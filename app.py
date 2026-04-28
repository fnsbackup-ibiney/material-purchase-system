# app.py
# 服裝物料採購單自動化處理系統
# Step 1:檔案上傳 + 合併儲存格(向下填充)+ 表格預覽

import pandas as pd
import streamlit as st


# ── 頁面設定 ─────────────────────────────────────────────
st.set_page_config(
    page_title="物料採購單處理系統",
    page_icon="📋",
    layout="wide",
)

st.title("📋 服裝物料採購單自動化處理系統")
st.caption("Step 1:上傳原始 Excel/CSV → 處理合併儲存格 → 預覽確認")


# ── 檔案上傳區 ────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "請選擇要處理的檔案(可一次多選)",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True,
    help="支援 .xlsx / .xls / .csv;Excel 多工作表會全部讀進來",
)

if not uploaded_files:
    st.info("👆 請從上方上傳一個或多個檔案")
    st.stop()


# ── 讀檔函式 ─────────────────────────────────────────────
def read_file(file) -> dict:
    """
    讀取單一上傳檔,回傳 {sheet_name: DataFrame}。
    重點:
      1. header=None  → 不預設第一行是標題,先把整張表原汁原味讀進來
      2. dtype=str    → 一律當字串讀,避免 pandas 把編號 "001" 變成 1
      3. ffill(axis=0)→ 向下填充,把因為「合併儲存格」造成的 NaN 補滿
    """
    name = file.name.lower()
    sheets: dict[str, pd.DataFrame] = {}

    if name.endswith(".csv"):
        df = pd.read_csv(file, header=None, dtype=str)
        df = df.ffill(axis=0)
        sheets["CSV"] = df
    else:
        # Excel:把所有工作表都讀進來
        xls = pd.ExcelFile(file, engine="openpyxl")
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(
                xls,
                sheet_name=sheet_name,
                header=None,
                dtype=str,
            )
            df = df.ffill(axis=0)
            sheets[sheet_name] = df

    return sheets


# ── 逐檔解析並展示 ────────────────────────────────────────
for file in uploaded_files:
    st.divider()
    st.subheader(f"📄 {file.name}")

    try:
        sheets = read_file(file)
    except Exception as e:
        st.error(f"❌ 讀取失敗:{e}")
        continue

    # 顯示是否還有殘留 NaN(理論上 ffill 後不該有,留作驗證)
    total_rows = sum(len(df) for df in sheets.values())
    total_nan = sum(int(df.isna().sum().sum()) for df in sheets.values())
    st.caption(
        f"工作表數:{len(sheets)}  /  總行數:{total_rows}  /  "
        f"殘留 NaN 格:{total_nan}(若 >0 表示某欄整欄都空白)"
    )

    # 單一 sheet 直接顯示;多 sheet 用 tab 分頁
    if len(sheets) == 1:
        sheet_name, df = next(iter(sheets.items()))
        st.markdown(f"**工作表:`{sheet_name}`** — {df.shape[0]} 行 × {df.shape[1]} 欄")
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        tabs = st.tabs([f"📑 {n}" for n in sheets.keys()])
        for tab, (sheet_name, df) in zip(tabs, sheets.items()):
            with tab:
                st.caption(f"{df.shape[0]} 行 × {df.shape[1]} 欄")
                st.dataframe(df, use_container_width=True, hide_index=True)
