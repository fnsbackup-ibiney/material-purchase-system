# app.py
# 服裝物料採購單自動化處理系統
# Step 1: 檔案上傳 + 合併儲存格(向下填充)+ 表格預覽
# Step 2: 雙層拆分(供應商 → 款號)+ 上傳記錄 sidebar + 清空按鈕

import pandas as pd
import streamlit as st


# ── 頁面設定 ─────────────────────────────────────────────
st.set_page_config(
    page_title="物料採購單處理系統",
    page_icon="📋",
    layout="wide",
)


# ── Session state 初始化 ─────────────────────────────────
# upload_history:累積這次 session 所有上傳過的檔名(關閉瀏覽器才清)
# uploader_key:清空按鈕按下時 +1,藉換 key 讓 file_uploader 重置
if "upload_history" not in st.session_state:
    st.session_state.upload_history = []
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0


# ── 左側 sidebar:本次上傳記錄 + 清空按鈕 ─────────────────
with st.sidebar:
    st.header("📜 本次上傳記錄")

    if st.session_state.upload_history:
        st.caption(f"共 {len(st.session_state.upload_history)} 筆")
        for i, fname in enumerate(st.session_state.upload_history, 1):
            st.markdown(f"`{i}.` {fname}")
    else:
        st.info("尚未上傳任何檔案")

    st.divider()

    if st.button("🔄 清空,重新開始", type="primary", use_container_width=True):
        st.session_state.upload_history = []
        st.session_state.uploader_key += 1
        st.rerun()


# ── 主畫面標題 ───────────────────────────────────────────
st.title("📋 服裝物料採購單自動化處理系統")
st.caption("Step 2:檔案上傳 → 合併儲存格 → 雙層拆分(供應商 → 款號)")


# ── 檔案上傳區 ────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "請選擇要處理的檔案(可一次多選)",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True,
    help="支援 .xlsx / .xls / .csv;Excel 多工作表會全部讀進來",
    key=f"file_uploader_{st.session_state.uploader_key}",
)

# 累積上傳記錄(去重 by 檔名)
if uploaded_files:
    for f in uploaded_files:
        if f.name not in st.session_state.upload_history:
            st.session_state.upload_history.append(f.name)

if not uploaded_files:
    st.info("👆 請從上方上傳一個或多個檔案")
    st.stop()


# ── 讀檔函式(同 Step 1)─────────────────────────────────
def read_file(file) -> dict:
    """讀取上傳檔,回傳 {sheet_name: DataFrame}(已 ffill)。"""
    name = file.name.lower()
    sheets: dict[str, pd.DataFrame] = {}

    if name.endswith(".csv"):
        df = pd.read_csv(file, header=None, dtype=str)
        df = df.ffill(axis=0)
        sheets["CSV"] = df
    else:
        xls = pd.ExcelFile(file, engine="openpyxl")
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=str)
            df = df.ffill(axis=0)
            sheets[sheet_name] = df

    return sheets


# ── Step 2 核心邏輯:抽資料區 + 雙層拆分 ─────────────────

# 從樣本分析得知:真表頭固定在第 10 行(0-indexed)
# 第 0~9 行是公司抬頭區、供應商/聯絡人/客戶資訊
HEADER_ROW = 10

# 抬頭區的「供應商」備援值位置(raw_sample_1 沒有「供應商」欄時用)
SUPPLIER_LABEL_ROW = 5
SUPPLIER_VALUE_COL = 2


def parse_sheet(df: pd.DataFrame, sheet_name: str) -> tuple[pd.DataFrame, dict]:
    """
    把一張 ffilled 過的 sheet 拆成「資料區 + 抬頭資訊」。

    回傳:
      data_df: 資料區(以第 10 行為欄名,從第 11 行起)
      header_info: {default_supplier, sheet_name}
    """
    if df.shape[0] <= HEADER_ROW:
        return pd.DataFrame(), {"default_supplier": None, "sheet_name": sheet_name}

    headers = df.iloc[HEADER_ROW].tolist()
    data_df = df.iloc[HEADER_ROW + 1:].copy()
    data_df.columns = headers
    data_df = data_df.reset_index(drop=True)

    # 去掉完全空白的行(資料區尾巴常見)
    data_df = data_df.dropna(how="all").reset_index(drop=True)

    # 從抬頭區抓供應商備援值(raw_sample_1 用,因為資料區沒有供應商欄)
    default_supplier = None
    if df.shape[0] > SUPPLIER_LABEL_ROW and df.shape[1] > SUPPLIER_VALUE_COL:
        cell = df.iat[SUPPLIER_LABEL_ROW, SUPPLIER_VALUE_COL]
        if cell and str(cell).strip() and str(cell).lower() != "nan":
            default_supplier = str(cell).strip()

    return data_df, {"default_supplier": default_supplier, "sheet_name": sheet_name}


def split_by_supplier_and_style(data_df: pd.DataFrame, header_info: dict) -> dict:
    """
    雙層拆分:先按供應商,再按款號。
    回傳:{ supplier: { style: sub_df } }
    """
    if data_df.empty:
        return {}

    # 找「供应商」欄(優先用資料區的;raw_sample_2 有這欄)
    supplier_col = None
    for col in data_df.columns:
        if isinstance(col, str) and "供应商" in col:
            supplier_col = col
            break

    # 找「款号」欄
    style_col = None
    for col in data_df.columns:
        if isinstance(col, str) and "款号" in col:
            style_col = col
            break

    if style_col is None:
        return {}

    # 第一層:按供應商分組
    if supplier_col:
        # 資料區有「供应商」欄
        groups_by_supplier = {}
        for s, g in data_df.groupby(supplier_col, dropna=False):
            if pd.isna(s) or not str(s).strip():
                continue
            groups_by_supplier[str(s).strip()] = g.copy()
    else:
        # 資料區沒有,退而用抬頭區的供應商;再退用 sheet 名
        default = (
            header_info.get("default_supplier")
            or header_info.get("sheet_name")
            or "未知供應商"
        )
        groups_by_supplier = {default: data_df.copy()}

    # 第二層:每個供應商裡按款號分組
    result = {}
    for supplier, supplier_df in groups_by_supplier.items():
        styles = {}
        for style, style_df in supplier_df.groupby(style_col, dropna=False):
            if pd.isna(style) or not str(style).strip():
                continue
            styles[str(style).strip()] = style_df.copy()
        if styles:
            result[supplier] = styles

    return result


# ── 逐檔解析並展示 ────────────────────────────────────────
for file in uploaded_files:
    st.divider()
    st.subheader(f"📄 {file.name}")

    try:
        sheets = read_file(file)
    except Exception as e:
        st.error(f"❌ 讀取失敗:{e}")
        continue

    # 總體統計
    total_rows = sum(len(df) for df in sheets.values())
    total_nan = sum(int(df.isna().sum().sum()) for df in sheets.values())
    st.caption(
        f"工作表數:{len(sheets)}  /  總行數:{total_rows}  /  "
        f"殘留 NaN 格:{total_nan}"
    )

    # 對每張 sheet 做「原始預覽 + 拆分結果」
    for sheet_name, df in sheets.items():
        with st.expander(
            f"📑 工作表:{sheet_name}  —  {df.shape[0]} 行 × {df.shape[1]} 欄",
            expanded=True,
        ):
            # ① 原始 ffilled 預覽(收合在 expander 內,預設展開)
            st.markdown("##### 1️⃣ 原始預覽(ffill 後)")
            st.dataframe(df, use_container_width=True, hide_index=True, height=250)

            # ② 雙層拆分結果
            st.markdown("##### 2️⃣ 雙層拆分(供應商 → 款號)")

            data_df, header_info = parse_sheet(df, sheet_name)

            if data_df.empty:
                st.warning(
                    f"⚠️ 此工作表資料區為空(可能總行數不足 {HEADER_ROW + 1} 行)"
                )
                continue

            split_result = split_by_supplier_and_style(data_df, header_info)

            if not split_result:
                st.warning(
                    "⚠️ 找不到「款号」欄,無法拆分。請確認第 10 行有「款号」標題"
                )
                continue

            n_suppliers = len(split_result)
            n_styles = sum(len(v) for v in split_result.values())
            st.caption(
                f"📊 拆出 **{n_suppliers}** 個供應商,共 **{n_styles}** 個款號"
            )

            # 第一層 tabs:供應商
            supplier_tabs = st.tabs([f"🏭 {s}" for s in split_result.keys()])
            for stab, (supplier, styles) in zip(supplier_tabs, split_result.items()):
                with stab:
                    # 第二層:每個款號一個 expander
                    for style, sdf in styles.items():
                        with st.expander(
                            f"🏷️ 款號:{style}  ({len(sdf)} 筆物料)",
                            expanded=False,
                        ):
                            st.dataframe(
                                sdf,
                                use_container_width=True,
                                hide_index=True,
                                height=200,
                            )
