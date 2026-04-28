# app.py
# 服裝物料採購單自動化處理系統
# Step 1: 檔案上傳 + 合併儲存格(向下填充)+ 表格預覽
# Step 2: 雙層拆分(供應商 → 款號)+ 上傳記錄 sidebar + 清空按鈕
# Step 3: 真款號 vs 備註分類 + 智慧取價(優先序:成本价 > 大货价 > 单价 > 报价)

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
# uploaded_blobs:對應的檔案二進位內容,讓 sidebar 可以提供下載
# uploader_key:清空按鈕按下時 +1,藉換 key 讓 file_uploader 重置
if "upload_history" not in st.session_state:
    st.session_state.upload_history = []
if "uploaded_blobs" not in st.session_state:
    st.session_state.uploaded_blobs = {}
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0


# ── 左側 sidebar:本次上傳記錄 + 清空按鈕 ─────────────────
with st.sidebar:
    st.header("📜 本次上傳記錄")

    if st.session_state.upload_history:
        st.caption(f"共 {len(st.session_state.upload_history)} 筆(點擊下載原始檔)")
        for i, fname in enumerate(st.session_state.upload_history, 1):
            blob = st.session_state.uploaded_blobs.get(fname)
            if blob is not None:
                # 有暫存的二進位 → 提供下載按鈕
                st.download_button(
                    label=f"📥 {i}. {fname}",
                    data=blob,
                    file_name=fname,
                    mime="application/octet-stream",
                    key=f"dl_{i}_{fname}",
                    use_container_width=True,
                )
            else:
                # 沒暫存(理論上不會發生,留作保險)
                st.markdown(f"`{i}.` {fname} _(內容已清)_")
    else:
        st.info("尚未上傳任何檔案")

    st.divider()

    # 「清空」只清掉目前選著的檔案,讓你可以上傳下一批
    # 上傳記錄(歷史檔名)會繼續累積,直到你關閉/重新整理瀏覽器
    if st.button("🔄 清空,重新開始", type="primary", use_container_width=True):
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

# 累積上傳記錄 + 暫存原始檔位元組(讓 sidebar 可下載)
# 若有新檔名加入,立刻 rerun 一次,讓 sidebar 看到最新記錄
if uploaded_files:
    new_names = []
    for f in uploaded_files:
        if f.name not in st.session_state.upload_history:
            new_names.append(f.name)
            # getvalue() 取出位元組,存進 session(只活在當前瀏覽器分頁)
            st.session_state.uploaded_blobs[f.name] = f.getvalue()
    if new_names:
        st.session_state.upload_history.extend(new_names)
        st.rerun()

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

    # 取第 10 行當欄名,並處理「空白欄名」「重複欄名」問題
    # 否則 Streamlit 顯示表格會 ValueError(duplicate columns / NaN columns)
    raw_headers = df.iloc[HEADER_ROW].tolist()
    headers, seen = [], {}
    for i, h in enumerate(raw_headers):
        # NaN / 空白 → 用 _col_X 占位
        if h is None or pd.isna(h) or not str(h).strip() or str(h).lower() == "nan":
            name = f"_col_{i}"
        else:
            name = str(h).strip()
        # 重複名 → 加序號
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        headers.append(name)

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


# ── Step 3 核心:真款號判別 + 備註分類 + 智慧取價 ─────────

# 黑名單(明顯不是款號的字串)
_FAKE_CODE_BLACKLIST = {"TOTAL", "注:", "注:", "備註:", "备注:", "—", "-"}


def is_real_style_code(value) -> bool:
    """
    判斷一個 cell 是否為「真款號」。
    真款號:短(< 20 字)、含數字、不在黑名單、不是「1./2./3.」這種編號條款。
    """
    if value is None or pd.isna(value):
        return False
    s = str(value).strip()
    if not s:
        return False
    if s.upper() in {x.upper() for x in _FAKE_CODE_BLACKLIST}:
        return False
    # 長度太長 → 多半是中文備註句
    if len(s) > 20:
        return False
    # 沒有任何數字 → 多半是標籤,不是編號
    if not any(c.isdigit() for c in s):
        return False
    # 開頭是「1.」「2、」這種編號條款
    if len(s) >= 2 and s[0].isdigit() and s[1] in ".．、 ":
        return False
    return True


def classify_data_rows(
    data_df: pd.DataFrame, style_col: str
) -> tuple[pd.DataFrame, list[str]]:
    """
    分離資料行為「真款號物料」與「備註行」。
    真款號 → real_df
    其他 → 取 style_col 的字串當作備註,集中為一個 list(去重保序)
    """
    if data_df.empty or style_col not in data_df.columns:
        return data_df.copy(), []

    real_mask = data_df[style_col].apply(is_real_style_code)
    real_df = data_df[real_mask].copy().reset_index(drop=True)
    notes_df = data_df[~real_mask]

    notes: list[str] = []
    seen: set[str] = set()
    for _, row in notes_df.iterrows():
        text = row[style_col]
        if text is None or pd.isna(text):
            continue
        text = str(text).strip()
        # 過濾無意義小標題、合計列
        if not text or text.upper() == "TOTAL" or text in {"注:", "注:"}:
            continue
        if text not in seen:
            seen.add(text)
            notes.append(text)

    return real_df, notes


# 智慧取價:依優先序找出「成本價」欄
# 順序:成本价 > 供应商成本价 > 大货价 > 单价 > 报价
_PRICE_PRIORITY = ["成本价", "供应商成本价", "大货价", "单价", "报价"]


def find_cost_column(columns) -> str | None:
    """從欄位名稱中找出最優先的價格欄;找不到回 None"""
    for keyword in _PRICE_PRIORITY:
        for col in columns:
            if isinstance(col, str) and keyword in col:
                return col
    return None


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

            # ② 雙層拆分結果(Step 3 強化:識別假款號為備註 + 智慧取價)
            st.markdown("##### 2️⃣ 雙層拆分(供應商 → 款號)")

            data_df, header_info = parse_sheet(df, sheet_name)

            if data_df.empty:
                st.warning(
                    f"⚠️ 此工作表資料區為空(可能總行數不足 {HEADER_ROW + 1} 行)"
                )
                continue

            # 找款號欄(在 split 之前先找出來,以便分類)
            style_col_for_class = next(
                (c for c in data_df.columns if isinstance(c, str) and "款号" in c),
                None,
            )
            if style_col_for_class is None:
                st.warning("⚠️ 找不到「款号」欄,無法處理")
                continue

            # ── Step 3:分離真款號物料 vs 備註 ──
            real_df, sheet_notes = classify_data_rows(data_df, style_col_for_class)

            # ── Step 3:智慧取價 ──
            cost_col = find_cost_column(real_df.columns)

            # 顯示 sheet 級的備註(出口規定、條款等)
            if sheet_notes:
                with st.container(border=True):
                    st.markdown(f"📌 **此 sheet 共 {len(sheet_notes)} 條備註**(會套用到此 sheet 所有款號)")
                    for note in sheet_notes:
                        st.markdown(f"- {note}")

            # 顯示成本價欄判斷結果
            if cost_col:
                st.success(f"💰 自動抓到成本價欄:**{cost_col}**(優先序:成本价 > 大货价 > 单价 > 报价)")
            else:
                st.info("💰 未找到成本價相關欄位")

            # 用真款號 df 做拆分
            split_result = split_by_supplier_and_style(real_df, header_info)

            if not split_result:
                st.warning("⚠️ 過濾掉假款號後,沒有任何真款號可拆分")
                continue

            n_suppliers = len(split_result)
            n_styles = sum(len(v) for v in split_result.values())
            st.caption(
                f"📊 拆出 **{n_suppliers}** 個供應商,共 **{n_styles}** 個真款號"
                f"(已過濾 {len(sheet_notes)} 條備註)"
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
                            # 顯示這筆款號套用了哪些備註(目前是 sheet 級)
                            if sheet_notes:
                                st.caption(f"📌 套用備註({len(sheet_notes)} 條)")
                            # 顯示成本價欄(若有)
                            if cost_col and cost_col in sdf.columns:
                                cost_values = sdf[cost_col].dropna().unique().tolist()
                                if cost_values:
                                    st.caption(
                                        f"💰 此款成本價(欄「{cost_col}」):"
                                        f"{', '.join(map(str, cost_values[:5]))}"
                                    )
                            st.dataframe(
                                sdf,
                                use_container_width=True,
                                hide_index=True,
                                height=200,
                            )
