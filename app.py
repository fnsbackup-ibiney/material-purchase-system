# app.py
# 服裝物料採購單自動化處理系統
# Step 1: 檔案上傳 + 合併儲存格(向下填充)+ 表格預覽
# Step 2: 雙層拆分(供應商 → 款號)+ 上傳記錄 sidebar + 清空按鈕
# Step 3: 真款號 vs 備註分類 + 智慧取價(優先序:成本价 > 大货价 > 单价 > 报价)
# Step 4A: 標準材料庫上傳 + 三階紅燈比對(🟢已驗證 / 🟡規格不同 / 🚨全新料)

import io
import re
import zipfile

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
# Step 4A:標準庫資料(只存解析後的清單,不存原始 bytes)
if "std_library" not in st.session_state:
    st.session_state.std_library = None  # 解析後的 dict
if "std_library_filename" not in st.session_state:
    st.session_state.std_library_filename = None


# ── 左側 sidebar 渲染函式(在主流程結尾處呼叫,避免函式未定義錯誤)───
def render_sidebar():
    with st.sidebar:
        st.header("📜 本次上傳記錄")

        if st.session_state.upload_history:
            st.caption(f"共 {len(st.session_state.upload_history)} 筆(點擊下載原始檔)")
            for i, fname in enumerate(st.session_state.upload_history, 1):
                blob = st.session_state.uploaded_blobs.get(fname)
                if blob is not None:
                    st.download_button(
                        label=f"📥 {i}. {fname}",
                        data=blob,
                        file_name=fname,
                        mime="application/octet-stream",
                        key=f"dl_{i}_{fname}",
                        use_container_width=True,
                    )
                else:
                    st.markdown(f"`{i}.` {fname} _(內容已清)_")
        else:
            st.info("尚未上傳任何檔案")

        st.divider()

        # ── Step 4A:標準材料庫上傳區 ────────────────
        st.subheader("📚 標準材料庫")
        if st.session_state.std_library:
            st.success(f"✅ 已載入:{st.session_state.std_library_filename}")
            n_names = len(st.session_state.std_library["name_set"])
            st.caption(f"共 {n_names} 筆標準品名")
            if st.button("🗑️ 移除標準庫", use_container_width=True):
                st.session_state.std_library = None
                st.session_state.std_library_filename = None
                st.rerun()
        else:
            std_file = st.file_uploader(
                "上傳標準材料庫(供比對用)",
                type=["xlsx"],
                accept_multiple_files=False,
                help="不上傳也能用,但無法做紅燈警告",
                key="std_lib_uploader",
            )
            if std_file is not None:
                try:
                    with st.spinner("解析標準庫中(12MB 約需 30 秒)..."):
                        st.session_state.std_library = parse_standard_library(std_file.getvalue())
                        st.session_state.std_library_filename = std_file.name
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 無法解析:{e}")

        st.divider()

        # 「清空」只清當前 file_uploader,記錄與標準庫都保留
        if st.button("🔄 清空,重新開始", type="primary", use_container_width=True):
            st.session_state.uploader_key += 1
            st.rerun()


# 注意:render_sidebar() 故意延後到主流程「最尾端」呼叫,
# 因為 sidebar 內會用到 parse_standard_library() 等下方定義的函式。
# Python 從上往下讀,函式呼叫前必須先看過函式定義。


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
    render_sidebar()  # 主流程提早結束前先把 sidebar 畫出來
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


# ── Step 4A:標準材料庫讀取 + 三階紅燈比對 ─────────────────

# 標準庫每張 sheet 的設定:表頭行(0-indexed)、品名/料號欄、規格欄、跳過旗標
# 可比對的欄位 → key 值用標準化字串
_STD_SHEET_CONFIG = {
    "3F-Brax Zippers": {"header_row": 1, "name_keys": ["品名", "款号"], "size_keys": ["规格", "尺码"]},
    "Metal Trims":     {"header_row": 0, "name_keys": ["Article No.品名", "品名"], "size_keys": ["Size尺寸", "Size"]},
    "Rivet":           {"header_row": 1, "name_keys": ["Item", "FS article number"], "size_keys": []},
    "Non Metal Trims": {"header_row": 0, "name_keys": ["Article No.品名", "品名"], "size_keys": ["Size尺寸", "Size"]},
    "Label":           {"header_row": 0, "name_keys": ["Article No.", "品名"], "size_keys": ["Size", "尺寸"]},
    "Tape":            {"header_row": 0, "name_keys": ["Article No.品名", "品名"], "size_keys": ["Size尺寸", "Size"]},
    "Paper Tags":      {"skip": True},  # 設計圖,無結構化資料
}


def _strip_styles_xlsx(file_bytes: bytes) -> bytes:
    """
    繞過 openpyxl 3.1.5 對某些 .xlsx 的 styles.xml 解析 crash 問題。
    把 styles.xml 換成最小骨架後重新打包。資料完整保留,只丟掉樣式。
    """
    src = io.BytesIO(file_bytes)
    dst = io.BytesIO()
    minimal_styles = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        b'<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        b'<borders count="1"><border/></borders>'
        b'<cellStyleXfs count="1"><xf/></cellStyleXfs>'
        b'<cellXfs count="1"><xf/></cellXfs>'
        b'</styleSheet>'
    )
    with zipfile.ZipFile(src, "r") as zin:
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == "xl/styles.xml":
                    data = minimal_styles
                zout.writestr(item, data)
    return dst.getvalue()


def normalize_for_compare(s) -> str:
    """字串正規化(供比對用):去全部空白、轉大寫、去除常見單位寫法差異。"""
    if s is None or pd.isna(s):
        return ""
    text = str(s).upper()
    # 去掉所有空白(含全形空白)
    text = re.sub(r"\s+", "", text)
    # 統一全形→半形數字/標點
    text = text.replace(",", ",").replace(":", ":")
    return text


def parse_standard_library(file_bytes: bytes) -> dict:
    """
    解析使用者上傳的標準庫 .xlsx,回傳:
    {
        'name_set': set[str],         # 所有「品名」正規化字串集合
        'name_size_set': set[tuple],  # (品名正規化, 規格正規化) 二元組集合
        'sheets_info': dict           # 每張 sheet 的統計
    }
    """
    safe_bytes = _strip_styles_xlsx(file_bytes)
    xls = pd.ExcelFile(io.BytesIO(safe_bytes), engine="openpyxl")

    name_set: set[str] = set()
    name_size_set: set[tuple] = set()
    sheets_info: dict = {}

    for sheet_name in xls.sheet_names:
        cfg = _STD_SHEET_CONFIG.get(sheet_name)
        if cfg is None:
            sheets_info[sheet_name] = {"status": "未知 sheet,已跳過", "count": 0}
            continue
        if cfg.get("skip"):
            sheets_info[sheet_name] = {"status": "跳過(無結構化資料)", "count": 0}
            continue

        df = pd.read_excel(xls, sheet_name=sheet_name, header=cfg["header_row"], dtype=str)

        # 找品名欄(用 name_keys 列表逐一嘗試,取第一個符合的)
        name_col = None
        for key in cfg["name_keys"]:
            for col in df.columns:
                if isinstance(col, str) and key in col:
                    name_col = col
                    break
            if name_col:
                break

        # 找規格欄
        size_col = None
        for key in cfg.get("size_keys", []):
            for col in df.columns:
                if isinstance(col, str) and key in col:
                    size_col = col
                    break
            if size_col:
                break

        if name_col is None:
            sheets_info[sheet_name] = {"status": "找不到品名欄,跳過", "count": 0}
            continue

        count = 0
        for _, row in df.iterrows():
            raw_name = row.get(name_col)
            n = normalize_for_compare(raw_name)
            if not n:
                continue
            name_set.add(n)
            if size_col:
                s = normalize_for_compare(row.get(size_col))
                if s:
                    name_size_set.add((n, s))
                else:
                    name_size_set.add((n, ""))
            else:
                name_size_set.add((n, ""))
            count += 1

        sheets_info[sheet_name] = {
            "status": f"✅ 已讀取(品名欄:{name_col} / 規格欄:{size_col or '無'})",
            "count": count,
        }

    return {
        "name_set": name_set,
        "name_size_set": name_size_set,
        "sheets_info": sheets_info,
    }


def _match_one(query: str, library: dict, size_query: str) -> str:
    """
    單一 key 的三階比對。回傳 🟢 / 🟡 / 🚨。
    寬鬆比對:正規化後雙向「包含」即視為命中。
    """
    if not query:
        return "🚨"
    matched_names = set()
    for n_lib in library["name_set"]:
        if not n_lib:
            continue
        if query in n_lib or n_lib in query:
            matched_names.add(n_lib)

    if not matched_names:
        return "🚨"

    for n_lib in matched_names:
        sizes = {s for (nm, s) in library["name_size_set"] if nm == n_lib}
        if "" in sizes:
            return "🟢"
        if not size_query:
            return "🟡"
        for s_lib in sizes:
            if size_query in s_lib or s_lib in size_query:
                return "🟢"
    return "🟡"


# 三階優先序:🟢 最好 → 🟡 → 🚨 最差
_RANK = {"🟢": 0, "🟡": 1, "🚨": 2, "⚪": 3}


def compare_with_library(name, code, size, library: dict) -> str:
    """
    雙 key 比對:品名 OR 客戶編號 任一命中即算找到。
    取兩個 key 比對結果中「最好」的那一個。

    🟢 任一命中 + 規格相符
    🟡 任一命中但規格不同
    🚨 兩個都找不到
    ⚪ 沒上傳標準庫 / 兩個都是空字串
    """
    if not library:
        return "⚪"

    n_query = normalize_for_compare(name)
    c_query = normalize_for_compare(code)
    s_query = normalize_for_compare(size)

    if not n_query and not c_query:
        return "⚪"

    results = []
    if n_query:
        results.append(_match_one(n_query, library, s_query))
    if c_query:
        results.append(_match_one(c_query, library, s_query))

    # 取最好的結果(rank 最小)
    best = min(results, key=lambda r: _RANK.get(r, 9))
    return best


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
                        # ── Step 4A:逐筆物料做紅燈比對 ──
                        # 找此筆款的「品名」「客戶編號」「規格」欄
                        name_col_for_compare = next(
                            (c for c in sdf.columns if isinstance(c, str) and "品名" in c),
                            None,
                        )
                        code_col_for_compare = next(
                            (c for c in sdf.columns
                             if isinstance(c, str) and ("客户编号" in c or "客戶編號" in c or "Article" in c)),
                            None,
                        )
                        size_col_for_compare = next(
                            (c for c in sdf.columns
                             if isinstance(c, str) and ("规格" in c or "尺码" in c or "尺寸" in c)),
                            None,
                        )
                        # 算出比對狀態 — 雙 key 比對:品名 OR 客戶編號任一命中即綠
                        statuses = []
                        if name_col_for_compare or code_col_for_compare:
                            for _, mat_row in sdf.iterrows():
                                name = mat_row.get(name_col_for_compare) if name_col_for_compare else None
                                code = mat_row.get(code_col_for_compare) if code_col_for_compare else None
                                size = mat_row.get(size_col_for_compare) if size_col_for_compare else None
                                statuses.append(
                                    compare_with_library(name, code, size, st.session_state.std_library)
                                )
                        else:
                            statuses = ["⚪"] * len(sdf)

                        # 統計
                        n_green = statuses.count("🟢")
                        n_yellow = statuses.count("🟡")
                        n_red = statuses.count("🚨")
                        n_white = statuses.count("⚪")
                        # 款號標題若有紅燈,加警告 emoji
                        title_warn = " 🚨" if n_red > 0 else (" 🟡" if n_yellow > 0 else "")

                        with st.expander(
                            f"🏷️ 款號:{style}  ({len(sdf)} 筆物料){title_warn}",
                            expanded=(n_red > 0),  # 有紅燈自動展開
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
                            # 比對統計
                            if st.session_state.std_library:
                                st.caption(
                                    f"🚦 比對結果: 🟢 {n_green} 筆 / 🟡 {n_yellow} 筆 / "
                                    f"🚨 {n_red} 筆 / ⚪ {n_white} 筆未比對"
                                )
                            else:
                                st.caption("⚪ 尚未上傳標準庫,無法比對(👈 在左側 sidebar 上傳)")

                            # 加上「比對狀態」欄到表格(放第一欄)
                            sdf_display = sdf.copy()
                            sdf_display.insert(0, "🚦", statuses)

                            st.dataframe(
                                sdf_display,
                                use_container_width=True,
                                hide_index=True,
                                height=220,
                            )


# ── 主流程結束,最後渲染 sidebar(此時所有函式都已定義完成)──────
render_sidebar()
