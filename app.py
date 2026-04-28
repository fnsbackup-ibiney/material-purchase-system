# app.py
# ÊúçË£ùÁâ©ÊñôÊé°Ë≥ºÂñÆËá™ÂãïÂåñËôïÁêÜÁ≥ªÁµ±
# Step 1: Ê™îÊ°à‰∏äÂÇ≥ + Âêà‰ΩµÂÑ≤Â≠òÊ†º(Âêë‰∏ãÂ°´ÂÖÖ)+ Ë°®Ê†ºÈ†êË¶Ω
# Step 2: ÈõôÂ±§ÊãÜÂàÜ(‰æõÊáâÂïÜ ‚Üí Ê¨æËôü)+ ‰∏äÂÇ≥Ë®òÈåÑ sidebar + Ê∏ÖÁ©∫ÊåâÈàï

import pandas as pd
import streamlit as st


# ‚îÄ‚îÄ È†ÅÈù¢Ë®≠ÂÆö ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
st.set_page_config(
    page_title="Áâ©ÊñôÊé°Ë≥ºÂñÆËôïÁêÜÁ≥ªÁµ±",
    page_icon="üìã",
    layout="wide",
)


# ‚îÄ‚îÄ Session state ÂàùÂßãÂåñ ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
# upload_history:Á¥ØÁ©çÈÄôÊ¨° session ÊâÄÊúâ‰∏äÂÇ≥ÈÅéÁöÑÊ™îÂêç(ÈóúÈñâÁÄèË¶ΩÂô®ÊâçÊ∏Ö)
# uploader_key:Ê∏ÖÁ©∫ÊåâÈàïÊåâ‰∏ãÊôÇ +1,ËóâÊèõ key ËÆì file_uploader ÈáçÁΩÆ
if "upload_history" not in st.session_state:
    st.session_state.upload_history = []
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0


# ‚îÄ‚îÄ Â∑¶ÂÅ¥ sidebar:Êú¨Ê¨°‰∏äÂÇ≥Ë®òÈåÑ + Ê∏ÖÁ©∫ÊåâÈàï ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
with st.sidebar:
    st.header("üìú Êú¨Ê¨°‰∏äÂÇ≥Ë®òÈåÑ")

    if st.session_state.upload_history:
        st.caption(f"ÂÖ± {len(st.session_state.upload_history)} Á≠Ü")
        for i, fname in enumerate(st.session_state.upload_history, 1):
            st.markdown(f"`{i}.` {fname}")
    else:
        st.info("Â∞öÊú™‰∏äÂÇ≥‰ªª‰ΩïÊ™îÊ°à")

    st.divider()

    if st.button("üîÑ Ê∏ÖÁ©∫,ÈáçÊñ∞ÈñãÂßã", type="primary", use_container_width=True):
        st.session_state.upload_history = []
        st.session_state.uploader_key += 1
        st.rerun()


# ‚îÄ‚îÄ ‰∏ªÁï´Èù¢Ê®ôÈ°å ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
st.title("üìã ÊúçË£ùÁâ©ÊñôÊé°Ë≥ºÂñÆËá™ÂãïÂåñËôïÁêÜÁ≥ªÁµ±")
st.caption("Step 2:Ê™îÊ°à‰∏äÂÇ≥ ‚Üí Âêà‰ΩµÂÑ≤Â≠òÊ†º ‚Üí ÈõôÂ±§ÊãÜÂàÜ(‰æõÊáâÂïÜ ‚Üí Ê¨æËôü)")


# ‚îÄ‚îÄ Ê™îÊ°à‰∏äÂÇ≥ÂçÄ ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
uploaded_files = st.file_uploader(
    "Ë´ãÈÅ∏ÊìáË¶ÅËôïÁêÜÁöÑÊ™îÊ°à(ÂèØ‰∏ÄÊ¨°Â§öÈÅ∏)",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True,
    help="ÊîØÊè¥ .xlsx / .xls / .csv;Excel Â§öÂ∑•‰ΩúË°®ÊúÉÂÖ®ÈÉ®ËÆÄÈÄ≤‰æÜ",
    key=f"file_uploader_{st.session_state.uploader_key}",
)

# Á¥ØÁ©ç‰∏äÂÇ≥Ë®òÈåÑ(ÂéªÈáç by Ê™îÂêç)
if uploaded_files:
    for f in uploaded_files:
        if f.name not in st.session_state.upload_history:
            st.session_state.upload_history.append(f.name)

if not uploaded_files:
    st.info("üëÜ Ë´ãÂæû‰∏äÊñπ‰∏äÂÇ≥‰∏ÄÂÄãÊàñÂ§öÂÄãÊ™îÊ°à")
    st.stop()


# ‚îÄ‚îÄ ËÆÄÊ™îÂáΩÂºè(Âêå Step 1)‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
def read_file(file) -> dict:
    """ËÆÄÂèñ‰∏äÂÇ≥Ê™î,ÂõûÂÇ≥ {sheet_name: DataFrame}(Â∑≤ ffill)„ÄÇ"""
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


# ‚îÄ‚îÄ Step 2 Ê†∏ÂøÉÈÇèËºØ:ÊäΩË≥áÊñôÂçÄ + ÈõôÂ±§ÊãÜÂàÜ ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ

# ÂæûÊ®£Êú¨ÂàÜÊûêÂæóÁü•:ÁúüË°®È†≠Âõ∫ÂÆöÂú®Á¨¨ 10 Ë°å(0-indexed)
# Á¨¨ 0~9 Ë°åÊòØÂÖ¨Âè∏Êä¨È†≠ÂçÄ„ÄÅ‰æõÊáâÂïÜ/ËÅØÁµ°‰∫∫/ÂÆ¢Êà∂Ë≥áË®ä
HEADER_ROW = 10

# Êä¨È†≠ÂçÄÁöÑ„Äå‰æõÊáâÂïÜ„ÄçÂÇôÊè¥ÂÄº‰ΩçÁΩÆ(raw_sample_1 Ê≤íÊúâ„Äå‰æõÊáâÂïÜ„ÄçÊ¨ÑÊôÇÁî®)
SUPPLIER_LABEL_ROW = 5
SUPPLIER_VALUE_COL = 2


def parse_sheet(df: pd.DataFrame, sheet_name: str) -> tuple[pd.DataFrame, dict]:
    """
    Êää‰∏ÄÂºµ ffilled ÈÅéÁöÑ sheet ÊãÜÊàê„ÄåË≥áÊñôÂçÄ + Êä¨È†≠Ë≥áË®ä„Äç„ÄÇ

    ÂõûÂÇ≥:
      data_df: Ë≥áÊñôÂçÄ(‰ª•Á¨¨ 10 Ë°åÁÇ∫Ê¨ÑÂêç,ÂæûÁ¨¨ 11 Ë°åËµ∑)
      header_info: {default_supplier, sheet_name}
    """
    if df.shape[0] <= HEADER_ROW:
        return pd.DataFrame(), {"default_supplier": None, "sheet_name": sheet_name}

    headers = df.iloc[HEADER_ROW].tolist()
    data_df = df.iloc[HEADER_ROW + 1:].copy()
    data_df.columns = headers
    data_df = data_df.reset_index(drop=True)

    # ÂéªÊéâÂÆåÂÖ®Á©∫ÁôΩÁöÑË°å(Ë≥áÊñôÂçÄÂ∞æÂ∑¥Â∏∏Ë¶ã)
    data_df = data_df.dropna(how="all").reset_index(drop=True)

    # ÂæûÊä¨È†≠ÂçÄÊäì‰æõÊáâÂïÜÂÇôÊè¥ÂÄº(raw_sample_1 Áî®,Âõ†ÁÇ∫Ë≥áÊñôÂçÄÊ≤íÊúâ‰æõÊáâÂïÜÊ¨Ñ)
    default_supplier = None
    if df.shape[0] > SUPPLIER_LABEL_ROW and df.shape[1] > SUPPLIER_VALUE_COL:
        cell = df.iat[SUPPLIER_LABEL_ROW, SUPPLIER_VALUE_COL]
        if cell and str(cell).strip() and str(cell).lower() != "nan":
            default_supplier = str(cell).strip()

    return data_df, {"default_supplier": default_supplier, "sheet_name": sheet_name}


def split_by_supplier_and_style(data_df: pd.DataFrame, header_info: dict) -> dict:
    """
    ÈõôÂ±§ÊãÜÂàÜ:ÂÖàÊåâ‰æõÊáâÂïÜ,ÂÜçÊåâÊ¨æËôü„ÄÇ
    ÂõûÂÇ≥:{ supplier: { style: sub_df } }
    """
    if data_df.empty:
        return {}

    # Êâæ„Äå‰æõÂ∫îÂïÜ„ÄçÊ¨Ñ(ÂÑ™ÂÖàÁî®Ë≥áÊñôÂçÄÁöÑ;raw_sample_2 ÊúâÈÄôÊ¨Ñ)
    supplier_col = None
    for col in data_df.columns:
        if isinstance(col, str) and "‰æõÂ∫îÂïÜ" in col:
            supplier_col = col
            break

    # Êâæ„ÄåÊ¨æÂè∑„ÄçÊ¨Ñ
    style_col = None
    for col in data_df.columns:
        if isinstance(col, str) and "Ê¨æÂè∑" in col:
            style_col = col
            break

    if style_col is None:
        return {}

    # Á¨¨‰∏ÄÂ±§:Êåâ‰æõÊáâÂïÜÂàÜÁµÑ
    if supplier_col:
        # Ë≥áÊñôÂçÄÊúâ„Äå‰æõÂ∫îÂïÜ„ÄçÊ¨Ñ
        groups_by_supplier = {}
        for s, g in data_df.groupby(supplier_col, dropna=False):
            if pd.isna(s) or not str(s).strip():
                continue
            groups_by_supplier[str(s).strip()] = g.copy()
    else:
        # Ë≥áÊñôÂçÄÊ≤íÊúâ,ÈÄÄËÄåÁî®Êä¨È†≠ÂçÄÁöÑ‰æõÊáâÂïÜ;ÂÜçÈÄÄÁî® sheet Âêç
        default = (
            header_info.get("default_supplier")
            or header_info.get("sheet_name")
            or "Êú™Áü•‰æõÊáâÂïÜ"
        )
        groups_by_supplier = {default: data_df.copy()}

    # Á¨¨‰∫åÂ±§:ÊØèÂÄã‰æõÊáâÂïÜË£°ÊåâÊ¨æËôüÂàÜÁµÑ
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


# ‚îÄ‚îÄ ÈÄêÊ™îËß£Êûê‰∏¶Â±ïÁ§∫ ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
for file in uploaded_files:
    st.divider()
    st.subheader(f"üìÑ {file.name}")

    try:
        sheets = read_file(file)
    except Exception as e:
        st.error(f"‚ùå ËÆÄÂèñÂ§±Êïó:{e}")
        continue

    # Á∏ΩÈ´îÁµ±Ë®à
    total_rows = sum(len(df) for df in sheets.values())
    total_nan = sum(int(df.isna().sum().sum()) for df in sheets.values())
    st.caption(
        f"Â∑•‰ΩúË°®Êï∏:{len(sheets)}  /  Á∏ΩË°åÊï∏:{total_rows}  /  "
        f"ÊÆòÁïô NaN Ê†º:{total_nan}"
    )

    # Â∞çÊØèÂºµ sheet ÂÅö„ÄåÂéüÂßãÈ†êË¶Ω + ÊãÜÂàÜÁµêÊûú„Äç
    for sheet_name, df in sheets.items():
        with st.expander(
            f"üìë Â∑•‰ΩúË°®:{sheet_name}  ‚Äî  {df.shape[0]} Ë°å √ó {df.shape[1]} Ê¨Ñ",
            expanded=True,
        ):
            # ‚ë† ÂéüÂßã ffilled È†êË¶Ω(Êî∂ÂêàÂú® expander ÂÖß,È†êË®≠Â±ïÈñã)
            st.markdown("##### 1Ô∏è‚É£ ÂéüÂßãÈ†êË¶Ω(ffill Âæå)")
            st.dataframe(df, use_container_width=True, hide_index=True, height=250)

            # ‚ë° ÈõôÂ±§ÊãÜÂàÜÁµêÊûú
            st.markdown("##### 2Ô∏è‚É£ ÈõôÂ±§ÊãÜÂàÜ(‰æõÊáâÂïÜ ‚Üí Ê¨æËôü)")

            data_df, header_info = parse_sheet(df, sheet_name)

            if data_df.empty:
                st.warning(
                    f"‚ö†Ô∏è Ê≠§Â∑•‰ΩúË°®Ë≥áÊñôÂçÄÁÇ∫Á©∫(ÂèØËÉΩÁ∏ΩË°åÊï∏‰∏çË∂≥ {HEADER_ROW + 1} Ë°å)"
                )
                continue

            split_result = split_by_supplier_and_style(data_df, header_info)

            if not split_result:
                st.warning(
                    "‚ö†Ô∏è Êâæ‰∏çÂà∞„ÄåÊ¨æÂè∑„ÄçÊ¨Ñ,ÁÑ°Ê≥ïÊãÜÂàÜ„ÄÇË´ãÁ¢∫Ë™çÁ¨¨ 10 Ë°åÊúâ„ÄåÊ¨æÂè∑„ÄçÊ®ôÈ°å"
                )
                continue

            n_suppliers = len(split_result)
            n_styles = sum(len(v) for v in split_result.values())
            st.caption(
                f"üìä ÊãÜÂá∫ **{n_suppliers}** ÂÄã‰æõÊáâÂïÜ,ÂÖ± **{n_styles}** ÂÄãÊ¨æËôü"
            )

            # Á¨¨‰∏ÄÂ±§ tabs:‰æõÊáâÂïÜ
            supplier_tabs = st.tabs([f"üè≠ {s}" for s in split_result.keys()])
            for stab, (supplier, styles) in zip(supplier_tabs, split_result.items()):
                with stab:
                    # Á¨¨‰∫åÂ±§:ÊØèÂÄãÊ¨æËôü‰∏ÄÂÄã expander
                    for style, sdf in styles.items():
                        with st.expander(
                            f"üè∑Ô∏è Ê¨æËôü:{style}  ({len(sdf)} Á≠ÜÁâ©Êñô)",
                            expanded=False,
                        ):
                            st.dataframe(
                                sdf,
                                use_container_width=True,
                                hide_index=True,
                                height=200,
                            )
