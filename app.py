"""
PUMA ZeCom Column Mapper (v2 — robust reader)
===============================================
Upload a MARKETPLACE file (Lazada / Shopee / Zalora / TikTok) that only has EAN
(Seller SKU) at the row level, and get back the SAME file with ZeCom pricing
(and any other ZeCom columns you choose) mapped in.

JOIN CHAIN (ZeCom has no EAN, so we bridge through Content):
  Marketplace EAN  --(Content file)-->  Color No / Article No  --(ZeCom file)-->  Selected columns

v2 changes (built after real ZeCom tracker files were tested):
  - Multi-engine, format-sniffing file reader (openpyxl / calamine / xlrd / HTML-in-disguise),
    with graceful fallback and a clear "what was tried" error message instead of one confusing
    engine-import error.
  - Sheet picker (ZeCom trackers frequently bundle MY + SG in one file, PH in another).
  - Header row is auto-detected AND manually overridable, with a raw-row preview, because
    tracker layouts change often and auto-detection can guess wrong.
  - Repeated column names (e.g. "MY RRP" appearing 6x for 6 different campaign tiers) are
    now handled: each gets a unique internal name plus a human label that includes the
    campaign/tier banner text sitting above it and its real Excel column letter, so nothing
    silently collides or gets overwritten.
  - Wider key-matching hints (Style#, PIM Article#, PIM_Article#, Color No, etc.) since this
    varies by region/file.
  - Optional key normalization (trim/uppercase/strip stray .0) to catch near-miss key mismatches.
"""

import io
import re

import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="ZeCom Column Mapper", layout="wide")

# ---------------------------------------------------------------------------
# Low-level file reading — format sniffing + multi-engine fallback
# ---------------------------------------------------------------------------

def _looks_like_html(b: bytes) -> bool:
    head = b[:512].lstrip().lower()
    return head.startswith(b"<html") or head.startswith(b"<!doctype") or b"<table" in head[:2000].lower()


def _sniff_kind(b: bytes) -> str:
    if b[:4] == b"PK\x03\x04":
        return "zip"       # genuine .xlsx/.xlsm (OOXML)
    if b[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "ole"        # legacy .xls
    if _looks_like_html(b):
        return "html"        # many marketplace "xlsx" exports are actually HTML tables
    return "unknown"


@st.cache_data(show_spinner=False)
def list_sheets(file_bytes, filename):
    """
    Try each engine/format in turn just to enumerate sheet names.
    Returns (engine_used, sheet_names, attempt_log).
    engine_used is one of: 'openpyxl', 'calamine', 'xlrd', 'html', 'csv'
    """
    attempts = []
    kind = _sniff_kind(file_bytes)

    if filename.lower().endswith(".csv"):
        return "csv", ["(csv)"], ["Detected .csv extension"]

    engine_order = ["openpyxl", "calamine", "xlrd"]
    for engine in engine_order:
        try:
            xls = pd.ExcelFile(io.BytesIO(file_bytes), engine=engine)
            return engine, xls.sheet_names, attempts + [f"{engine}: OK"]
        except ImportError:
            attempts.append(f"{engine}: not installed (pip install {('python-calamine' if engine=='calamine' else engine)})")
        except Exception as e:
            attempts.append(f"{engine}: {type(e).__name__}: {e}")

    if kind == "html" or True:  # always try html as last resort even if kind unknown
        try:
            tables = pd.read_html(io.BytesIO(file_bytes))
            if tables:
                return "html", [f"Sheet1 ({len(tables)} table(s) found, using 1st)"], attempts + ["html: OK"]
        except Exception as e:
            attempts.append(f"html: {type(e).__name__}: {e}")

    return None, [], attempts


@st.cache_data(show_spinner=False)
def read_preview(file_bytes, filename, engine, sheet_name, nrows=40):
    """Cheap, small read — just enough rows to detect the header row and show a preview."""
    bio = io.BytesIO(file_bytes)
    if engine == "csv":
        raw = pd.read_csv(bio, header=None, dtype=str, nrows=nrows)
    elif engine == "html":
        tables = pd.read_html(bio, header=None)
        raw = tables[0].head(nrows)
        raw.columns = range(raw.shape[1])
    else:
        raw = pd.read_excel(bio, sheet_name=sheet_name, header=None, dtype=str, engine=engine, nrows=nrows)
        raw.columns = range(raw.shape[1])
    return raw


@st.cache_data(show_spinner=False)
def read_full(file_bytes, filename, engine, sheet_name, header_row):
    """
    ONE full read of the whole file/sheet, with pandas applying the header natively
    (header=N). This is far more memory-efficient than reading a full header-less
    grid and then slicing/copying it again — that pattern effectively held two full
    copies of a large file in memory at once, which is the likely cause of the app
    crashing on Streamlit Cloud's memory-constrained free tier for large exports.
    dtype=str throughout to preserve leading zeros / exact IDs (e.g. "076646_01").
    Pandas also automatically de-duplicates repeated column names (Col, Col.1, Col.2...).
    """
    bio = io.BytesIO(file_bytes)
    if engine == "csv":
        df = pd.read_csv(bio, header=header_row, dtype=str)
    elif engine == "html":
        tables = pd.read_html(bio, header=header_row)
        df = tables[0]
    else:
        df = pd.read_excel(bio, sheet_name=sheet_name, header=header_row, dtype=str, engine=engine)
    df.columns = [str(c) for c in df.columns]
    return df


def excel_col_letter(idx: int) -> str:
    letters = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def find_header_row(raw_df: pd.DataFrame, keywords, max_scan=20):
    best_row, best_hits = 0, -1
    for i in range(min(max_scan, len(raw_df))):
        row_vals = [str(v).strip().lower() for v in raw_df.iloc[i].tolist()]
        hits = sum(1 for kw in keywords if any(kw in v for v in row_vals))
        if hits > best_hits:
            best_hits, best_row = hits, i
    return best_row


def build_label_map(columns, banner_vals):
    """
    Build a {column_name: display_label} map. Column names themselves are already
    unique (pandas guarantees this via header=N parsing), so this only builds the
    human-friendly label: real Excel column letter + the banner/campaign text
    sitting directly above the header (common in ZeCom trackers where one campaign
    name spans several repeated sub-columns like RRP/SRP/DISC%).
    """
    labels = {}
    for i, col in enumerate(columns):
        letter = excel_col_letter(i)
        banner_val = banner_vals[i] if i < len(banner_vals) else None
        banner_str = (
            str(banner_val).strip()
            if pd.notna(banner_val) and str(banner_val).strip().lower() not in ("", "nan", "none")
            else ""
        )
        display_name = f"Col_{letter}" if str(col).startswith("Unnamed:") else str(col)
        label = f"{letter}: " + (f"{banner_str} — {display_name}" if banner_str else display_name)
        labels[col] = label
    return labels


def clean_id_str(val, normalize=False):
    if pd.isna(val):
        return None
    if isinstance(val, float):
        s = str(int(val)) if val.is_integer() else str(val)
    else:
        s = str(val).strip()
        if s == "" or s.lower() == "nan":
            return None
        if re.fullmatch(r"\d+\.0+", s):
            s = s.split(".")[0]
    if normalize:
        s = s.strip().upper()
    return s if s != "" else None


# ---------------------------------------------------------------------------
# Column-guessing hints
# ---------------------------------------------------------------------------

MARKETPLACE_EAN_COLUMN_HINTS = {
    "Lazada": ["sellersku", "seller sku"],
    "Shopee": ["sku", "seller sku", "sku reference no"],
    "Zalora": ["sellersku", "seller sku"],
    "TikTok Shop": ["seller sku"],
}
MARKETPLACE_HEADER_HINTS = ["sellersku", "seller sku", "sku", "product", "seller"]

CONTENT_EAN_HINTS = ["ean"]
CONTENT_PARENT_HINTS = ["color no", "article no", "colorno", "articleno", "style#", "style #"]
CONTENT_HEADER_HINTS = CONTENT_EAN_HINTS + CONTENT_PARENT_HINTS

ZECOM_PARENT_HINTS = [
    "pim article", "pim_article", "pim style",
    "article no", "articleno", "color no", "colorno",
    "style#", "style #",
]
ZECOM_HEADER_HINTS = ZECOM_PARENT_HINTS + ["price", "srp", "rrp", "md price"]


def guess_column(columns, hints):
    """
    Hint-priority order matters: e.g. for ZeCom files "PIM Article#" (has the
    color suffix, the correct join key) must win over "PIM Style" (truncated,
    wrong), even though "PIM Style" is an *exact* match for a lower-priority
    hint and "PIM Article#" is only a *substring* match for the higher-priority
    hint. So for each hint (highest priority first) we check both exact and
    substring matches before ever moving on to the next hint.
    """
    cols_lower = {c: str(c).strip().lower() for c in columns}
    for h in hints:
        for c, cl in cols_lower.items():
            if cl == h:
                return c
        for c, cl in cols_lower.items():
            if h in cl:
                return c
    return None


# ---------------------------------------------------------------------------
# Reusable "upload + configure" block for one file
# ---------------------------------------------------------------------------

def configure_file(label, uploaded_file, header_hints, key_prefix, default_sheet_hint=None):
    """
    Full pipeline for one uploaded file: sheet pick -> raw preview -> header row
    (auto + manual override) -> build final dataframe + label map.
    Returns (df, labels) or (None, None) if not ready.
    """
    if uploaded_file is None:
        return None, None

    file_bytes = uploaded_file.getvalue()
    engine, sheet_names, attempt_log = list_sheets(file_bytes, uploaded_file.name)

    if engine is None:
        all_missing = all(
            ("not installed" in a) or ("importerror" in a.lower())
            for a in attempt_log
        )
        if all_missing:
            st.error(
                f"**Could not read {label} — but this isn't a problem with your file.**\n\n"
                "None of the Excel-reading packages (openpyxl, python-calamine, xlrd, lxml) "
                "are installed in the Python environment currently running this app:\n\n"
                + "\n".join(f"- {a}" for a in attempt_log)
                + "\n\n**Fix:** close the app, then in the same terminal run:\n\n"
                "```\npython -m pip install -r requirements.txt\n```\n\n"
                "and relaunch with:\n\n"
                "```\npython -m streamlit run app.py\n```\n\n"
                "Using `python -m pip` / `python -m streamlit` (instead of bare `pip` / "
                "`streamlit`) makes sure both use the same Python interpreter — the usual "
                "cause is `pip install` landing in a different Python install than the one "
                "running Streamlit."
            )
        else:
            st.error(
                f"**Could not read {label}.** Tried multiple formats and none worked:\n\n"
                + "\n".join(f"- {a}" for a in attempt_log)
                + "\n\nIf this is a genuine Excel file, try re-saving it as .xlsx from Excel first. "
                "If it was exported from a marketplace seller center, some exports are actually "
                "HTML tables saved with an .xlsx extension — try opening and re-saving it in Excel."
            )
        return None, None

    sheet_name = sheet_names[0]
    if len(sheet_names) > 1:
        default_idx = 0
        if default_sheet_hint:
            for i, s in enumerate(sheet_names):
                if default_sheet_hint.lower() == str(s).lower():
                    default_idx = i
                    break
        sheet_name = st.selectbox(f"{label} — sheet", options=sheet_names, index=default_idx, key=f"{key_prefix}_sheet")

    raw = read_preview(file_bytes, uploaded_file.name, engine, sheet_name)

    auto_header_row = find_header_row(raw, header_hints)
    with st.expander(f"Preview raw rows of {label} (use this to check/pick the header row)"):
        st.dataframe(raw.head(12), use_container_width=True, height=250)

    header_row = st.number_input(
        f"{label} — which row number is the actual header? (0 = first row)",
        min_value=0,
        max_value=500,
        value=int(auto_header_row),
        key=f"{key_prefix}_header_row",
        help="Auto-detected based on keyword matches. Override if the tracker layout changed and this guessed wrong.",
    )
    header_row = int(header_row)

    # If someone picks a header row beyond what the small preview covered, widen
    # the preview just enough to still read the correct banner row above it.
    if header_row >= len(raw):
        raw = read_preview(file_bytes, uploaded_file.name, engine, sheet_name, nrows=header_row + 10)

    df = read_full(file_bytes, uploaded_file.name, engine, sheet_name, header_row)
    banner_vals = raw.iloc[header_row - 1].tolist() if header_row > 0 else [None] * len(df.columns)
    labels = build_label_map(df.columns, banner_vals)
    # Drop fully-empty columns only after labels are built against original positions,
    # so Excel-letter/banner labels stay correctly aligned (dropping first would shift
    # positions and silently corrupt the labels of every column after a dropped one).
    df = df.dropna(axis=1, how="all")
    labels = {k: v for k, v in labels.items() if k in df.columns}
    return df, labels


# ---------------------------------------------------------------------------
# Sidebar — setup
# ---------------------------------------------------------------------------

st.title("🔗 ZeCom Column Mapper")
st.caption(
    "Upload your marketplace export + Content file + ZeCom file. "
    "Get the marketplace file back with ZeCom pricing (or any other column you pick) mapped in."
)

with st.sidebar:
    st.header("1. Setup")
    marketplace = st.selectbox("Marketplace", ["Lazada", "Shopee", "Zalora", "TikTok Shop"])
    region = st.selectbox("Region", ["MY", "PH", "SG"])

    st.header("2. Upload files")
    mp_file = st.file_uploader(f"Marketplace file ({marketplace})", type=["xlsx", "xls", "csv"])
    content_file = st.file_uploader("Content file (EAN → Color No)", type=["xlsx", "xls", "csv"])
    zecom_file = st.file_uploader("ZeCom file (pricing etc.)", type=["xlsx", "xls", "csv"])

    st.header("3. Options")
    normalize_keys = st.checkbox(
        "Normalize join keys (trim spaces, uppercase, strip stray .0)",
        value=True,
        help="Turn on if rows aren't matching due to minor formatting differences between files.",
    )

if not (mp_file and content_file and zecom_file):
    st.info("Upload all three files in the sidebar to get started.")
    st.stop()

# ---------------------------------------------------------------------------
# Step 1 — Marketplace file
# ---------------------------------------------------------------------------

st.subheader("Step 1 — Marketplace file")
mp_ean_hints = MARKETPLACE_EAN_COLUMN_HINTS[marketplace]
mp_df, mp_labels = configure_file(
    f"Marketplace file ({marketplace})", mp_file, MARKETPLACE_HEADER_HINTS, "mp"
)
if mp_df is None:
    st.stop()

mp_ean_guess = guess_column(mp_df.columns, mp_ean_hints)
mp_ean_col = st.selectbox(
    "Which column is the EAN / Seller SKU?",
    options=list(mp_df.columns),
    index=list(mp_df.columns).index(mp_ean_guess) if mp_ean_guess in mp_df.columns else 0,
    format_func=lambda c: mp_labels.get(c, c),
)
st.dataframe(mp_df.head(5), use_container_width=True, height=180)

# ---------------------------------------------------------------------------
# Step 2 — Content file
# ---------------------------------------------------------------------------

st.divider()
st.subheader("Step 2 — Content file (bridge: EAN → Color No)")
content_df, content_labels = configure_file(
    "Content file", content_file, CONTENT_HEADER_HINTS, "content"
)
if content_df is None:
    st.stop()

content_ean_guess = guess_column(content_df.columns, CONTENT_EAN_HINTS)
content_parent_guess = guess_column(content_df.columns, CONTENT_PARENT_HINTS)

cc1, cc2 = st.columns(2)
with cc1:
    content_ean_col = st.selectbox(
        "EAN column in Content file",
        options=list(content_df.columns),
        index=list(content_df.columns).index(content_ean_guess) if content_ean_guess in content_df.columns else 0,
        format_func=lambda c: content_labels.get(c, c),
    )
with cc2:
    content_parent_col = st.selectbox(
        "Color No / Article No / Style# column in Content file",
        options=list(content_df.columns),
        index=list(content_df.columns).index(content_parent_guess) if content_parent_guess in content_df.columns else 0,
        format_func=lambda c: content_labels.get(c, c),
    )
st.dataframe(content_df[[content_ean_col, content_parent_col]].head(5), use_container_width=True, height=150)

# ---------------------------------------------------------------------------
# Step 3 — ZeCom file
# ---------------------------------------------------------------------------

st.divider()
st.subheader("Step 3 — ZeCom file (pricing + other columns)")
zecom_df, zecom_labels = configure_file(
    "ZeCom file", zecom_file, ZECOM_HEADER_HINTS, "zecom", default_sheet_hint=region
)
if zecom_df is None:
    st.stop()

zecom_parent_guess = guess_column(zecom_df.columns, ZECOM_PARENT_HINTS)
zecom_parent_col = st.selectbox(
    "PIM_Article# / Color No / Style# column in ZeCom file",
    options=list(zecom_df.columns),
    index=list(zecom_df.columns).index(zecom_parent_guess) if zecom_parent_guess in zecom_df.columns else 0,
    format_func=lambda c: zecom_labels.get(c, c),
)

other_zecom_cols = [c for c in zecom_df.columns if c != zecom_parent_col]
st.caption(
    "Tracker columns often repeat per campaign tier (BAU / Payday / Mega / Shopee-specific, etc). "
    "Each option below shows its real Excel column letter and the campaign banner text above it — "
    "use that to pick the right tier(s)."
)
selected_zecom_cols = st.multiselect(
    "Which ZeCom columns do you want mapped into the output?",
    options=other_zecom_cols,
    format_func=lambda c: zecom_labels.get(c, c),
)

if not selected_zecom_cols:
    st.warning("Pick at least one ZeCom column above to see mapped results.")
    st.dataframe(zecom_df.head(5), use_container_width=True, height=180)
    st.stop()

st.dataframe(zecom_df[[zecom_parent_col] + selected_zecom_cols].head(5), use_container_width=True, height=150)

# ---------------------------------------------------------------------------
# Build the join
# ---------------------------------------------------------------------------

st.divider()
st.subheader("Step 4 — Mapped output")

mp_df["_EAN_KEY"] = mp_df[mp_ean_col].apply(lambda v: clean_id_str(v, normalize_keys))
content_df["_EAN_KEY"] = content_df[content_ean_col].apply(lambda v: clean_id_str(v, normalize_keys))
content_df["_PARENT_KEY"] = content_df[content_parent_col].apply(lambda v: clean_id_str(v, normalize_keys))
zecom_df["_PARENT_KEY"] = zecom_df[zecom_parent_col].apply(lambda v: clean_id_str(v, normalize_keys))

dup_parents = zecom_df["_PARENT_KEY"].value_counts()
dup_parents = set(dup_parents[dup_parents > 1].index) - {None}

ean_to_parent = (
    content_df.dropna(subset=["_EAN_KEY"])
    .drop_duplicates(subset=["_EAN_KEY"], keep="first")
    .set_index("_EAN_KEY")["_PARENT_KEY"]
)
zecom_lookup = (
    zecom_df.dropna(subset=["_PARENT_KEY"])
    .drop_duplicates(subset=["_PARENT_KEY"], keep="first")
    .set_index("_PARENT_KEY")[selected_zecom_cols]
)

mp_df["_PARENT_KEY"] = mp_df["_EAN_KEY"].map(ean_to_parent)
mapped = mp_df.join(zecom_lookup, on="_PARENT_KEY")

for c in selected_zecom_cols:
    mapped[c] = mapped[c].where(mapped["_PARENT_KEY"].notna(), "Not Available (no EAN→Color No match)")
    mapped[c] = mapped[c].where(
        ~(mapped["_PARENT_KEY"].notna() & mapped[c].isna()),
        "Not Available (Color No not in ZeCom)",
    )
mapped["ZeCom_Duplicate_Flag"] = mapped["_PARENT_KEY"].apply(
    lambda k: "⚠ Multiple ZeCom entries for this Color No" if k in dup_parents else ""
)

output_label_map = {c: zecom_labels.get(c, c) for c in selected_zecom_cols}
output_cols = list(mp_df.columns.drop(["_EAN_KEY", "_PARENT_KEY"])) + selected_zecom_cols + ["ZeCom_Duplicate_Flag"]
final_df = mapped[output_cols].copy()
# Give the appended ZeCom columns their descriptive labels in the actual output file
final_df = final_df.rename(columns=output_label_map)

total_rows = len(final_df)
no_content_match = int((mapped["_PARENT_KEY"].isna()).sum())
first_sel = selected_zecom_cols[0]
no_zecom_match = int(
    ((mapped["_PARENT_KEY"].notna()) & (mapped[first_sel].astype(str).str.startswith("Not Available"))).sum()
)
dup_hits = int((mapped["ZeCom_Duplicate_Flag"] != "").sum())

s1, s2, s3, s4 = st.columns(4)
s1.metric("Total rows", total_rows)
s2.metric("EAN not found in Content", no_content_match)
s3.metric("Color No not found in ZeCom", no_zecom_match)
s4.metric("Duplicate ZeCom entries hit", dup_hits)

st.dataframe(final_df.head(20), use_container_width=True, height=350)

buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    final_df.to_excel(writer, index=False, sheet_name="Mapped Output")
buf.seek(0)

st.download_button(
    "⬇️ Download mapped file (.xlsx)",
    data=buf,
    file_name=f"{marketplace}_{region}_ZeCom_Mapped.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
