"""
PUMA ZeCom Column Mapper
=========================
Upload a MARKETPLACE file (Lazada / Shopee / Zalora / TikTok) that only has EAN
(Seller SKU) at the row level, and get back the SAME file with ZeCom pricing
(and any other ZeCom columns you choose) mapped in.

JOIN CHAIN (ZeCom has no EAN, so we bridge through Content):
  Marketplace EAN  --(Content file)-->  Color No / Article No  --(ZeCom file)-->  Selected columns

Output = original marketplace file, same rows/columns/order, with new columns appended.
"""

import io
import re
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="ZeCom Column Mapper", layout="wide")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def clean_id_str(val):
    """Turn 18890032587.0 -> '18890032587', keep strings/blank as-is."""
    if pd.isna(val):
        return None
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val)
    s = str(val).strip()
    if s == "" or s.lower() == "nan":
        return None
    # strip trailing .0 artifacts that survive as strings
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


def excel_col_letter(idx: int) -> str:
    """0-indexed column position -> Excel column letter (A, B, ..., AA, ...)."""
    letters = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def find_header_row(raw_df: pd.DataFrame, keywords, max_scan=15):
    """Scan first N rows for a row containing enough keyword hits to be a header."""
    best_row, best_hits = 0, -1
    for i in range(min(max_scan, len(raw_df))):
        row_vals = [str(v).strip().lower() for v in raw_df.iloc[i].tolist()]
        hits = sum(1 for kw in keywords if any(kw in v for v in row_vals))
        if hits > best_hits:
            best_hits, best_row = hits, i
    return best_row


@st.cache_data(show_spinner=False)
def read_any_excel(file_bytes, filename, header_hint_keywords, sheet_name=None):
    """
    Robust reader: tries a couple of engines, auto-detects the header row by
    scanning for keyword hits (handles double-header / banner-row exports).
    Returns (dataframe, detected_header_row_index).
    """
    bio = io.BytesIO(file_bytes)
    engines_to_try = ["openpyxl", "calamine"]
    last_err = None
    for engine in engines_to_try:
        try:
            bio.seek(0)
            if filename.lower().endswith(".csv"):
                raw = pd.read_csv(bio, header=None, dtype=str)
            else:
                xls = pd.ExcelFile(bio, engine=engine)
                sheet = sheet_name if sheet_name in xls.sheet_names else xls.sheet_names[0]
                raw = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=str)
            header_row = find_header_row(raw, header_hint_keywords)
            df = raw.iloc[header_row + 1:].copy()
            df.columns = [str(c).strip() for c in raw.iloc[header_row].tolist()]
            df = df.dropna(axis=1, how="all")
            df = df.loc[:, ~df.columns.str.fullmatch(r"nan|None|", case=False, na=False)]
            df = df.reset_index(drop=True)
            return df, header_row
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"Could not read {filename}: {last_err}")


MARKETPLACE_EAN_COLUMN_HINTS = {
    "Lazada": ["sellersku", "seller sku"],
    "Shopee": ["sku", "seller sku", "sku reference no"],
    "Zalora": ["sellersku", "seller sku"],
    "TikTok Shop": ["seller sku"],
}

CONTENT_EAN_HINTS = ["ean"]
CONTENT_PARENT_HINTS = ["color no", "article no", "colorno", "articleno"]
ZECOM_PARENT_HINTS = ["pim_article", "pim article", "article no", "articleno"]


def guess_column(columns, hints):
    cols_lower = {c: str(c).strip().lower() for c in columns}
    # exact-ish match first
    for c, cl in cols_lower.items():
        for h in hints:
            if cl == h:
                return c
    # then contains
    for c, cl in cols_lower.items():
        for h in hints:
            if h in cl:
                return c
    return None


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

# ---------------------------------------------------------------------------
# Main flow
# ---------------------------------------------------------------------------

if not (mp_file and content_file and zecom_file):
    st.info("Upload all three files in the sidebar to get started.")
    st.stop()

# --- Read Marketplace file ---
mp_ean_hints = MARKETPLACE_EAN_COLUMN_HINTS[marketplace]
try:
    mp_df, mp_header_row = read_any_excel(
        mp_file.getvalue(), mp_file.name, mp_ean_hints + ["seller", "sku", "product"]
    )
except Exception as e:
    st.error(f"Could not read the marketplace file: {e}")
    st.stop()

mp_ean_col = guess_column(mp_df.columns, mp_ean_hints)

st.subheader("Step 1 — Marketplace file")
c1, c2 = st.columns([2, 1])
with c1:
    mp_ean_col = st.selectbox(
        "Which column is the EAN / Seller SKU?",
        options=list(mp_df.columns),
        index=list(mp_df.columns).index(mp_ean_col) if mp_ean_col in mp_df.columns else 0,
    )
with c2:
    st.metric("Rows detected", len(mp_df))
st.dataframe(mp_df.head(5), use_container_width=True, height=180)

# --- Read Content file ---
try:
    content_df, content_header_row = read_any_excel(
        content_file.getvalue(), content_file.name, CONTENT_EAN_HINTS + CONTENT_PARENT_HINTS
    )
except Exception as e:
    st.error(f"Could not read the Content file: {e}")
    st.stop()

content_ean_col = guess_column(content_df.columns, CONTENT_EAN_HINTS)
content_parent_col = guess_column(content_df.columns, CONTENT_PARENT_HINTS)

st.subheader("Step 2 — Content file (bridge: EAN → Color No)")
cc1, cc2 = st.columns(2)
with cc1:
    content_ean_col = st.selectbox(
        "EAN column in Content file",
        options=list(content_df.columns),
        index=list(content_df.columns).index(content_ean_col) if content_ean_col in content_df.columns else 0,
    )
with cc2:
    content_parent_col = st.selectbox(
        "Color No / Article No column in Content file",
        options=list(content_df.columns),
        index=list(content_df.columns).index(content_parent_col) if content_parent_col in content_df.columns else 0,
    )
st.dataframe(content_df[[content_ean_col, content_parent_col]].head(5), use_container_width=True, height=150)

# --- Read ZeCom file ---
try:
    zecom_df, zecom_header_row = read_any_excel(
        zecom_file.getvalue(), zecom_file.name, ZECOM_PARENT_HINTS + ["price", "srp"]
    )
except Exception as e:
    st.error(f"Could not read the ZeCom file: {e}")
    st.stop()

zecom_parent_col = guess_column(zecom_df.columns, ZECOM_PARENT_HINTS)

st.subheader("Step 3 — ZeCom file (pricing + other columns)")
zc1, zc2 = st.columns([1, 2])
with zc1:
    zecom_parent_col = st.selectbox(
        "PIM_Article# / Color No column in ZeCom file",
        options=list(zecom_df.columns),
        index=list(zecom_df.columns).index(zecom_parent_col) if zecom_parent_col in zecom_df.columns else 0,
    )

# Dynamic column selector, labelled with real Excel column letters like the voucher tool
other_zecom_cols = [c for c in zecom_df.columns if c != zecom_parent_col]
col_labels = {c: f"{excel_col_letter(list(zecom_df.columns).index(c))}: {c}" for c in zecom_df.columns}

with zc2:
    selected_zecom_cols = st.multiselect(
        "Which ZeCom columns do you want mapped into the output?",
        options=other_zecom_cols,
        default=[c for c in other_zecom_cols if any(k in str(c).lower() for k in ["price", "srp", "mrp"])][:3],
        format_func=lambda c: col_labels.get(c, c),
    )

st.dataframe(zecom_df[[zecom_parent_col] + selected_zecom_cols].head(5), use_container_width=True, height=150)

if not selected_zecom_cols:
    st.warning("Pick at least one ZeCom column to map, then results will appear below.")
    st.stop()

# ---------------------------------------------------------------------------
# Build the join
# ---------------------------------------------------------------------------

st.divider()
st.subheader("Step 4 — Mapped output")

# Clean join keys
mp_df["_EAN_KEY"] = mp_df[mp_ean_col].apply(clean_id_str)
content_df["_EAN_KEY"] = content_df[content_ean_col].apply(clean_id_str)
content_df["_PARENT_KEY"] = content_df[content_parent_col].apply(clean_id_str)
zecom_df["_PARENT_KEY"] = zecom_df[zecom_parent_col].apply(clean_id_str)

# Flag duplicate PIM_Article# entries in ZeCom (data-quality flag, not silently resolved)
dup_parents = zecom_df["_PARENT_KEY"].value_counts()
dup_parents = set(dup_parents[dup_parents > 1].index) - {None}

# EAN -> Color No (first match; flag if an EAN maps to multiple parents)
ean_to_parent = (
    content_df.dropna(subset=["_EAN_KEY"])
    .drop_duplicates(subset=["_EAN_KEY"], keep="first")
    .set_index("_EAN_KEY")["_PARENT_KEY"]
)

# Color No -> selected ZeCom columns (first match if duplicates, since duplicates are flagged separately)
zecom_lookup = (
    zecom_df.dropna(subset=["_PARENT_KEY"])
    .drop_duplicates(subset=["_PARENT_KEY"], keep="first")
    .set_index("_PARENT_KEY")[selected_zecom_cols]
)

mp_df["_PARENT_KEY"] = mp_df["_EAN_KEY"].map(ean_to_parent)

mapped = mp_df.join(zecom_lookup, on="_PARENT_KEY")

# Not Available handling — never silently blank
for c in selected_zecom_cols:
    mapped[c] = mapped[c].where(mapped["_PARENT_KEY"].notna(), "Not Available (no EAN→Color No match)")
    mapped[c] = mapped[c].where(
        ~(mapped["_PARENT_KEY"].notna() & mapped[c].isna()),
        "Not Available (Color No not in ZeCom)",
    )
mapped["ZeCom_Duplicate_Flag"] = mapped["_PARENT_KEY"].apply(
    lambda k: "⚠ Multiple ZeCom entries for this Color No" if k in dup_parents else ""
)

output_cols = list(mp_df.columns.drop(["_EAN_KEY", "_PARENT_KEY"])) + selected_zecom_cols + ["ZeCom_Duplicate_Flag"]
final_df = mapped[output_cols].copy()

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

total_rows = len(final_df)
no_content_match = int((mapped["_PARENT_KEY"].isna()).sum())
no_zecom_match = int(
    ((mapped["_PARENT_KEY"].notna()) & (mapped[selected_zecom_cols[0]].astype(str).str.startswith("Not Available"))).sum()
)
dup_hits = int((mapped["ZeCom_Duplicate_Flag"] != "").sum())

s1, s2, s3, s4 = st.columns(4)
s1.metric("Total rows", total_rows)
s2.metric("EAN not found in Content", no_content_match)
s3.metric("Color No not found in ZeCom", no_zecom_match)
s4.metric("Duplicate ZeCom entries hit", dup_hits)

st.dataframe(final_df.head(20), use_container_width=True, height=350)

# ---------------------------------------------------------------------------
# Download — same structure as marketplace file, plus new columns
# ---------------------------------------------------------------------------

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
