# ZeCom Column Mapper (v2)

Upload your marketplace export (Lazada / Shopee / Zalora / TikTok Shop) and get
the same file back with ZeCom pricing (or any other ZeCom column you pick)
mapped in — without touching your original file's structure.

## Why the Content file is needed
Your marketplace files only carry the EAN (Seller SKU). ZeCom pricing lives at
the **Article No / Color No (parent)** level. So the app bridges:
`Marketplace EAN → Content file → Color No → ZeCom → your columns`.

## How to run

**Locally:**
```bash
pip install -r requirements.txt
streamlit run app.py
```

**On Streamlit Community Cloud:** push `app.py` + `requirements.txt` to your GitHub
repo, then deploy/reboot from share.streamlit.io — it installs `requirements.txt`
automatically, you don't need to run pip yourself.

> If the app ever shows "not installed" for every single reading engine
> (including plain `openpyxl`), that means the whole `pip install -r requirements.txt`
> step failed during build — usually because one package in the list couldn't
> build on that platform, which aborts the entire install. Check the build log
> via **Manage app → logs** on share.streamlit.io to see the real pip error, and
> keep `requirements.txt` as minimal as possible (this is why `python-calamine`
> was dropped from the default list — it needs a compiled Rust wheel that isn't
> always available, and openpyxl alone already handles standard marketplace
> exports fine). After changing `requirements.txt`, reboot the app so it
> actually rebuilds the environment — Streamlit Cloud doesn't always do this
> automatically.


Then in the browser:
1. Pick your marketplace + region in the sidebar.
2. Upload the 3 files: Marketplace export, Content file, ZeCom file.
3. If a file has multiple sheets (e.g. MY + SG in one workbook), pick the sheet.
4. Check the auto-detected header row (a preview of the raw rows is shown) —
   override it if the tracker's layout changed and the guess is wrong.
5. Confirm/correct the EAN and Color No / PIM_Article# / Style# columns.
6. Pick which ZeCom columns to map in. Repeated columns (e.g. "MY RRP"
   appearing once per campaign tier) are shown as separate options, each
   labelled with its real Excel column letter and the campaign/tier banner
   text above it (e.g. `BV: BAU (35%) — PH EC RRP`), so nothing gets mixed up.
7. Review the summary counts, then download the mapped .xlsx — same rows and
   columns as your original marketplace file, plus the new ones.

## Built for trackers that change often
This was rebuilt after testing against real PH/MY/SG ZeCom tracker files,
which don't share the same headers, column order, or even column names
(`PIM Article#` vs `Style#` vs `STYLE#`) release to release. To handle that:

- **Sheet + header row are never hardcoded.** Header row is auto-detected by
  scanning the first 20 rows for keyword hits, but you can always override it
  with the row-picker, and a raw preview is shown so you can check.
- **Repeated column names never silently collide.** Many ZeCom trackers reuse
  the same column name (e.g. "PH EC RRP") once per campaign tier. Each
  occurrence gets a unique internal name and a label built from its real Excel
  column letter + the banner text sitting above it, so you can tell them apart
  and the join never accidentally picks the wrong one.
- **Multi-format file reading.** Tries openpyxl → python-calamine → xlrd →
  HTML-table parsing (some marketplace exports are actually HTML tables saved
  with an `.xlsx` extension) in order, and tells you exactly which of those
  were tried and why they failed if none work.
- **Optional key normalization** (trim spaces / uppercase / strip stray `.0`)
  for when two files format the same key slightly differently.

## Notes
- Marketplace SKU column auto-detected per platform:
  Lazada → SellerSKU, Zalora → SellerSku, Shopee → SKU, TikTok → Seller sku
  (you can override the guess in the UI).
- Rows are never silently dropped or blanked:
  - "Not Available (no EAN match)" = EAN wasn't found in the Content file
  - "Not Available (Color No not in ZeCom)" = Color No has no ZeCom pricing
  - Duplicate PIM_Article#/Color No entries in ZeCom are flagged, not silently
    averaged or overwritten.
