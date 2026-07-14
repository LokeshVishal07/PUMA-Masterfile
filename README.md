# ZeCom Column Mapper

Upload your marketplace export (Lazada / Shopee / Zalora / TikTok Shop) and get
the same file back with ZeCom pricing (or any other ZeCom column you pick)
mapped in — without ever touching your original file's structure.

## Why the Content file is needed
Your marketplace files only carry the EAN (Seller SKU). ZeCom pricing lives at
the **Article No / Color No (parent)** level — it has no EAN. So the app
bridges: `Marketplace EAN → Content file → Color No → ZeCom → your columns`.

## How to run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then in the browser:
1. Pick your marketplace + region in the sidebar.
2. Upload the 3 files: Marketplace export, Content file, ZeCom file.
3. Confirm/correct the auto-detected EAN and Color No / PIM_Article# columns.
4. Pick which ZeCom columns to bring in (labelled with real Excel column
   letters, e.g. `D: Price`).
5. Review the summary counts, then download the mapped .xlsx — same rows and
   columns as your original marketplace file, plus the new ones.

## Notes
- Handles banner rows / double headers automatically (scans the first 15 rows
  for the real header).
- Marketplace SKU column auto-detected per platform:
  Lazada → SellerSKU, Zalora → SellerSku, Shopee → SKU, TikTok → Seller sku
  (you can override the guess in the UI).
- Rows are never silently dropped or blanked:
  - "Not Available (no EAN match)" = EAN wasn't found in the Content file
  - "Not Available (Color No not in ZeCom)" = Color No has no ZeCom pricing
  - Duplicate PIM_Article# entries in ZeCom are flagged, not silently averaged.
