# Cutlist / BOM PWA (from `TEMPLATE.xlsx`)

This repo contains a **static, responsive, offline-capable PWA** that mirrors the Excel template logic in `TEMPLATE.xlsx`.

## What it does
- Shows a **spreadsheet-like grid** for `KITCHEN` and `WARDROBE`
- Lets you edit the **input cells** (dimensions, counts, material codes)
- Recalculates outputs using the same formulas (supports `IF`, `SUM`, `TRIM`, arithmetic, `&` concat, cell refs/ranges)
- Works **offline** after first load (service worker caches all assets)

## Run locally
Service workers require `http://` (not `file://`).

From the repo root:

```bash
cd public
python3 -m http.server 5173
```

Then open `http://localhost:5173`.

## Install as PWA (mobile)
- **Android Chrome**: open the site → browser menu → “Install app”
- After the first load, try airplane mode: the app should still open and work.

## Updating when the Excel changes
If you edit `TEMPLATE.xlsx`, regenerate the JSON model:

```bash
python3 scripts/extract_excel_model.py \
  --xlsx /Users/vd519252/vicky-codes/Aravind-instalation/TEMPLATE.xlsx \
  --out  /Users/vd519252/vicky-codes/Aravind-instalation/public/model.json
```

Then refresh the web page.

## Files
- `public/index.html`: UI shell
- `public/app.js`: UI + recalculation
- `public/formula.js`: minimal Excel formula evaluator
- `public/model.json`: extracted workbook (values + formulas)
- `public/sw.js`: offline caching service worker
- `public/manifest.webmanifest`: PWA manifest
- `scripts/extract_excel_model.py`: XLSX → JSON extractor (no external deps)


