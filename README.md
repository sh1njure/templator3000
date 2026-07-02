# templator3000 — supplier price-file generator

Turns a supplier price-list workbook into the three import files used by the
ERP / POS system.

## Web app (no install)

`index.html` at the repo root is a self-contained browser page: drag in the
Excel file, get the 3 files back. Everything runs client-side (SheetJS is
vendored in `vendor/`, so it works offline — no upload, no CDN).

It is served on **GitHub Pages** in *Deploy from a branch* mode
(**Settings → Pages → Source → Deploy from a branch → this branch → `/ (root)`**).
Because `index.html` sits at the root it becomes the homepage; the `.nojekyll`
file makes Pages serve the assets as-is. Live at
`https://sh1njure.github.io/templator3000/`.

To run it locally: open `index.html` in a browser (or `python3 -m http.server`
from the repo root).

## 4th template — ECI / Magnalux supplier import (optional)

Tick **"Also build ECI + Magnalux supplier import"** and pick the two supplier
price lists (they are read in your browser only — never uploaded or committed,
so the confidential trade prices stay private). This produces
`PRODUCT_SUPPLIER_IMPORT_<stem>.xlsx` (`ProdSuppRecImport` sheet):

- Every product is duplicated into **two rows** — `E18` (ECI) and `M68`
  (Magnalux).
- **Price** is read from the source's *"All candidates (with arithmetic)"*
  column (e.g. `ECI … 10.50. Magnalux 7.68.` → 10.50 to E18, 7.68 to M68).
- **Supplier Product Code** is looked up in that supplier's price list by
  matching product family + colour + variant, confirmed by price. Only
  confident matches are filled — anything uncertain is left blank (so it won't
  be for every row).
- Product Code is the raw source code (no `W/` prefix); date is `=TODAY()`.

Requires the source to have the *"All candidates (with arithmetic)"* column.

## The three outputs

| Output | Sheet | Purpose |
| --- | --- | --- |
| `PRODUCT_WCEW_<name>.xlsx` | `PRODUCT TEMPLATE` | product master (WCEW company) |
| `PRODUCT_LIGHTING_<name>.xlsx` | `PRODUCT TEMPLATE` | product master (LIGHTING company) |
| `SUPPLIER_DETAILS_<name>.xlsx` | `ProdSuppRecImport` | supplier / product-supplier link |

Column structure, number formats, column widths and calculation logic are
copied 1:1 from the reference TERESA examples.

## Usage

Brand, supplier code and prefix default to the usual values, so the common
case is just:

```bash
pip install openpyxl
python3 generate_price_files.py SOURCE.xlsx -o output -n TERESA \
    --sheet "TERESA G9" --teresa-cct
```

| flag | meaning |
| --- | --- |
| `-o/--outdir` | output directory (default `.`) |
| `-n/--name` | stem for the output filenames (default: source stem) |
| `--sheet` | source sheet name (default: first sheet) |
| `--brand` | Bin Location / brand (default `FUMAGALLI`) |
| `--supplier-code` | supplier code (default `W001`) |
| `--prefix` | product-code prefix for LIGHTING/SUPPLIER files (default `W/`) |
| `--teresa-cct` | append CCT suffix from the code (`…Z1L` → ` - 4K`, `…Z1R` → ` - 3K`) |

## Calculation logic

`base` = the source **"Our price (highest of the two)"** column (auto-detected;
falls back to Trade Price / RRP / Cost / Price).

```
WCEW  Current Cost (ex VAT) = base                     (RAW, unrounded *)
LIGHTING Current Cost (ex VAT) = base * 1.10           ("+10%")
LIGHTING Trade Price (ex VAT)  = base * 1.25
LIGHTING Retail Price (ex VAT) = base * 1.45
LIGHTING CPOS Price (inc VAT)  = ceil(base * 1.45 * 1.23)   (round UP, whole)
LIGHTING Price A               = base * 1.35
SUPPLIER Price                 = base * 1.10
SUPPLIER Last Purchased Date   = =TODAY()
```

Static columns: WCEW group `077` / LIGHTING group `009`, analysis `001`,
VAT `3`, Unit `EA`; WCEW markups Trade 25 / Retail 45 / CPOS 45, Price A 35,
B 20, H 10, all other Price columns `-100`; LIGHTING Price B..N `-100`.

Codes are stored as text (`@`) so leading zeros are preserved.

### * WCEW rounding note

The original written spec asked for `round_up_half(base)` on the WCEW cost,
but every reference example row uses the raw unrounded base (e.g. `18.95`, not
`19.0`). The generator follows the examples. To switch to the written-spec
behaviour set `WCEW_ROUND_HALF = True` in `generate_price_files.py`.

`round_up_half(x) = ceil(x * 2) / 2` is still provided and used by callers that
need it.
