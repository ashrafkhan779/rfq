# HP Valves — RFQ Intelligence Dashboard

A self-contained, offline-capable executive dashboard built from `HP-FIN.xlsx` (sheet **MAIN TABLE**). The older `HP.xlsx` layout (Tender Category column, no BUDGETORY column) is still accepted.
No server, no build step, no external chart library: open `index.html` by double-clicking it.

## Files

| File | Purpose |
|---|---|
| `index.html` | The dashboard. Single file (HTML + CSS + JS + SheetJS + world outlines for the 3D globe + embedded dataset). Works from `file://` and GitHub Pages. |
| `convert.py` | Reads `HP.xlsx`, cleans and enriches every RFQ line, writes `data.json` and embeds it into `index.html`. |
| `data.json` | The clean dataset + metadata, for inspection or reuse in other tools. |
| `README.md` | This file. |
| `images/` | Country artwork for the Executive footprint card: transparent-background PNGs for all 10 countries (KSA, UAE, PAK, TURK, ALG, SIR, BNG, QTR, BAH, JRDN); compressed WebP copies are embedded in index.html so the folder is optional. Add more as `images/<name>.png` and map the code in `COUNTRY_IMAGE` inside `index.html`. |

## Refresh the dashboard with a new HP.xlsx

```bash
pip install pandas openpyxl        # once
python convert.py                  # uses HP-FIN.xlsx (or HP.xlsx) in the same folder
python convert.py "C:\path\HP-FIN.xlsx" --sheet "MAIN TABLE"
```

`convert.py` prints a status summary and rewrites both `data.json` and the dataset inside `index.html`.
Keep `index.html` next to `convert.py`; the script only replaces the `<script id="hp-data">` block, so any design changes you make to `index.html` are preserved.

### Or update without Python — the Data & upload tab

Open the **Data & upload** tab, drop the new `HP.xlsx` on it. The workbook is parsed in the browser (SheetJS is embedded; nothing leaves your machine) with the same rules as `convert.py`, and every KPI, chart, table and remark refreshes immediately. The upload log shows row/status counts and warns about missing headers or unrecognised status values.

### Publish live on GitHub Pages

When served over `http(s)` the page loads **`data.json` from the repository first** (with a cache-busting timestamp, `cache: no-store`), and only falls back to the embedded copy if that fails. So the live dashboard updates by replacing one file:

1. Data & upload → drop the new `HP.xlsx` → check the log.
2. Click **Download data.json**.
3. Commit it to the repo as `data.json`, next to `index.html` (GitHub web UI: open the file → Edit, or drag-and-drop onto the repo).
4. GitHub Pages redeploys in ~1 minute; the sidebar footer shows the new "Rebuilt" time.

Opening `index.html` from disk (`file://`) uses the embedded data, so run `python convert.py` when you also want the offline copy refreshed.

## Pages

1. **Executive summary** — Total quoted value (EUR) + RFQ count, then PO Received, Under Pipeline, Open Quote, Bid Lost, RFQ Declined (value, count, share), a value-share ribbon, and charts: monthly intake (count by status with totals above each column, value line with labels), status mix, value by country, top customers, win-rate trend, RFQ-to-closing turnaround, declined-reason categories, and an animated 3D customer globe (distinct customers per country, drag to rotate) with a ranked country list. Clicking a country (or setting the Country filter) swaps the globe for a zoomed map of that country with its customers ranked by quoted value; clearing the filter brings the globe back.
2. **Countries** — "All countries" plus one card per country (value, RFQs, POs, win rate, status mini-bar). Clicking a card sets the Country filter, so every chart, KPI and the full-data table follow. The table includes **RFQ Received**, **Closing Date**, **Days to close** (closing − received) and **Due in** (closing − today, for open/pipeline lines; red when overdue).
3. **Customers** — Searchable customer list with highlight data (value, RFQs, POs, country, last RFQ). Clicking a customer sets the Customer filter and shows the full profile: first/latest RFQ, last PO, PO value, win rate, status ribbon, full-width monthly intake, then "Where the RFQs sit" (count · value per stage) beside a Countries chart (count · value per country), and the full RFQ table.
4. **Open quotes** — Count, value (with and without budgetary quotes), overdue count, closing within 7 days, median age; charts by stage, aging, country, closing-date week and biggest open opportunities; register sorted by closing date.
5. **Under pipeline** — Count, value, Awaiting-PO value, clarification stage, days since quotation; stage funnel, aging, country, customer; register with quotation and clarification dates.
6. **Customer performance** — Matrix per customer: Country, Segment, **Last order**, **Days since**, **Months active**, **Lifetime value** (PO received, EUR), POs, Avg PO, **RFQ count**, Quoted value, Win rate, Open/pipeline, Avg days to PO, First/Latest RFQ. Segments: Active (≤90 d since order), Cooling (≤180), At risk (≤365), Dormant (>365), No order yet. Click a customer name to open its Customers page.
7. **Data & upload** — Excel upload (in-browser), current-dataset summary with status counts, **Download data.json** / clean CSV, and the GitHub publishing steps.

## Filters (global, apply to every page)

Year · Quarter · Month · Week · Country · Customer · Status · Vendor · free-text search · **Reset filters** button.
Options cascade (e.g. choosing 2025 only lists 2025 months). Active filters appear as removable chips.
Clicking bars, slices, country cards or customer rows also sets the matching filter (cross-filtering).

**Exclude budgetary quotes** toggle: driven by the **BUDGETORY** column in the workbook (Yes = excluded when the toggle is on; No = always shown). If a workbook has no BUDGETORY column, the older keyword detection (*BDGTRY / BUDGETORY / Budgetary / Indicative* in the description) is used instead.

Themes: **Control room** (dark), **Datasheet** (light — #CBDDE9 background, #2872A1 accent), **Boardroom** (navy). Every table exports the current view to CSV.

## Definitions

| Metric | Definition |
|---|---|
| Value | `Value` column, treated as EUR. `--`, blanks and 0 count as unpriced (shown as n/a) |
| Win rate | PO Received ÷ (PO Received + Bid Lost) — decided bids only |
| Days to close | Closing Date − RFQ Received |
| Due in | Closing Date − today (open / pipeline lines only) |
| Days to PO | PO Received − RFQ Received |
| Days since quote | today − QTN Date (falls back to RFQ Received) |
| Last order | Most recent PO Received date for the customer |
| Months active | (latest of RFQ / QTN / PO date − first RFQ) ÷ 30.44, minimum 1 |
| Lifetime value | Sum of `Value` for lines with status PO Received |
| Year / Quarter / Month | Derived from RFQ Received; Week uses the sheet's `Week Num` |

## Data notes

- Source header `Staus` (misspelt) is mapped to `status`; the corrected spelling is also accepted.
- Status `Open Quote` is the canonical open state. Older spellings (`Open RFQ's`, `Open RFQ`, `Open Quotes`) are normalised to it by both `convert.py` and the in-browser upload.
- `MGE Cost Price` is shown as recorded in the row details but is not used for margin, because its currency is inconsistent with `Value` (often ≈ 4.8× the EUR value, suggesting SAR/AED). Confirm the currency before adding a margin KPI.
- Declined reasons are free text; `convert.py` buckets them (`declinedGroup`) with keyword rules that you can edit in `DECLINE_BUCKETS`.
- Rows with no Customer and no RFQ date are dropped.
- Country values written as names ("Qatar", "Saudi Arabia") are normalised to the tracker codes (QTR, KSA …) so filters and the map don't split; edit `COUNTRY_ALIASES` / `XL_COUNTRY_ALIAS` to add more.
- Map: country codes are mapped to ISO3 in `ISO3` / `CENTROID` inside `index.html` (KSA, UAE, PAK, JRDN, QTR, BNG, BAH, ALG, SIR, TURK, OMN, KWT, EGY, IRQ, IND + a few European codes). Add a new code there if a new country appears; bubbles use the centroid so tiny states (Bahrain, Qatar) always show.

## Design notes (skills applied)

- **skill-for-dashbiard**: executive-first layout, insights that explain what/why/risk/opportunity, multi-theme, cross-filtering, drill-through, CSV export, aging and Pareto analysis.
- **uiux-pro-max-design-intelligence** priority checks: WCAG contrast on all themes (status colours darkened on light theme), 44 px touch targets on nav/list/buttons, visible focus rings, SVG icons only, 16 px base type with tabular numerals, semantic colour tokens (no raw hex in components), motion ≤ 300 ms and disabled under `prefers-reduced-motion`, status never conveyed by colour alone (always paired with a label), responsive down to mobile.
