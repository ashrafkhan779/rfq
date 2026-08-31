# HP Valves — RFQ Intelligence Dashboard

A self-contained, offline-capable executive dashboard built from `HP.xlsx` (sheet **MAIN TABLE**).
No server, no build step, no external chart library: open `index.html` by double-clicking it.

## Files

| File | Purpose |
|---|---|
| `index.html` | The dashboard. Single file (HTML + CSS + JS + embedded dataset). Works from `file://`. |
| `convert.py` | Reads `HP.xlsx`, cleans and enriches every RFQ line, writes `data.json` and embeds it into `index.html`. |
| `data.json` | The clean dataset + metadata, for inspection or reuse in other tools. |
| `README.md` | This file. |

## Refresh the dashboard with a new HP.xlsx

```bash
pip install pandas openpyxl        # once
python convert.py                  # HP.xlsx in the same folder
python convert.py "C:\path\HP.xlsx" --sheet "MAIN TABLE"
```

`convert.py` prints a status summary and rewrites both `data.json` and the dataset inside `index.html`.
Keep `index.html` next to `convert.py`; the script only replaces the `<script id="hp-data">` block, so any design changes you make to `index.html` are preserved.

If the embedded block is empty (e.g. you copied a fresh template), the page falls back to fetching `data.json` — that only works over `http://`, not `file://`, which is why the data is embedded.

## Pages

1. **Executive summary** — Total quoted value (EUR) + RFQ count, then PO Received, Under Pipeline, Open RFQ's, Bid Lost, RFQ Declined (value, count, share), a value-share ribbon, auto-generated insights, and charts: monthly intake (count by status + value line), status mix, value by country, top customers, win-rate trend, RFQ-to-closing turnaround, declined-reason categories, deal-owner workload.
2. **Countries** — "All countries" plus one card per country (value, RFQs, POs, win rate, status mini-bar). Clicking a card sets the Country filter, so every chart, KPI and the full-data table follow. The table includes **RFQ Received**, **Closing Date**, **Days to close** (closing − received) and **Due in** (closing − today, for open/pipeline lines; red when overdue).
3. **Customers** — Searchable customer list with highlight data (value, RFQs, POs, country, last RFQ). Clicking a customer sets the Customer filter and shows the full profile: first/latest RFQ, last PO, PO value, win rate, status ribbon, monthly intake, stage breakdown and the full RFQ table.
4. **Open RFQ's** — Count, value (with and without budgetary quotes), overdue count, closing within 7 days, median age; charts by stage, aging, country, closing-date week and biggest open opportunities; register sorted by closing date.
5. **Under pipeline** — Count, value, Awaiting-PO value, clarification stage, days since quotation; stage funnel, aging, country, customer; register with quotation and clarification dates.
6. **Customer performance** — Matrix per customer: Country, Segment, **Last order**, **Days since**, **Months active**, **Lifetime value** (PO received, EUR), POs, Avg PO, **RFQ count**, Quoted value, Win rate, Open/pipeline, Avg days to PO, First/Latest RFQ. Segments: Active (≤90 d since order), Cooling (≤180), At risk (≤365), Dormant (>365), No order yet. Click a customer name to open its Customers page.

## Filters (global, apply to every page)

Year · Quarter · Month · Week · Country · Customer · Status · Vendor · free-text search · **Reset filters** button.
Options cascade (e.g. choosing 2025 only lists 2025 months). Active filters appear as removable chips.
Clicking bars, slices, country cards or customer rows also sets the matching filter (cross-filtering).

**Exclude budgetary quotes** toggle: lines whose description/feedback contain *BDGTRY / BUDGETORY / Budgetary / Indicative* are flagged. Two such quotes (SEC ≈ €12.9M, SWA ≈ €11.8M) account for ~90% of the open value, so the toggle shows the bankable pipeline.

Themes: **Control room** (dark), **Datasheet** (light), **Boardroom** (navy). Every table exports the current view to CSV.

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
- `MGE Cost Price` is shown as recorded in the row details but is not used for margin, because its currency is inconsistent with `Value` (often ≈ 4.8× the EUR value, suggesting SAR/AED). Confirm the currency before adding a margin KPI.
- Declined reasons are free text; `convert.py` buckets them (`declinedGroup`) with keyword rules that you can edit in `DECLINE_BUCKETS`.
- Rows with no Customer and no RFQ date are dropped.

## Design notes (skills applied)

- **skill-for-dashbiard**: executive-first layout, insights that explain what/why/risk/opportunity, multi-theme, cross-filtering, drill-through, CSV export, aging and Pareto analysis.
- **uiux-pro-max-design-intelligence** priority checks: WCAG contrast on all themes (status colours darkened on light theme), 44 px touch targets on nav/list/buttons, visible focus rings, SVG icons only, 16 px base type with tabular numerals, semantic colour tokens (no raw hex in components), motion ≤ 300 ms and disabled under `prefers-reduced-motion`, status never conveyed by colour alone (always paired with a label), responsive down to mobile.
