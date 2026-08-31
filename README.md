# RFQ Executive Intelligence Dashboard

A portable, browser-based RFQ dashboard generated from `HP.xlsx`. It is designed for executive review and commercial/tender follow-up, with one shared filter context across every tab.

## Files

- `index.html` — the full interactive dashboard (no framework or chart library required).
- `data.json` — dashboard-ready data generated from the Excel workbook.
- `convert.py` — dependency-free Python converter for refreshing `data.json` from a new Excel file with the same structure.
- `README.md` — setup, metric definitions and refresh instructions.

## Dashboard structure

The dashboard contains six views: Executive Summary, Country, Customer, Open RFQs, Under Pipeline, and Customer Performance. The global filter bar includes Year, Quarter, Month, Week, Country, Customer, Status and Vendor, plus Reset. Filters apply consistently across all views.

The Executive Summary displays Total RFQ Value in EUR, RFQ Count, PO Received, Under Pipeline, Open RFQs, Bid Lost and RFQ Declined, followed by trend, status, geography, customer and closing-exposure visuals. Country and Customer views support click-through drilldowns. Open RFQs and Under Pipeline have operational registers with dynamic deadline calculations.

## Run the dashboard

Because browsers generally block `fetch()` from `file://` pages, serve the folder locally instead of double-clicking the HTML file.

```bash
cd rfq_dashboard
python -m http.server 8000
```

Then open `http://localhost:8000` in your browser.

## Refresh with a newer Excel file

Place the replacement workbook in this folder and run:

```bash
python convert.py HP.xlsx data.json
```

Then refresh the browser. `convert.py` uses only the Python standard library; no packages need to be installed.

## Metric definitions

- **Total RFQ Value**: sum of the Excel `Value` column, displayed as EUR per the requested dashboard convention.
- **RFQ Count**: number of RFQ rows in the current filter context.
- **Under Pipeline**: stages `Awaiting PO`, `Clarification Complete`, or `Clarification Completed`.
- **Open RFQs**: stages `Floated To Supplier/Internal`, `Quote Received Supplier/Internal`, or `Quoted To Client`.
- **PO Received / Bid Lost / RFQ Declined**: direct status mapping from the Stage column.
- **Days to Close**: `Closing Date - today's date`; negative values are overdue. This recalculates every time the dashboard opens.
- **RFQ Window Days**: `Closing Date - RFQ Received`.
- **Customer Last Order**: latest non-blank `PO Received` date for the customer.
- **Days Since**: today minus Customer Last Order.
- **Relationship Months**: first RFQ Received date through today.
- **Lifetime Value**: sum of RFQ `Value` for that customer in the active global filter context.

## Data notes

The source workbook contains 251 RFQ records. The generated snapshot contains 100 Open RFQs, 53 PO Received, 35 Under Pipeline, 27 Bid Lost and 36 RFQ Declined. The source `Value` field totals approximately EUR 31.04M before filters.
