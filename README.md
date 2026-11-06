# HP Valves — RFQ Intelligence Dashboard

Executive dashboard for the HP Valves RFQ tracker (`HP-FIN.xlsx`, sheet **MAIN TABLE**).

**Architecture (Oct 2026):** the frontend is one static file (`index.html`) hosted on **GitHub Pages**; the data and the user accounts live in **Supabase**. Users sign in with e-mail + password; administrators upload the Excel on the Data & upload tab and every user sees the new figures.

## Files

| File | Purpose |
|---|---|
| `index.html` | The whole frontend (HTML + CSS + JS + SheetJS + supabase-js + world outlines + country artwork). Deploy this to GitHub Pages. |
| `supabase_schema.sql` | Database schema, admin list, row-level security. Run once in Supabase → SQL Editor. |
| `convert.py` | Builds `data.json` from the workbook on a PC (optional — the browser upload does the same). |
| `push_to_supabase.py` | Publishes `data.json` to the database from the command line (optional, needs the service key). |
| `data.json` | Latest dataset as JSON (for inspection / scripted use). |
| `images/` | Transparent PNG artwork per country (also embedded in `index.html`). |

## One-time setup

1. **Database** — Supabase → SQL Editor → paste `supabase_schema.sql` → Run. It creates `app_admins`, `rfq_uploads`, `rfq_lines`, the `rfq_current` view and the security policies, and registers the two admins:
   - `Paul.Paran@meraki-energy.com`
   - `it.admin@meraki-energy.com`
2. **Users** — Supabase → Authentication → Users → *Add user* (e-mail + password, "Auto confirm"). Anyone added here can sign in and view; only e-mails in `app_admins` see **Data & upload** and can upload.
   To add another admin: `insert into public.app_admins(email) values ('name@company.com');`
3. **Auth URL** — Supabase → Authentication → URL Configuration → set *Site URL* to the GitHub Pages address (e.g. `https://<user>.github.io/<repo>/`) and add it to *Redirect URLs*. This is needed for the "Forgot password?" e-mail link.
4. **GitHub Pages** — commit `index.html` (and `images/` if you like) to the repo, Settings → Pages → deploy from the branch. The project URL and publishable key are already in `index.html` (top of the script, `SUPABASE_URL` / `SUPABASE_KEY`). The publishable key is safe in a browser: the SQL policies decide what each signed-in user may do.
5. **First data** — sign in as an admin, open **Data & upload**, drop `HP-FIN.xlsx`. The embedded sample data is shown to admins until the first upload is saved.

## Updating the data

Sign in as an admin → **Data & upload** → drop the new workbook. Three steps, each with a button:
1. **Preview** — the file is parsed in the browser and the dashboard shows it on your screen only (marked PREVIEW).
2. **Save to Supabase** — writes it to the database as a new upload (the last 5 are kept). **Discard** throws the preview away.
3. **Refresh database** — reloads the dashboard from Supabase. Every other user gets the new data on their next page load (or by pressing Refresh database if they are admins). No GitHub commit is needed for data changes any more.

Command-line alternative: `python convert.py HP-FIN.xlsx --no-inject` then `python push_to_supabase.py` with `SUPABASE_SERVICE_KEY` set (never commit that key).

## Roles

| | Viewer (any user) | Admin (`app_admins`) |
|---|---|---|
| Executive, Countries, Customers, Open quotes, Under pipeline, Customer performance, Shipment | ✓ | ✓ |
| Data & upload tab, uploading, Download data.json / CSV | — | ✓ |

The tab is hidden in the UI *and* uploads are blocked by row-level security, so a non-admin cannot write even with developer tools.

## Workbook rules (apply to both the browser upload and convert.py)

- The header row is found automatically (must contain *Customer* and *RFQ Received*); headers match ignoring case, spacing and punctuation (`Staus` / `Status`, `BUDGETORY` / `Budgetary Quote Y/N`).
- **Budgetary Quote Y/N** (any header starting with "Budget…"): Y / Yes = excluded when the *Exclude budgetary quotes* toggle is on; N / No = always shown. Without such a column the old keyword detection is used.
- Unknown columns (e.g. Y–AC) are kept per row under `extra` for later use.
- If the `Staus` formula column has no saved values, status is derived from *Stage* with the same rule as the workbook formula.
- **End-Client is the client dimension** (Nov 2026). The *End-Client* column drives the End-Client filter, tabs, charts and performance matrix; the *Customer* column is kept as the trading partner (shown as "Customer (trader)" in tables and row details, and listed as "via …" on the end-client profile). If a workbook has no End-Client column, *Customer* is used as the client dimension as before. A blank End-Client falls back to the Customer value.
- **Statuses are dynamic**: every distinct value in `Staus` appears as typed in the Status filter, KPI tiles, ribbon and charts; each is classified by keyword into a role (PO / pipeline / open / lost / declined) that drives page logic. Unmatched values show as their own category and are listed on Data & upload.
- Country names ("Qatar") are normalised to tracker codes (QTR). **TURK = Turkmenistan** (not Türkiye). Dates are read as Excel serials (no time-zone day shift).
- `MGE Cost Price` is shown in row details only; its currency is inconsistent with `Value`, so no margin KPI is built on it.

## Pages

1. **Executive summary** — total quoted value + RFQ count, one KPI tile per status, value ribbon, monthly intake, status mix, value by country/customer, win-rate trend, turnaround, declined reasons, and the **customer footprint** (3D globe; selecting a country shows its artwork with the customer count and ranked customer list).
2. **Summary** — every tracker column for the current filters in one sortable table (incl. extra columns Y–AC), with the KPI strip; the Executive KPI tiles (Total, PO Received, Under Pipeline, Open Quotes, Open RFQ, Bid Lost, RFQ Declined) are clickable and open this tab filtered to that status.
3. **Countries** — country cards, profile, monthly intake, customer globe/artwork, full RFQ table with RFQ Received, Closing Date, Days to close and Due in.
3. **Customers** — searchable list, profile, monthly intake, stage and country charts (count · value), full table.
4. **Open quotes** / 5. **Under pipeline** — registers with aging, stage and closing-date views.
6. **Customer performance** — recency segments, win rate by customer, matrix (last order, days since, months active, lifetime value, RFQ count …).
7. **Shipment** — placeholder (Coming soon).
8. **Data & upload** (admins) — Excel upload to the database, current-dataset summary, Download data.json / CSV.

Filters (Year, Quarter, Month, Week, Country, Customer, Status, Vendor, search, Reset) apply to every page; three themes; CSV export on every table.

## Local / offline use

Set `SUPABASE_URL = ''` at the top of the script in `index.html` and the page runs in the old single-file mode (embedded data, `python convert.py` to refresh, everyone is admin). Keep a separate copy for that; the GitHub copy should keep the Supabase settings.
