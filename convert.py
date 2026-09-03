#!/usr/bin/env python3
"""
convert.py — HP Valves RFQ Intelligence Dashboard
--------------------------------------------------
Reads the "MAIN TABLE" sheet of HP.xlsx, cleans and enriches every RFQ line,
then writes:

  1. data.json   — the clean dataset + metadata (for inspection / other tools)
  2. index.html  — the same JSON is injected into the <script id="hp-data"> tag
                   so the dashboard opens by double-click (file://) with no server.

Usage
  python convert.py                       # uses HP.xlsx next to this script
  python convert.py path/to/HP.xlsx       # explicit source workbook
  python convert.py HP.xlsx --sheet "MAIN TABLE" --html index.html --json data.json

Requires: pandas, openpyxl   (pip install pandas openpyxl)
"""
import argparse
import json
import math
import re
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd

# --------------------------------------------------------------------------
# Column mapping: Excel header  ->  JSON key
# --------------------------------------------------------------------------
COLUMNS = {
    "SY": "sy",
    "Week Num": "week",
    "Country": "country",
    "Customer": "customer",
    "Deal Name": "deal",
    "Deal Owner": "owner",
    "RFQ Received": "rfq",
    "Closing Date": "close",
    "Tender Category": "category",
    "Description": "desc",
    "Rfx Type": "rfx",
    "Vendor": "vendor",
    "Internal Ref.": "ref",
    "HP QTN REFERENCE": "hpRef",
    "Value": "value",
    "Stage": "stage",
    "Staus": "status",          # sic — header is misspelt in the source workbook
    "Status": "status",         # tolerate the corrected spelling too
    "QTN Date": "qtn",
    "MGE Cost Price": "cost",
    "Clarification Received": "clarRecv",
    "Clarification Completed": "clarDone",
    "Declined Reason": "declined",
    "Bid Feedback": "feedback",
    "PO Received": "po",
}

DATE_KEYS = ["rfq", "close", "qtn", "clarRecv", "clarDone", "po"]
STATUS_ORDER = ["PO Received", "Under Pipeline", "Open Quote", "Bid Lost", "RFQ Declined"]

# Normalise a few status spellings that show up in hand-maintained trackers
STATUS_ALIASES = {
    "open rfq": "Open Quote", "open rfqs": "Open Quote", "open rfq's": "Open Quote",
    "open quote": "Open Quote", "open quotes": "Open Quote", "open quotation": "Open Quote",
    "po received": "PO Received", "under pipeline": "Under Pipeline",
    "bid lost": "Bid Lost", "rfq declined": "RFQ Declined", "declined": "RFQ Declined",
}


def clean_text(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    s = str(v).strip()
    return "" if s.lower() in {"nan", "nat", "none"} else s


def to_number(v):
    """'--', blanks and text become None; everything numeric becomes float."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    if s in {"", "-", ".", "-."}:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def to_iso(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if isinstance(v, (datetime, date, pd.Timestamp)):
        if pd.isna(v):
            return None
        return pd.Timestamp(v).strftime("%Y-%m-%d")
    try:
        ts = pd.to_datetime(str(v), errors="coerce", dayfirst=False)
        return None if pd.isna(ts) else ts.strftime("%Y-%m-%d")
    except Exception:
        return None


def days_between(a, b):
    if not a or not b:
        return None
    return (date.fromisoformat(b) - date.fromisoformat(a)).days


def normalise_status(s):
    key = re.sub(r"\s+", " ", s.strip().lower()).replace("’", "'")
    return STATUS_ALIASES.get(key, s.strip())


DECLINE_BUCKETS = [
    ("No supplier quote",      r"supplier|no quote received|quoted to esnad|for esnad|awarded to hp|pre ?payment|agent"),
    ("Insufficient details",   r"detail|information|not matching|no sec bid"),
    ("Out of scope",           r"out of scope|not in .*sow|scope"),
    ("Portal / deadline",      r"portal|deadline|due date|closure|submit"),
    ("Duplicate / re-floated", r"refloat|already|po received|duplicate"),
    ("Not quoted",             r"not quoted"),
]


def bucket_declined(text: str) -> str:
    t = (text or "").lower()
    if not t:
        return "No reason recorded"
    for label, pat in DECLINE_BUCKETS:
        if re.search(pat, t):
            return label
    return "Other"


def load(path: Path, sheet: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    # keep only rows that have a customer or an RFQ date
    keep = [c for c in ("Customer", "RFQ Received") if c in df.columns]
    if keep:
        df = df.dropna(subset=keep, how="all")
    return df.reset_index(drop=True)


def build_rows(df: pd.DataFrame):
    rows = []
    for i, rec in df.iterrows():
        r = {"id": int(i) + 1}
        for col, key in COLUMNS.items():
            if col not in df.columns:
                continue
            v = rec[col]
            if key in DATE_KEYS:
                r[key] = to_iso(v)
            elif key in ("value", "cost"):
                r[key] = to_number(v)
            elif key in ("sy", "week"):
                n = to_number(v)
                r[key] = int(n) if n is not None else None
            else:
                r[key] = clean_text(v)
        r["status"] = normalise_status(r.get("status", "")) or "Unknown"
        r["deal"] = clean_text(r.get("deal", ""))

        # --- derived time fields (based on RFQ Received) --------------------
        rfq = r.get("rfq")
        if rfq:
            d = date.fromisoformat(rfq)
            r["year"] = d.year
            r["quarter"] = f"Q{(d.month - 1) // 3 + 1}"
            r["month"] = f"{d.year}-{d.month:02d}"
            if r.get("week") is None:
                r["week"] = d.isocalendar()[1]
            if r.get("sy") is None:
                r["sy"] = d.year
        else:
            r["year"] = r.get("sy")
            r["quarter"] = None
            r["month"] = None

        # --- budgetary / indicative quotes (inflate pipeline value) ---------
        blob = f"{r.get('desc','')} {r.get('feedback','')}".lower()
        r["budgetary"] = bool(re.search(r"bdgtry|budgetory|budgetary|budgetry|indicative", blob))

        # --- declined-reason bucket (free text -> analysable category) -----
        r["declinedGroup"] = bucket_declined(r.get("declined", "")) if r["status"] == "RFQ Declined" else ""

        # --- derived cycle-time fields --------------------------------------
        r["daysToClose"] = days_between(rfq, r.get("close"))      # RFQ -> Closing date
        r["daysToQtn"] = days_between(rfq, r.get("qtn"))          # RFQ -> Quotation
        r["daysToPO"] = days_between(rfq, r.get("po"))            # RFQ -> PO received
        r["daysClarification"] = days_between(r.get("clarRecv"), r.get("clarDone"))
        rows.append(r)
    return rows


def build_meta(rows, source: Path, sheet: str):
    def uniq(key):
        return sorted({r[key] for r in rows if r.get(key) not in (None, "")}, key=lambda x: str(x))

    statuses = [s for s in STATUS_ORDER if any(r["status"] == s for r in rows)]
    statuses += [s for s in uniq("status") if s not in statuses]
    dates = [r["rfq"] for r in rows if r.get("rfq")]
    return {
        "title": "HP Valves — RFQ Intelligence",
        "source": source.name,
        "sheet": sheet,
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "currency": "EUR",
        "rowCount": len(rows),
        "dateRange": [min(dates), max(dates)] if dates else [None, None],
        "statuses": statuses,
        "countries": uniq("country"),
        "customers": uniq("customer"),
        "vendors": uniq("vendor"),
        "owners": uniq("owner"),
        "years": uniq("year"),
    }


def inject_html(html_path: Path, payload: str):
    html = html_path.read_text(encoding="utf-8")
    pattern = re.compile(r'(<script id="hp-data" type="application/json">)(.*?)(</script>)', re.S)
    if not pattern.search(html):
        sys.exit(f"ERROR: {html_path} has no <script id=\"hp-data\"> tag to inject into.")
    # protect against a literal </script> inside free-text fields
    safe = payload.replace("</", "<\\/")
    html = pattern.sub(lambda m: m.group(1) + safe + m.group(3), html, count=1)
    html_path.write_text(html, encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description="Convert HP.xlsx into data.json + self-contained index.html")
    ap.add_argument("xlsx", nargs="?", default="HP.xlsx")
    ap.add_argument("--sheet", default="MAIN TABLE")
    ap.add_argument("--json", default="data.json")
    ap.add_argument("--html", default="index.html")
    ap.add_argument("--no-inject", action="store_true", help="write data.json only")
    args = ap.parse_args()

    here = Path(__file__).resolve().parent
    src = Path(args.xlsx)
    if not src.exists():
        src = here / args.xlsx
    if not src.exists():
        sys.exit(f"ERROR: workbook not found: {args.xlsx}")

    df = load(src, args.sheet)
    rows = build_rows(df)
    meta = build_meta(rows, src, args.sheet)
    payload = json.dumps({"meta": meta, "rows": rows}, ensure_ascii=False, separators=(",", ":"))

    json_path = Path(args.json) if Path(args.json).is_absolute() else here / args.json
    json_path.write_text(json.dumps({"meta": meta, "rows": rows}, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"✓ {json_path.name}: {len(rows)} rows, {len(meta['customers'])} customers, "
          f"{len(meta['countries'])} countries, {meta['dateRange'][0]} → {meta['dateRange'][1]}")

    if not args.no_inject:
        html_path = Path(args.html) if Path(args.html).is_absolute() else here / args.html
        if not html_path.exists():
            sys.exit(f"ERROR: {html_path} not found (keep index.html next to convert.py)")
        inject_html(html_path, payload)
        print(f"✓ {html_path.name}: dataset embedded ({len(payload)//1024} KB). Double-click to open.")

    # quick console summary
    by_status = {}
    for r in rows:
        s = by_status.setdefault(r["status"], [0, 0.0])
        s[0] += 1
        s[1] += r["value"] or 0
    for s in meta["statuses"]:
        c, v = by_status.get(s, (0, 0))
        print(f"  {s:<16} {c:>4} RFQs   € {v:,.0f}")


if __name__ == "__main__":
    main()
