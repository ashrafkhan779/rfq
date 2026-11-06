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
    "End-Client": "customer",     # end-client is the client dimension of the dashboard
    "End Client": "customer",
    "Customer": "trader",         # trading partner / contractual customer
    "Deal Name": "deal",
    "Deal Owner": "owner",
    "RFQ Received": "rfq",
    "Closing Date": "close",
    "Tender Category": "category",
    "BUDGETORY": "budgetaryFlag",   # Yes / No column added Sept 2026 (spelling as in the workbook)
    "BUDGETARY": "budgetaryFlag",
    "Budgetary": "budgetaryFlag",
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
# Status values are kept exactly as typed in the workbook (whitespace collapsed; case variants merged to the first
# spelling seen). Each is classified into a ROLE by keyword — the dashboard's logic runs on the role, not the label.
ROLE_ORDER = ["po", "pipe", "open", "lost", "dec", "other"]


def role_of(status: str) -> str:
    k = (status or "").lower()
    if re.search(r"\bpo\b|purchase order|awarded|won", k): return "po"
    if re.search(r"pipeline|pipe", k): return "pipe"
    if re.search(r"open|quoted|quote", k): return "open"
    if re.search(r"lost|reject", k): return "lost"
    if re.search(r"declin|no bid|not quoted|regret", k): return "dec"
    return "other"


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


COUNTRY_ALIASES = {
    "qatar": "QTR", "saudi arabia": "KSA", "saudi": "KSA", "ksa": "KSA", "uae": "UAE", "united arab emirates": "UAE", "uaw": "UAE",
    "pakistan": "PAK", "jordan": "JRDN", "bangladesh": "BNG", "bahrain": "BAH", "algeria": "ALG", "syria": "SIR",
    "turkmenistan": "TURK", "turkmen": "TURK", "tkm": "TURK", "oman": "OMN", "kuwait": "KWT", "egypt": "EGY", "iraq": "IRQ", "india": "IND",
}


def normalise_country(c: str) -> str:
    key = re.sub(r"\s+", " ", c.strip().lower())
    return COUNTRY_ALIASES.get(key, c.strip().upper() if len(c.strip()) <= 4 else c.strip())


def derive_status(stage: str) -> str:
    """Mirrors the Staus formula in the workbook (plus clarification stages the formula misses)."""
    k = (stage or "").strip().lower()
    if not k:
        return ""
    if k == "po received": return "PO Received"
    if k == "bid lost": return "Bid Lost"
    if k == "rfq declined": return "RFQ Declined"
    if re.search(r"awaiting po|clarification", k): return "Under Pipeline"
    if re.search(r"quote received supplier|quoted to client", k): return "Open Quotes"
    if re.search(r"rfq acknowledged|rfq received|floated to supplier", k): return "Open RFQ"
    return ""


def normalise_status(s, stage=""):
    return re.sub(r"\s+", " ", (s or "").strip()) or derive_status(stage) or "Unknown"


def merge_status_spellings(rows):
    """Case/spacing variants of the same status collapse to the most common spelling in the file."""
    from collections import Counter, defaultdict
    counts = defaultdict(Counter)
    for r in rows:
        counts[r["status"].lower().replace("’", "'")][r["status"]] += 1
    canon = {k: c.most_common(1)[0][0] for k, c in counts.items()}
    for r in rows:
        r["status"] = canon[r["status"].lower().replace("’", "'")]
        r["role"] = role_of(r["status"])


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


def hdr_key(h) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(h or "").lower())


_KEY_BY_NORM = {hdr_key(c): k for c, k in COLUMNS.items()}


def match_header(h, headers=()):
    k = hdr_key(h)
    if not k:
        return None
    if k == "customer":            # 'Customer' = client unless an End-Client column exists
        return "trader" if any(hdr_key(x) == "endclient" for x in headers) else "customer"
    if k == "endclient":
        return "customer"
    if k in _KEY_BY_NORM:
        return _KEY_BY_NORM[k]
    if k.startswith("budget"):            # BUDGETORY / Budgetary Quote Y/N / Budgetary?
        return "budgetaryFlag"
    if re.fullmatch(r"sta?tus", k):
        return "status"
    if k in ("weeknum", "week", "wk"):
        return "week"
    if k in ("internalref", "internalreference"):
        return "ref"
    if k in ("hpqtnref", "hpqtnreference", "hpquotationref"):
        return "hpRef"
    if k in ("mgecostprice", "costprice"):
        return "cost"
    return None


def load(path: Path, sheet: str) -> pd.DataFrame:
    """Read the sheet, locating the header row automatically (it must contain Customer + RFQ Received)
    and renaming columns to their canonical workbook names so that renamed/extra columns don't break anything."""
    xl = pd.ExcelFile(path)
    if sheet not in xl.sheet_names:
        alt = next((n for n in xl.sheet_names if n.strip().upper() == sheet.strip().upper()), xl.sheet_names[0])
        print(f"  sheet '{sheet}' not found — using '{alt}'")
        sheet = alt
    probe = pd.read_excel(path, sheet_name=sheet, header=None, nrows=15)
    hi = next((i for i, row in probe.iterrows()
               if any(hdr_key(v) in ("customer", "endclient") for v in row) and any(hdr_key(v) == "rfqreceived" for v in row)), None)
    if hi is None:
        sys.exit("ERROR: could not find the header row (needs 'Customer' and 'RFQ Received').")
    df = pd.read_excel(path, sheet_name=sheet, header=hi)
    canon_by_key = {}
    for c, k in COLUMNS.items():
        canon_by_key.setdefault(k, c)
    rename, extra, used = {}, [], set()
    for c in df.columns:
        k = match_header(c, list(df.columns))
        if k and k not in used:
            used.add(k)
            rename[c] = canon_by_key[k]
        elif str(c).strip() and not str(c).startswith("Unnamed"):
            extra.append(str(c).strip())
    df = df.rename(columns=rename)
    df = df.loc[:, ~df.columns.duplicated()]
    df.attrs["extra_cols"] = extra
    if extra:
        print(f"  extra columns kept: {', '.join(extra)}")
    # keep only rows that have a customer or an RFQ date
    keep = [c for c in ("Customer", "RFQ Received") if c in df.columns]
    if keep:
        df = df.dropna(subset=keep, how="all")
    return df.reset_index(drop=True)


def build_rows(df: pd.DataFrame):
    rows = []
    has_flag = any(c in df.columns for c, k in COLUMNS.items() if k == "budgetaryFlag")
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
        if not r.get("customer") and r.get("trader"):
            r["customer"] = r["trader"]       # End-Client blank → fall back to the trading partner
        r["status"] = normalise_status(r.get("status", ""), r.get("stage", ""))
        r["role"] = role_of(r["status"])
        r["country"] = normalise_country(r.get("country", ""))
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
        # Rule: the BUDGETORY column (Yes/No) decides. Only when the workbook has no such column
        # do we fall back to keyword detection in the description / feedback text.
        flag = clean_text(r.pop("budgetaryFlag", "")).lower()
        if has_flag:
            r["budgetary"] = flag in {"yes", "y", "true", "1", "x", "✓"}
        else:
            blob = f"{r.get('desc','')} {r.get('feedback','')}".lower()
            r["budgetary"] = bool(re.search(r"bdgtry|budgetory|budgetary|budgetry|indicative", blob))

        # --- declined-reason bucket (free text -> analysable category) -----
        r["declinedGroup"] = bucket_declined(r.get("declined", "")) if r["role"] == "dec" else ""

        # --- derived cycle-time fields --------------------------------------
        r["daysToClose"] = days_between(rfq, r.get("close"))      # RFQ -> Closing date
        r["daysToQtn"] = days_between(rfq, r.get("qtn"))          # RFQ -> Quotation
        r["daysToPO"] = days_between(rfq, r.get("po"))            # RFQ -> PO received
        r["daysClarification"] = days_between(r.get("clarRecv"), r.get("clarDone"))
        extra = {}
        for c in df.attrs.get("extra_cols", []):
            if c not in df.columns:
                continue
            v = rec[c]
            if isinstance(v, (datetime, date, pd.Timestamp)):
                v = to_iso(v)
            elif isinstance(v, float) and v.is_integer():
                v = str(int(v))
            extra[c] = clean_text(v)
        extra = {k: v for k, v in extra.items() if v not in ("", None)}
        if extra:
            r["extra"] = extra
        rows.append(r)
    merge_status_spellings(rows)
    return rows


def build_meta(rows, source: Path, sheet: str):
    def uniq(key):
        return sorted({r[key] for r in rows if r.get(key) not in (None, "")}, key=lambda x: str(x))

    statuses = sorted(uniq("status"), key=lambda s: (ROLE_ORDER.index(role_of(s)), s))
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
    ap.add_argument("xlsx", nargs="?", default=None, help="source workbook (default: HP-FIN.xlsx, else HP.xlsx)")
    ap.add_argument("--sheet", default="MAIN TABLE")
    ap.add_argument("--json", default="data.json")
    ap.add_argument("--html", default="index.html")
    ap.add_argument("--no-inject", action="store_true", help="write data.json only")
    args = ap.parse_args()

    here = Path(__file__).resolve().parent
    if args.xlsx:
        src = Path(args.xlsx)
        if not src.exists():
            src = here / args.xlsx
    else:
        src = next((here / n for n in ("HP-FIN.xlsx", "HP.xlsx") if (here / n).exists()), here / "HP-FIN.xlsx")
    if not src.exists():
        sys.exit(f"ERROR: workbook not found: {src}")

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
        print(f"  {s:<18} [{role_of(s):<5}] {c:>4} RFQs   € {v:,.0f}")


if __name__ == "__main__":
    main()
