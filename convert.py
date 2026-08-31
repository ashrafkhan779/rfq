#!/usr/bin/env python3
"""Convert HP.xlsx into dashboard-ready data.json using only Python stdlib.

Usage:
    python convert.py HP.xlsx data.json
"""
from __future__ import annotations

import json
import math
import re
import sys
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
REL_NS = {"p": "http://schemas.openxmlformats.org/package/2006/relationships"}
DATE_HEADERS = {
    "RFQ Received", "Closing Date", "QTN Date", "Clarification Received",
    "Clarification Completed", "PO Received"
}
STATUS_MAP = {
    "Awaiting PO": "Under Pipeline",
    "Clarification Complete": "Under Pipeline",
    "Clarification Completed": "Under Pipeline",
    "Floated To Supplier/Internal": "Open RFQ's",
    "Quote Received Supplier/Internal": "Open RFQ's",
    "Quoted To Client": "Open RFQ's",
    "PO Received": "PO Received",
    "Bid Lost": "Bid Lost",
    "RFQ Declined": "RFQ Declined",
}


def excel_date(n):
    try:
        n = float(n)
    except (TypeError, ValueError):
        return None
    # Excel's 1900 date system (including the historic leap-year bug).
    dt = datetime(1899, 12, 30) + timedelta(days=n)
    return dt.date().isoformat()


def col_index(cell_ref: str) -> int:
    letters = re.match(r"([A-Z]+)", cell_ref).group(1)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def load_shared_strings(zf):
    name = "xl/sharedStrings.xml"
    if name not in zf.namelist():
        return []
    root = ET.fromstring(zf.read(name))
    strings = []
    for si in root.findall("m:si", NS):
        parts = []
        for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"):
            parts.append(t.text or "")
        strings.append("".join(parts))
    return strings


def resolve_first_sheet(zf):
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    sheet = wb.find("m:sheets/m:sheet", NS)
    if sheet is None:
        raise RuntimeError("Workbook contains no worksheets")
    rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    target = None
    for rel in rels.findall("p:Relationship", REL_NS):
        if rel.attrib.get("Id") == rid:
            target = rel.attrib.get("Target")
            break
    if not target:
        raise RuntimeError("Unable to resolve worksheet relationship")
    target = target.lstrip("/")
    if not target.startswith("xl/"):
        target = "xl/" + target
    return target, sheet.attrib.get("name", "Sheet1")


def cell_value(c, shared):
    typ = c.attrib.get("t")
    v = c.find("m:v", NS)
    if typ == "inlineStr":
        t = c.find("m:is/m:t", NS)
        return t.text if t is not None else ""
    if v is None:
        return None
    raw = v.text
    if typ == "s":
        try:
            return shared[int(raw)]
        except Exception:
            return raw
    if typ == "b":
        return raw == "1"
    if typ in ("str", "e"):
        return raw
    try:
        num = float(raw)
        return int(num) if num.is_integer() else num
    except (TypeError, ValueError):
        return raw


def normalize_text(v):
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    return str(v).strip()


def derive_status(stage, existing):
    stage = normalize_text(stage)
    existing = normalize_text(existing)
    return STATUS_MAP.get(stage, existing or stage or "Unclassified")


def quarter(month):
    return f"Q{((month - 1)//3)+1}"


def month_name(month):
    return datetime(2000, month, 1).strftime("%b")


def parse_date(s):
    if not s:
        return None
    try:
        return datetime.fromisoformat(s).date()
    except Exception:
        return None


def convert(input_path: Path, output_path: Path):
    with zipfile.ZipFile(input_path) as zf:
        shared = load_shared_strings(zf)
        sheet_path, sheet_name = resolve_first_sheet(zf)
        root = ET.fromstring(zf.read(sheet_path))
        parsed_rows = []
        for row in root.findall(".//m:sheetData/m:row", NS):
            cells = {}
            for c in row.findall("m:c", NS):
                ref = c.attrib.get("r", "A1")
                cells[col_index(ref)] = cell_value(c, shared)
            parsed_rows.append(cells)

    if not parsed_rows:
        raise RuntimeError("No rows found")
    max_col = max(parsed_rows[0].keys())
    headers = [normalize_text(parsed_rows[0].get(i)) for i in range(max_col + 1)]
    header_index = {h: i for i, h in enumerate(headers)}

    records = []
    for source_row_num, row in enumerate(parsed_rows[1:], start=2):
        if not any(v not in (None, "") for v in row.values()):
            continue
        rec = {}
        for i, h in enumerate(headers):
            if not h:
                continue
            value = row.get(i)
            if h in DATE_HEADERS and isinstance(value, (int, float)):
                value = excel_date(value)
            elif isinstance(value, float) and not math.isfinite(value):
                value = None
            rec[h] = value

        rec["Status"] = derive_status(rec.get("Stage"), rec.get("Staus"))
        rec.pop("Staus", None)
        rec["Source Row"] = source_row_num

        received = parse_date(rec.get("RFQ Received"))
        closing = parse_date(rec.get("Closing Date"))
        po = parse_date(rec.get("PO Received"))
        if received:
            rec["Year"] = received.year
            rec["Quarter"] = quarter(received.month)
            rec["Month"] = f"{received.month:02d} - {month_name(received.month)}"
            rec["MonthNum"] = received.month
            rec["Week"] = rec.get("Week Num") or int(received.strftime("%V"))
        else:
            rec["Year"] = rec.get("SY") or None
            rec["Quarter"] = None
            rec["Month"] = None
            rec["MonthNum"] = None
            rec["Week"] = rec.get("Week Num") or None

        rec["RFQ Window Days"] = (closing - received).days if received and closing else None
        rec["Order Cycle Days"] = (po - received).days if received and po else None
        try:
            rec["Value"] = float(rec.get("Value")) if rec.get("Value") not in (None, "", "--") else 0.0
        except (TypeError, ValueError):
            rec["Value"] = 0.0
        records.append(rec)

    metadata = {
        "source_file": input_path.name,
        "sheet": sheet_name,
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "currency": "EUR",
        "record_count": len(records),
        "status_logic": STATUS_MAP,
        "notes": [
            "Value is displayed as EUR per dashboard requirement.",
            "Days to close is calculated in the browser from Closing Date against the viewer's current date.",
            "RFQ Window Days is Closing Date minus RFQ Received.",
            "Customer Last Order uses the latest PO Received date for that customer."
        ]
    }
    output_path.write_text(json.dumps({"metadata": metadata, "records": records}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {len(records)} records to {output_path}")


if __name__ == "__main__":
    src = Path(sys.argv[1] if len(sys.argv) > 1 else "HP.xlsx")
    dst = Path(sys.argv[2] if len(sys.argv) > 2 else "data.json")
    convert(src, dst)
