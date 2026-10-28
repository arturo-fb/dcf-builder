#!/usr/bin/env python3
"""Fill the DCF template's blue input cells from a JSON file, keeping all template formatting.

Usage:  python3 fill_dcf.py TEMPLATE.xlsx inputs.json OUTPUT.xlsx

inputs.json:
{
  "Quarterly!B3": "Airbnb, Inc.",                      # plain value (number or text)
  "Quarterly!D7": {"date": "2025-12-31"},              # date
  "DCF!M7": {"text": "Q2'26 shareholder letter", "url": "https://..."},   # text with hyperlink
  "DCF!C13": {"text": "Closing price on 2026-10-01 (IBKR)"}               # note text (style kept)
}
Only writes the cells listed. Refuses to overwrite a template formula unless the key ends with '!'
(e.g. "DCF!H19!" to force) – formulas are the model, inputs are the only thing that changes.
"""
import datetime as dt
import json
import sys

from openpyxl import load_workbook
from openpyxl.styles import Font


def main(tpl, js, out):
    wb = load_workbook(tpl)
    data = json.load(open(js))
    for key, val in data.items():
        force = key.endswith("!")
        key = key.rstrip("!")
        sheet, ref = key.rsplit("!", 1)
        c = wb[sheet][ref]
        if isinstance(c.value, str) and c.value.startswith("=") and not force:
            raise SystemExit(f"{key} holds a formula ({c.value[:40]}…) – not overwriting")
        default_font = c.font is None or c.font.name in (None, "Calibri")
        if isinstance(val, dict):
            if default_font and "date" not in val:
                # unstyled cell (e.g. a Source / notes column): note style, or link style if it has a URL
                c.font = Font(name="Arial", size=9, italic=not val.get("url"),
                              color="FF0563C1" if val.get("url") else "FF595959",
                              underline="single" if val.get("url") else None)
            if "date" in val:
                c.value = dt.datetime.strptime(val["date"], "%Y-%m-%d")
            else:
                c.value = val.get("text")
            if val.get("url"):
                c.hyperlink = val["url"]
        else:
            c.value = val
    wb.save(out)
    print(f"filled {len(data)} cells -> {out}")


if __name__ == "__main__":
    main(*sys.argv[1:4])
