#!/usr/bin/env python3
"""Build the Suitcase Rex web database from data/Suitcase_Rex.xlsx.

Reads Sheet1 of the spreadsheet, cleans each row, and writes:
  data/tracks.json  - cleaned records
  data/tracks.csv   - the same records as a flat CSV
  index.html        - src/index.template.html with the records embedded

Usage:  python3 scripts/build.py [--fragment OUT.html]
Requires: openpyxl
"""
import csv
import json
import re
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "data" / "Suitcase_Rex.xlsx"
TEMPLATE = ROOT / "src" / "index.template.html"

# Suitcase box sets: 4 discs each, numbered 1-16 across the series in the sheet.
SUITCASES = {
    1: {"title": "Failed Experiments and Trashed Aircraft", "released": 2000},
    2: {"title": "American Magnetic", "released": 2005},
    3: {"title": "Up We Go Now", "released": 2009},
    4: {"title": "Captain Kangaroo Won the War", "released": 2015},
}


def clean_text(value):
    if value is None:
        return None
    text = str(value).replace("\xa0", " ")
    # Excel-escaped Windows-1252 apostrophe (0x92)
    text = text.replace("_x0092_", "’")
    text = re.sub(r"\s+", " ", text).strip()
    return text or None


def parse_year(raw):
    """Return (year, year_label, approximate)."""
    if raw is None:
        return None, None, False
    if isinstance(raw, (int, float)):
        y = int(raw)
        return y, str(y), False
    text = clean_text(raw) or ""
    m = re.fullmatch(r"(\d{4})\?", text)
    if m:
        y = int(m.group(1))
        return y, f"{y}?", True
    m = re.fullmatch(r"(\d{2})-'(\d{2})\?", text)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        return 1900 + a, f"19{a:02d}–{b:02d}?", True
    # "?" and stray values like "e" mean the year is unknown
    return None, None, False


def parse_location(raw):
    text = clean_text(raw) or ""
    m = re.fullmatch(r"Disc (\d+), track (\d+)", text)
    if m:
        disc, track = int(m.group(1)), int(m.group(2))
        suitcase = (disc - 1) // 4 + 1
        return {
            "collection": f"Suitcase {suitcase}",
            "suitcase": suitcase,
            "disc": disc,
            "suitcase_disc": (disc - 1) % 4 + 1,
            "track_no": track,
        }
    m = re.fullmatch(r"Briefcase (\d+)", text)
    if m:
        return {
            "collection": f"Briefcase {int(m.group(1))}",
            "briefcase": int(m.group(1)),
        }
    return None


def split_writers(credit):
    if not credit:
        return []
    return [w.strip() for w in credit.split("-") if w.strip()]


def load_records():
    ws = openpyxl.load_workbook(XLSX, read_only=True)["Sheet1"]
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        band, title, year_raw, notes, credit, loc_raw = (list(row) + [None] * 6)[:6]
        loc = parse_location(loc_raw)
        if loc is None:
            continue  # blank rows and section labels such as "remakes"
        year, year_label, approx = parse_year(year_raw)
        credit = clean_text(credit)
        rec = {
            "band": clean_text(band),
            "title": clean_text(title),
            "year": year,
            "year_label": year_label,
            "year_approx": approx,
            "notes": clean_text(notes),
            "credit": credit,
            "writers": split_writers(credit),
            "source_location": clean_text(loc_raw),
            **loc,
        }
        records.append(rec)

    def order(r):
        return (
            r.get("suitcase") or 99,
            r.get("disc") or 0,
            r.get("track_no") or 0,
            r.get("briefcase") or 0,
            r["title"] or "",
        )

    records.sort(key=order)
    for i, r in enumerate(records, 1):
        r["id"] = i
    return records


def main():
    records = load_records()
    payload = {"suitcases": SUITCASES, "tracks": records}

    (ROOT / "data" / "tracks.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=1) + "\n", encoding="utf-8"
    )

    fields = [
        "id", "collection", "disc", "suitcase_disc", "track_no", "band", "title",
        "year", "year_label", "year_approx", "credit", "notes", "source_location",
    ]
    with open(ROOT / "data" / "tracks.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in records:
            w.writerow(r)

    embedded = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    embedded = embedded.replace("</", "<\\/")
    page = TEMPLATE.read_text(encoding="utf-8").replace("/*__DATA__*/null", embedded)
    head, body = page.split('<div class="wrap">', 1)
    html = (
        '<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
        + head.strip() + '\n</head>\n<body>\n<div class="wrap">' + body.rstrip() + "\n</body>\n</html>\n"
    )
    (ROOT / "index.html").write_text(html, encoding="utf-8")
    if len(sys.argv) > 2 and sys.argv[1] == "--fragment":
        Path(sys.argv[2]).write_text(page, encoding="utf-8")

    dated = sum(1 for r in records if r["year"] is not None)
    print(f"{len(records)} tracks ({dated} dated) -> index.html, data/tracks.json, data/tracks.csv")


if __name__ == "__main__":
    main()
