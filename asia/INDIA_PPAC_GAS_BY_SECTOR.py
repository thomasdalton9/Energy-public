"""
India natural gas consumption by sector, monthly, from PPAC (Petroleum Planning & Analysis Cell, Ministry of
Petroleum and Natural Gas) - the raw government source. Sectors as PPAC reports them: fertiliser, power, city gas
distribution (CGD), refineries, petrochemicals, sponge iron, other industry, LPG shrinkage, others.

STATUS: UNVERIFIED. The Claude sandbox blocks ppac.gov.in, so the layout of the PPAC files has never been seen by
this script's author. It therefore does not hard-code file URLs or cell positions: it finds Excel/CSV links on the
PPAC consumption pages whose link text or file name mentions consumption, then reads each sheet tolerantly (a
header row holding month labels, rows whose first cell names a sector). If no sheet yields a month x sector
table the script exits non-zero with diagnostics and does NOT overwrite the workbook. Run
the manual workflow discovery_archive/workflows/world_gas_sector_probe.yml first (it fetches these pages from Actions and lists the
links) and adjust PAGES / SECTOR_KEYWORDS if PPAC's layout differs.

Unit: million standard cubic metres per month (MMSCM), as PPAC publishes (some PPAC tables are in MMSCMD, a daily
rate: these are multiplied by days in month when the sheet title says "MMSCMD").

Incremental: PPAC publishes whole files, so each file is downloaded only when its Last-Modified / Content-Length
differs from the value recorded on the "Source files" tab (CLAUDE.md); monthly values already saved are kept and
files only add or revise months.

    python3 asia/INDIA_PPAC_GAS_BY_SECTOR.py --out "output/Data and Chart Outputs/india_ppac_gas_by_sector.xlsx"
"""
import argparse
import io
import os
import re
import sys
from urllib.parse import urljoin

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root
import xlsx_notes  # noqa: E402

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 180)
PAGES = ["https://ppac.gov.in/natural-gas/consumption", "https://ppac.gov.in/natural-gas/consumption-of-natural-gas",
         "https://ppac.gov.in/natural-gas"]
LINK_RE = re.compile(r'href=["\']([^"\']+\.(?:xlsx?|csv))["\'][^>]*>(.*?)</a>', re.I | re.S)
SECTOR_KEYWORDS = [   # (regex on the row label, sector name)
    (r"fertili[sz]er", "Fertiliser"), (r"power|electric", "Power"), (r"\bcgd\b|city gas|cng|png", "City gas distribution"),
    (r"refiner", "Refineries"), (r"petrochem", "Petrochemicals"), (r"sponge", "Sponge iron"),
    (r"lpg", "LPG shrinkage"), (r"industr|manufactur", "Industrial and manufacturing"),
    (r"^total", "Total"), (r"other", "Others"),
]
RAW_COLS = ["month", "sector", "value", "unit", "source_file"]
SECTIONS = {"UNITS", "SOURCE", "CAVEATS", "SOURCE FILES"}


def out(*a):
    print(*a, flush=True)


def find_files(session):
    found = {}
    for page in PAGES:
        try:
            r = session.get(page, headers=H, timeout=T)
        except requests.RequestException as exc:
            out(f"  page {page}: {type(exc).__name__}")
            continue
        out(f"  page {page}: HTTP {r.status_code}, {len(r.text)} chars")
        if r.status_code != 200:
            continue
        for href, text in LINK_RE.findall(r.text):
            label = re.sub(r"<[^>]+>|\s+", " ", text).strip()
            if re.search(r"consum|sector|demand", href + " " + label, re.I):
                found[urljoin(page, href)] = label
    return found


MONTH_RE = re.compile(r"^([A-Za-z]{3,9})[-\s/'.,]*(\d{2}|\d{4})$")
ISO_RE = re.compile(r"^(\d{4})-(\d{2})(?:-\d{2})?")


def _month(cell):
    """Header cell -> first-of-month Timestamp, or None. Accepts datetimes, 'Apr-24', 'April 2024', '2024-04'."""
    if isinstance(cell, (pd.Timestamp, __import__("datetime").datetime)):
        return pd.Timestamp(cell).replace(day=1)
    s = str(cell).strip()
    m = MONTH_RE.match(s)
    if m:
        year = int(m.group(2))
        year += 2000 if year < 100 else 0
        t = pd.to_datetime(f"{m.group(1)[:3]} {year}", format="%b %Y", errors="coerce")
        return None if pd.isna(t) else t
    m = ISO_RE.match(s)
    if m and 1 <= int(m.group(2)) <= 12:
        return pd.Timestamp(int(m.group(1)), int(m.group(2)), 1)
    return None


def parse_sheet(df, title, source):
    """Find a header row with >= 3 month labels, then sector rows below it -> long frame."""
    daily = bool(re.search(r"mmscmd", title, re.I))
    rows = []
    for hi in range(min(len(df), 40)):
        months = {j: _month(v) for j, v in enumerate(df.iloc[hi]) if _month(v) is not None}
        if len(months) < 3:
            continue
        for ri in range(hi + 1, len(df)):
            label = str(df.iat[ri, 0]).strip() if df.shape[1] else ""
            if not label or label.lower() == "nan":
                label = str(df.iat[ri, 1]).strip() if df.shape[1] > 1 else ""
            sector = next((name for rx, name in SECTOR_KEYWORDS if re.search(rx, label.lower())), None)
            if sector is None:
                continue
            for j, m in months.items():
                v = pd.to_numeric(df.iat[ri, j], errors="coerce")
                if pd.notna(v):
                    rows.append({"month": m, "sector": sector, "value": float(v) * (m.days_in_month if daily else 1),
                                 "unit": "MMSCM", "source_file": source})
        if rows:
            break
    return pd.DataFrame(rows, columns=RAW_COLS)


def parse_file(content, name):
    if name.lower().endswith(".csv"):
        return parse_sheet(pd.read_csv(io.BytesIO(content), header=None), name, name)
    parts = []
    for sheet, df in pd.read_excel(io.BytesIO(content), sheet_name=None, header=None).items():
        got = parse_sheet(df, f"{name} {sheet}", name)
        out(f"    sheet {sheet!r}: {len(got)} cells")
        parts.append(got)
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=RAW_COLS)


def load_state(path):
    if not os.path.exists(path):
        return pd.DataFrame(columns=RAW_COLS), pd.DataFrame(columns=["url", "last_modified", "length"])
    x = pd.ExcelFile(path)
    raw = pd.read_excel(x, "Raw") if "Raw" in x.sheet_names else pd.DataFrame(columns=RAW_COLS)
    src = pd.read_excel(x, "Source files") if "Source files" in x.sheet_names else pd.DataFrame(
        columns=["url", "last_modified", "length"])
    return raw, src


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join("output", "Data and Chart Outputs", "india_ppac_gas_by_sector.xlsx"))
    args = ap.parse_args()
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    raw, src = load_state(args.out)
    seen = {r.url: (str(r.last_modified), str(r.length)) for r in src.itertuples()}
    s = requests.Session()
    files = find_files(s)
    out(f"{len(files)} candidate file link(s)")
    for u, lab in files.items():
        out(f"  {u}  [{lab}]")
    new_parts, new_src = [], dict(seen)
    for url in files:
        head = s.head(url, headers=H, timeout=T, allow_redirects=True)
        sig = (head.headers.get("Last-Modified", ""), head.headers.get("Content-Length", ""))
        if seen.get(url) == sig and sig != ("", ""):
            out(f"  unchanged, skipped: {url}")
            continue
        r = s.get(url, headers=H, timeout=T)
        r.raise_for_status()
        got = parse_file(r.content, url.rsplit("/", 1)[-1])
        out(f"  {url}: {len(got)} month x sector cells")
        if len(got):
            new_parts.append(got)
            new_src[url] = sig
    if not new_parts and raw.empty:
        raise SystemExit("No month x sector table could be read from any PPAC file - see log; nothing written")
    raw["month"] = pd.to_datetime(raw["month"]) if len(raw) else raw["month"]
    allr = pd.concat([raw] + new_parts, ignore_index=True).drop_duplicates(["month", "sector"], keep="last")
    allr = allr.sort_values(["month", "sector"]).reset_index(drop=True)
    wide = allr.pivot_table(index="month", columns="sector", values="value", aggfunc="first")
    wide.index.name = "Month"
    src_df = pd.DataFrame([{"url": u, "last_modified": a, "length": b} for u, (a, b) in new_src.items()])
    notes = [
        "UNITS", "MMSCM = million standard cubic metres per month, as published by PPAC (MMSCMD tables converted "
        "with days in month).", "",
        "SOURCE", "PPAC, Ministry of Petroleum and Natural Gas, India: https://ppac.gov.in/natural-gas/consumption",
        "Sector consumption of natural gas (fertiliser, power, city gas distribution, refineries, petrochemicals, "
        "sponge iron, other).", "",
        "CAVEATS", "Layout-tolerant parser, first built without access to the live files (UNVERIFIED until the "
        "first Actions run). PPAC revises recent months; later files overwrite earlier months. Includes domestic gas "
        "and imported LNG; PPAC's 'Total' is its own sum, shown as a separate column.", "",
        "SOURCE FILES", "Tab 'Source files' lists each downloaded file with Last-Modified and size: a file is only "
        "downloaded again when these change.",
        f"Updated: {pd.Timestamp.now('UTC'):%Y-%m-%d %H:%M} UTC.",
    ]
    xlsx_notes.write_workbook(args.out, {"Monthly by sector": wide, "Raw": allr.set_index("month"),
                                         "Source files": src_df.set_index("url")}, notes, SECTIONS)
    out(f"wrote {args.out}: {len(allr)} cells, {wide.index.min():%b/%y} to {wide.index.max():%b/%y}")


if __name__ == "__main__":
    main()
