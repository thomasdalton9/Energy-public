"""
Jordan power ANNUAL statistics from the National Electric Power Company (NEPCO) annual reports and the Ministry of
Energy and Mineral Resources (MEMR) annual reports -> output/Data and Chart Outputs/jordan_power_annual.xlsx

SOURCES (text PDFs read with pdfplumber; probe discovery_archive/middle_east/JORDAN_ANNUAL_REPORTS_PROBE.py, text
dump discovery_archive/results/jordan_annual/, Oct 2026):
  NEPCO English  https://www.nepco.com.jo/store/docs/web/<yyyy>_en.pdf (2003-2018), /store/DOCS/web/<yyyy>_EN.pdf
                 (2019 on); list page https://www.nepco.com.jo/en/AnnualReports.aspx. 2007, 2009 and 2016 are
                 image-only PDFs (no text layer): nothing is read from them, their years come from the next reports'
                 multi-year tables.
  NEPCO Arabic   /store/DOCS/web/<yyyy>_AR.pdf - read ONLY for a year whose English report does not exist (2025: the
                 Arabic report was posted 12 Aug 2026, the English one is not out). Arabic text comes out in visual
                 order (numbers left to right = newest year first, labels reversed); each Arabic table is ACCEPTED
                 only if its values for the years it shares with the English reports agree (>= 70% of at least 6
                 shared values within 1% / 0.15), else it is left out and the year stays blank (note on 'Releases').
  MEMR English   https://www.memr.gov.jo/En/List/Annual_Reports (links /ebv4.0/root_storage/en/eb_list_page/...,
                 2015-2024): prose only. The sentence giving natural gas consumed for electricity generation (MMcf
                 per year and/or MMcf per day) and the Risha field's gas where stated; kept verbatim on 'MEMR
                 sentences'. Years without such a sentence (2015, 2018) are blank.

WHAT IS READ (as published, no unit conversion except MMcf/d / 1000 = Bcf/d; '--' in a NEPCO table = 0):
  * Energy purchased by fuel, GWh (NEPCO table 'Electrical Energy Purchased by type of Fuel', reports 2018 on, 4-5
    years each: heavy fuel, natural gas, diesel, oil shale, renewables, others, imported, total) and wind / solar /
    hydro from 'Purchased by type of Generation'. This is NEPCO's PURCHASE (sent-out) basis = the interconnected
    system, 2015 on.
  * Production by fuel, GWh, OLDER BASIS (reports 2008-2017, 'Electrical Energy Production by Type of Fuel in
    Jordan'): national GROSS production incl. industrial self-generation, imports excluded, 2007-2017. Kept in its own
    columns (Prod_*) and its own chart, never mixed with the purchase basis.
  * Fuel consumption for electricity generation, thousand toe (T.T.O.E): natural gas, oil shale, heavy fuel, diesel,
    total (2005 on). Up to the 2017 report the table also has the industrial self-generators' fuel; the 'Total' is
    the table's total.
  * Peak and minimum load of the interconnected system, MW with dates (table 'Peak load for Interconnected System':
    morning/evening peaks with dates in the 2020 reports on, without dates 2013-2019) and the report-year peak and
    its month from the report's prose (2003 on).
  * Imports (NEPCO purchases from Egypt, Syria ...) and exports by counterparty (NEPCO sales table: Egypt, Syria,
    Jerusalem District Electricity Co. (Jericho), Iraq at Trabeel border, Iraq General Company for Electricity
    Transmission), GWh.
  * Key figures (table 1): peak MW, purchased and sold GWh, for the latest two years.

PRECEDENCE: the NEWEST report that publishes a year/item wins (every value of every report is on 'By report', the
report and page used on 'Sources used'). 'Checks': fuels add up to the total, wind + solar + hydro = renewables,
export parts = exported total, MEMR MMcf / days = MMcf/d, MEMR MMcf per NEPCO gas ktoe, values revised between reports,
and NEPCO's annual peak against the daily NEPCO Max Load Archive workbook (jordan_peak_load_daily.xlsx).

SCHEDULE (owner: the data are annual): workflow_dispatch plus a light check on 1 Sep and 1 Dec
(.github/workflows/jordan_power_annual.yml). A check sends a HEAD request for each known report of the last three
years and tries the reports not yet found (next year's English and Arabic NEPCO report, missing English reports,
MEMR list page), downloads only new or changed files (Last-Modified, then SHA-256 on 'Releases') and exits without
rewriting anything when nothing is new. NEPCO cuts a runner off after about 65 requests, so a full run (about 25
NEPCO downloads) stays well under that.

Usage: python3 middle_east/JORDAN_NEPCO_ANNUAL.py [--out ...] [--reparse] [--recheck-all]
       [--text-dump discovery_archive/results/jordan_annual/report_text.json.gz  (no downloads, for testing)]
"""

import argparse
import gzip
import hashlib
import io
import json
import os
import re
import sys
import time
from datetime import datetime, timezone

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

requests.packages.urllib3.disable_warnings()
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs")
DEFAULT_OUT = os.path.join(DATA_DIR, "jordan_power_annual.xlsx")
DAILY_PEAK = os.path.join(DATA_DIR, "jordan_peak_load_daily.xlsx")
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "en-US,en;q=0.9"}
NEPCO = "https://www.nepco.com.jo"
MEMR = "https://www.memr.gov.jo"
MEMR_LIST = MEMR + "/En/List/Annual_Reports"
FIRST_NEPCO, FIRST_MEMR = 2003, 2015
PAUSE_S = 2

MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
          "November", "December"]

# Annual sheet columns (item -> unit); the order is the sheet's column order
ITEMS = {
    # NEPCO purchases (sent-out) by fuel, reports 2018 on
    "Purch_Gas_GWh": "GWh", "Purch_Oil_shale_GWh": "GWh", "Purch_Heavy_fuel_GWh": "GWh", "Purch_Diesel_GWh": "GWh",
    "Purch_Renewables_GWh": "GWh", "Purch_Wind_GWh": "GWh", "Purch_Solar_GWh": "GWh", "Purch_Hydro_GWh": "GWh",
    "Purch_Imports_GWh": "GWh", "Purch_Other_GWh": "GWh", "Purch_Total_GWh": "GWh",
    # older basis: national production by fuel (reports 2008-2017)
    "Prod_Heavy_fuel_GWh": "GWh", "Prod_Gas_GWh": "GWh", "Prod_Diesel_GWh": "GWh", "Prod_Other_GWh": "GWh",
    "Prod_Renewables_GWh": "GWh", "Prod_Total_GWh": "GWh",
    # fuel for power
    "Fuel_Gas_ktoe": "ktoe", "Fuel_Oil_shale_ktoe": "ktoe", "Fuel_Heavy_fuel_ktoe": "ktoe", "Fuel_Diesel_ktoe": "ktoe",
    "Fuel_Total_ktoe": "ktoe",
    # MEMR gas for power
    "MEMR_Gas_power_MMcf": "MMcf", "MEMR_Gas_power_MMcf_d": "MMcf/d", "MEMR_Gas_power_Bcf_d": "Bcf/d",
    "MEMR_Risha_MMcf": "MMcf", "MEMR_Risha_MMcf_d": "MMcf/d",
    # peak / minimum load
    "Peak_MW": "MW", "Peak_date": "date", "Peak_morning_MW": "MW", "Peak_morning_date": "date",
    "Peak_evening_MW": "MW", "Peak_evening_date": "date", "Min_MW": "MW", "Min_date": "date", "Peak_month_prose": "",
    # trade
    "Import_Egypt_GWh": "GWh", "Import_Syria_GWh": "GWh", "Import_Other_GWh": "GWh", "Import_Total_GWh": "GWh",
    "Export_Egypt_GWh": "GWh", "Export_Syria_GWh": "GWh", "Export_Palestine_Jericho_GWh": "GWh",
    "Export_Iraq_Trabeel_GWh": "GWh", "Export_Iraq_transmission_GWh": "GWh", "Export_Other_GWh": "GWh",
    "Export_Total_GWh": "GWh",
    # key figures (table 1)
    "Key_Peak_MW": "MW", "Key_Purchased_GWh": "GWh", "Key_Sold_GWh": "GWh",
}


def log(msg):
    print(msg, flush=True)


# ------------------------------------------------------------------ tokens

DASH = {"--", "-", "_", "–", "—", "---", "−"}
NUMTOK = re.compile(r"^[\(\)]?-?[\d,]*\d(?:\.\d+)?[\(\)%]?$")
GROWTHTOK = re.compile(r"^(>100|100>|<100|N/A|n/a)$")
YEARTOK = re.compile(r"^((?:19|20)\d\d)\*{0,3}$")


def is_val(tok):
    return tok in DASH or bool(NUMTOK.match(tok)) or bool(GROWTHTOK.match(tok))


def num(tok):
    """Value token -> float ('--' = 0.0); growth-like tokens -> None."""
    if tok in DASH:
        return 0.0
    if GROWTHTOK.match(tok):
        return None
    t = tok.replace(",", "").strip("()%")
    try:
        return float(t)
    except ValueError:
        return None


def clean_lines(text):
    return [re.sub(r"\s+", " ", l.replace("ـ", "")).strip() for l in (text or "").split("\n") if l.strip()]


def header_years(line):
    """Years in a header line ('2020 2021 2022 2023 2024* (%)'); tokens with '/' (growth headers) are ignored."""
    ys = [int(m.group(1)) for t in line.split() for m in [YEARTOK.match(t)] if m]
    return ys if len(ys) >= 2 and len(set(ys)) == len(ys) and max(ys) - min(ys) == len(ys) - 1 else None


def split_row(line):
    """'Natural Gas 15768 15172 15515.9 (2.0)' -> ('Natural Gas', ['15768', ..., '(2.0)']) using the trailing run."""
    toks = line.split()
    k = len(toks)
    while k > 0 and is_val(toks[k - 1]):
        k -= 1
    return " ".join(toks[:k]), toks[k:]


def year_table(lines, start, stop_re=None):
    """Year-column table: header with the years, then 'label v1 .. vn [growth]' rows. Returns (years, rows) with
    rows = [(label, [values per year])]. A row without a label takes the text-only line before it (plus the one
    after it, for labels printed on two lines around the numbers)."""
    years, rows = None, []
    i = start
    while i < len(lines) and i < start + 6 and years is None:
        years = header_years(lines[i])
        i += 1
    if not years:
        return None, []
    n = len(years)
    years = sorted(years)            # EN reports print the columns oldest first (2010 sales header is reversed)
    prev_text = None
    while i < len(lines):
        line = lines[i]
        if stop_re and re.search(stop_re, line, re.I):
            break
        label, vals = split_row(line)
        if len(vals) >= n and not header_years(line):
            v = [num(t) for t in vals[:n]]
            if not label and prev_text:
                label = prev_text
                if i + 1 < len(lines):
                    nl, nv = split_row(lines[i + 1])
                    if nl and not nv and not re.match(r"^(Total|[A-H][\.\-]|\d\.)", nl):
                        label = f"{label} {nl}"
                        i += 1
            rows.append((label, v))
            prev_text = None
        elif label and not vals:
            prev_text = label
        i += 1
    return years, rows


# ------------------------------------------------------------------ English NEPCO tables

def find_table(pages, title_re):
    """(page number, line index, lines) of the table title 'Table (n): <title_re>' (not a contents-page entry, which
    is followed by a page number), else of the first line matching title_re."""
    for strict in (True, False):
        for p, text in enumerate(pages):
            lines = clean_lines(text)
            for j, l in enumerate(lines):
                if strict and not re.search(r"Table\s*\(\s*\d+\s*\)\s*:?\s*(?:" + title_re + ")", l, re.I):
                    continue
                if not strict and not re.search(title_re, l, re.I):
                    continue
                if re.search(r"\.{3,}|\s\d{1,3}$", l) and not header_years(l):
                    continue      # contents-page entry
                return p + 1, j, lines
    return None, None, None


NEXT_TABLE = r"^Table\s*\(\s*\d+\s*\)"


def fuel_key(label, scheme):
    l = label.lower().strip(" -")
    if scheme == "purch":
        for pat, key in ((r"heavy fuel", "Purch_Heavy_fuel_GWh"), (r"natural gas|^n\.? ?gas", "Purch_Gas_GWh"),
                         (r"diesel", "Purch_Diesel_GWh"), (r"oil shale", "Purch_Oil_shale_GWh"),
                         (r"renewable", "Purch_Renewables_GWh"), (r"import", "Purch_Imports_GWh"),
                         (r"industrial sector|others?$|^[bc]- ?others|^[bc]- others", "Purch_Other_GWh"),
                         (r"^total", "Purch_Total_GWh")):
            if re.search(pat, l):
                return key
    if scheme == "gentype":
        for pat, key in ((r"^wind", "Purch_Wind_GWh"), (r"^solar", "Purch_Solar_GWh"), (r"^hydro", "Purch_Hydro_GWh")):
            if re.search(pat, l):
                return key
    if scheme == "fuel":
        for pat, key in ((r"heavy fuel", "Fuel_Heavy_fuel_ktoe"), (r"natural gas|^n\.? ?gas", "Fuel_Gas_ktoe"),
                         (r"diesel", "Fuel_Diesel_ktoe"), (r"oil shale", "Fuel_Oil_shale_ktoe"),
                         (r"^total$", "Fuel_Total_ktoe")):
            if re.search(pat, l):
                return key
    return None


def parse_en(pages, report_year):
    """English NEPCO report -> list of (year, item, value, page, note)."""
    out = []

    def add(y, item, v, page, note=""):
        if v is not None and y <= report_year:
            out.append((int(y), item, v, page, note))

    # purchases by fuel (2018 reports on)
    p, j, lines = find_table(pages, r"Purchased by type of Fuel")
    if p:
        years, rows = year_table(lines, j + 1, NEXT_TABLE)
        seen = set()
        for label, vals in rows:
            key = fuel_key(label, "purch")
            if key and key not in seen:
                seen.add(key)
                for y, v in zip(years or [], vals):
                    add(y, key, v, p)
    # wind / solar / hydro from purchases by type of generation, kept only if they add up to the fuel table's
    # renewables for every year (the 2019 report's table 12 prints its labels one row out of line)
    p, j, lines = find_table(pages, r"Purchased by type of Generation")
    if p:
        years, rows = year_table(lines, j + 1, NEXT_TABLE)
        seen, gt = set(), []
        for label, vals in rows:
            key = fuel_key(label, "gentype")
            if key and key not in seen:
                seen.add(key)
                gt += [(y, key, v) for y, v in zip(years or [], vals) if v is not None and y <= report_year]
        ren = {r[0]: r[2] for r in out if r[1] == "Purch_Renewables_GWh"}
        sums = {}
        for y, key, v in gt:
            sums[y] = sums.get(y, 0.0) + v
        bad = [y for y in sums if y in ren and abs(sums[y] - ren[y]) > 3.0]
        if gt and (bad or not ren):
            out.append((report_year, "_note", f"wind/solar/hydro rows not used: they do not add up to renewables "
                        f"({', '.join(map(str, bad))})" if bad else "no renewables row to check wind/solar/hydro",
                        p, ""))
        else:
            for y, key, v in gt:
                add(y, key, v, p)
    # older basis: production by type of fuel (transposed: 'Total 2017 a b c d e' / '2016 a b c d e')
    p, j, lines = find_table(pages, r"Production by Type of Fuel")
    if p:
        head = " ".join(lines[j + 1:j + 5])
        fourth = "Prod_Other_GWh"
        # the 2017 report prints each year's numbers on the line BEFORE 'Total 2017' / '2016': join them
        joined = []
        for k, line in enumerate(lines[j + 1:], start=j + 1):
            nxt = lines[k + 1] if k + 1 < len(lines) else ""
            if line and all(is_val(t) for t in line.split()) and re.match(r"^(?:Total\s+)?(?:19|20)\d\d$", nxt):
                joined.append(f"{nxt} {line}")
            else:
                joined.append(line)
        for line in joined:
            m = re.match(r"^(?:Total\s+)?((?:19|20)\d\d)\s+(.*)$", line)
            if not m:
                if re.search(NEXT_TABLE, line) or line.startswith("*"):
                    if out and any(r[1].startswith("Prod_") for r in out):
                        break
                continue
            y = int(m.group(1))
            toks = m.group(2).split()
            if not toks or not all(is_val(t) for t in toks):
                continue
            v = [num(t) for t in toks]
            cols = (["Prod_Heavy_fuel_GWh", "Prod_Gas_GWh", "Prod_Diesel_GWh", "Prod_Renewables_GWh", "Prod_Total_GWh"]
                    if len(v) == 5 else
                    ["Prod_Heavy_fuel_GWh", "Prod_Gas_GWh", "Prod_Diesel_GWh", fourth, "Prod_Renewables_GWh",
                     "Prod_Total_GWh"] if len(v) == 6 else None)
            if not cols or None in v:
                continue
            if abs(sum(v[:-1]) - v[-1]) > max(1.0, 0.005 * v[-1]):
                continue          # columns not as expected: leave the row out
            for c, x in zip(cols, v):
                add(y, c, x, p, "coal/others column" if c == fourth and "Coal" in head else "")
            if len(v) == 5:   # no coal/others column in this table: the row's parts are complete without it
                add(y, fourth, 0.0, p, "no coal/others column in this report's table")
    # fuel for power by type (ktoe): 'Fuel Consumption by Type of Fuel' (to 2017) or 'Fuel Consumption for
    # Electricity Generation' (2018 on, by fuel); older reports: 'Electricity Fuel Consumption Type'
    for title in (r"Fuel Consumption by Type of Fuel|Electricity Fuel Consumption Type",
                  r"Fuel Consumption for Electricity Generation"):
        p, j, lines = find_table(pages, title)
        if not p:
            continue
        years, rows = year_table(lines, j + 1, NEXT_TABLE + r"|Contributing|^\*")
        keys = {fuel_key(l, "fuel") for l, _ in rows}
        if "Fuel_Gas_ktoe" not in keys:
            continue
        seen, industrial = set(), any(re.search(r"industr", l, re.I) for l, _ in rows)
        for label, vals in rows:
            key = fuel_key(label, "fuel")
            if key and key not in seen:
                seen.add(key)
                for y, v in zip(years or [], vals):
                    add(y, key, v, p, "table includes industrial self-generation" if industrial else "")
        break
    # peak / minimum load table
    p, j, lines = find_table(pages, r"Peak load for Interconnected System")
    if p:
        for line in lines[j + 1:j + 20]:
            if re.search(NEXT_TABLE, line):
                break
            m = re.match(r"^((?:19|20)\d\d)\s+(\d{3,4})\s+(\d\d/\d\d/\d{4})\s+(\d{3,4})\s+(\d\d/\d\d/\d{4})\s+"
                         r"(\d{3,4})\s+(\d\d/\d\d/\d{4})$", line)
            if m:
                y = int(m.group(1))
                vals = [(int(m.group(2)), m.group(3)), (int(m.group(4)), m.group(5)), (int(m.group(6)), m.group(7))]
                d = [pd.to_datetime(x[1], format="%d/%m/%Y").strftime("%Y-%m-%d") for x in vals]
                add(y, "Peak_morning_MW", vals[0][0], p)
                add(y, "Peak_morning_date", d[0], p)
                add(y, "Peak_evening_MW", vals[1][0], p)
                add(y, "Peak_evening_date", d[1], p)
                add(y, "Min_MW", vals[2][0], p)
                add(y, "Min_date", d[2], p)
                continue
            m = re.match(r"^((?:19|20)\d\d)\s+(\d{3,4})\s+(\d{3,4})\s+(\d{3,4})$", line)
            if m:
                y = int(m.group(1))
                add(y, "Peak_morning_MW", int(m.group(2)), p, "columns labelled Morning/Evening or Summer/Winter")
                add(y, "Peak_evening_MW", int(m.group(3)), p, "columns labelled Morning/Evening or Summer/Winter")
                add(y, "Min_MW", int(m.group(4)), p)
    # older 'System Peak Loads (MW)' tables: interconnected system (sent-out column where the table has one)
    p, j, lines = find_table(pages, r"System Peak Loads")
    if p:
        head = " ".join(lines[j + 1:j + 6])
        sent_out = "Sent-out" in head or "Sent-Out" in head
        for line in lines[j + 1:j + 14]:
            m = re.match(r"^((?:19|20)\d\d)\s+(.*)$", line)
            if not m:
                continue
            v = [num(t) for t in m.group(2).replace("***", "").replace("**", "").replace("*", "").split()
                 if is_val(t)]
            v = [x for x in v if x is not None]
            if len(v) < 4:
                continue
            val = v[1] if sent_out else v[-2]
            if 500 <= val <= 9000:
                add(int(m.group(1)), "Peak_MW_table", val, p,
                    "interconnected system, sent-out" if sent_out else "interconnected system column")
    # report-year peak in the prose
    flat = re.sub(r"-\s*\n\s*", "", "\n".join(pages))
    flat = re.sub(r"\s+", " ", flat)
    for pat in (r"annual peak load (?:for|of) the interconnected system amounted to \((\d{3,4})\) MW (?:during |and "
                r"occurred in |occurred in |in )?([A-Z][a-z]+)?",
                r"peak load of the electric(?:al)? (?:power )?system (?:in %d )?(?:was|amounted to) \((\d{3,4})\) MW "
                r"(?:during |in )?([A-Z][a-z]+)?" % report_year):
        m = re.search(pat, flat, re.I)
        if m:
            pg = next((i + 1 for i, t in enumerate(pages) if m.group(1) in t and "MW" in t), None)
            add(report_year, "Peak_MW_prose", int(m.group(1)), pg, m.group(0)[:200])
            if m.group(2) and m.group(2).capitalize() in MONTHS:
                add(report_year, "Peak_month_prose", m.group(2).capitalize(), pg)
            break
    # purchases table: imports by counterparty
    p, j, lines = find_table(pages, r"NEPCO.s (Electrical Energy|Electric Power) Purchases|NEPCO.s Purchased Energy")
    if p:
        years, rows = year_table(lines, j + 1, NEXT_TABLE)
        tot = {}
        for label, vals in rows:
            l = label.lower()
            if "import" in l:
                key = ("Import_Egypt_GWh" if re.search(r"egypt|eygpt", l) else "Import_Syria_GWh" if "syria" in l
                       else "Import_Other_GWh")
                for y, v in zip(years, vals):
                    add(y, key, v, p, label)
                    tot[y] = tot.get(y, 0.0) + (v or 0.0)
        for y, v in tot.items():
            add(y, "Import_Total_GWh", round(v, 3), p, "sum of the purchases table's import rows")
    # sales table: exports by counterparty
    p, j, lines = find_table(pages, r"NEPCO.s (Electrical Energy|Electric Power) Sales")
    if p:
        years, rows = year_table(lines, j + 1, NEXT_TABLE)
        inside = False
        for label, vals in rows:
            l = label.lower()
            if re.search(r"export", l):
                inside = True
                for y, v in zip(years, vals):
                    add(y, "Export_Total_GWh", v, p, label)
                continue
            if inside and (re.match(r"^[a-h][\.\-\)] ?", l) or l.startswith("total")):
                break
            if inside:
                key = export_key(l)
                for y, v in zip(years, vals):
                    add(y, key, v, p, label)
    # key figures (table 1, 2018 reports on: 'label v_now v_prev %')
    p, j, lines = find_table(pages, r"NEPCO.s Significant Figures")
    if p and j is not None:
        hdr = next((header_years(l) for l in lines[j + 1:j + 4] if header_years(l)), None)
        if hdr and len(hdr) == 2:
            yrs = sorted(hdr, reverse=True) if lines[j + 1].split()[0] == str(max(hdr)) else sorted(hdr)
            for line in lines[j + 1:j + 25]:
                label, vals = split_row(line)
                key = ("Key_Peak_MW" if re.search(r"peak load", label, re.I) else
                       "Key_Purchased_GWh" if re.search(r"^purchased", label, re.I) else
                       "Key_Sold_GWh" if re.search(r"^sold", label, re.I) else None)
                if key and len(vals) >= 2:
                    for y, t in zip(yrs, vals[:2]):
                        add(y, key, num(t), p)
    return out


def export_key(l):
    return ("Export_Egypt_GWh" if "egypt" in l else "Export_Syria_GWh" if "syria" in l else
            "Export_Palestine_Jericho_GWh" if re.search(r"jericho|jerusalem|palestin", l) else
            "Export_Iraq_Trabeel_GWh" if re.search(r"trabeel|border", l) else
            "Export_Iraq_transmission_GWh" if re.search(r"iraq|general company", l) else "Export_Other_GWh")


# ------------------------------------------------------------------ Arabic NEPCO tables (visual order)

AR_TABLES = {   # key -> (title fragment as extracted, {label fragment: item})
    "fuel": ("عونلا بسح ديلوتلا تاطحم يف مدختسملا دوقولا",
             {"ليقث دوقو": "Fuel_Heavy_fuel_ktoe", "يعيبط زاغ": "Fuel_Gas_ktoe", "لزيد": "Fuel_Diesel_ktoe",
              "يتيزلا رخصلا": "Fuel_Oil_shale_ktoe", "عومجملا": "Fuel_Total_ktoe"}),
    "purch_fuel": ("دوقولا عون بسح ةارتشملا",
                   {"ليقث دوقو": "Purch_Heavy_fuel_GWh", "يعيبط زاغ": "Purch_Gas_GWh", "لزيد": "Purch_Diesel_GWh",
                    "يتيز رخص": "Purch_Oil_shale_GWh", "ةددجتم ةقاط": "Purch_Renewables_GWh",
                    "ىرخأ": "Purch_Other_GWh", "ةدروتسملا ةقاطلا": "Purch_Imports_GWh", "عومجملا": "Purch_Total_GWh"}),
    "purch_type": ("ديلوتلا ةيعون بسح ةارتشملا",
                   {"حاير": "Purch_Wind_GWh", "ةيسمش ةقاط": "Purch_Solar_GWh", "يئام": "Purch_Hydro_GWh"}),
    "purch": (")س.و.ج( ةارتشملا ةيئابرهكلا ةقاطلا", {"رصم نم ةدروتسملا ةقاطلا": "Import_Egypt_GWh"}),
    "sales": ("ةلمجلاب ةينطولا ءابرهكلا ةكرش تاعيبم",
              {"ةردصملا ةقاطلا": "Export_Total_GWh", "سدقلا ةكرش": "Export_Palestine_Jericho_GWh",
               "يدودحلا زكرملا": "Export_Iraq_Trabeel_GWh", "ةقاطلا لقنل ةماعلا ةكرشلا": "Export_Iraq_transmission_GWh"}),
}


def lead_run(line):
    toks = line.split()
    k = 0
    while k < len(toks) and is_val(toks[k]):
        k += 1
    return toks[:k], " ".join(toks[k:])


def parse_ar(pages, report_year):
    """Arabic NEPCO report -> {table key: [(year, item, value, page, note)]}. Rows: '[%] v_newest .. v_oldest label'."""
    found = {}
    for key, (title, labels) in AR_TABLES.items():
        recs = []
        for p, text in enumerate(pages):
            lines = clean_lines(text)
            hit = next((j for j, l in enumerate(lines) if title in l), None)
            if hit is None:
                continue
            years = None
            for j in range(hit, min(hit + 5, len(lines))):
                ys = header_years(lines[j])
                if ys:
                    years, start = ys, j + 1
                    break
            if not years:
                continue
            n = len(years)          # header newest first, as printed
            prev_text = ""
            seen = set()
            for line in lines[start:]:
                if "مقر لودج" in line:
                    break
                vals, label = lead_run(line)
                if len(vals) in (n, n + 1):
                    v = [num(t) for t in (vals[1:] if len(vals) == n + 1 else vals)]
                    lab = label or prev_text
                    for frag, item in labels.items():
                        if frag in lab and item not in seen:
                            seen.add(item)
                            for y, x in zip(years, v):
                                if x is not None and y <= report_year:
                                    recs.append((y, item, x, p + 1, "Arabic report"))
                            break
                    prev_text = ""
                elif label and not vals:
                    prev_text = label
            break
        if key == "purch":
            recs = [r for r in recs if r[1] == "Import_Egypt_GWh"]
            recs += [(y, "Import_Total_GWh", v, pg, "sum of the purchases table's import rows")
                     for y, it, v, pg, _ in list(recs)]
        found[key] = recs
    # peak / minimum table: 'min_date min_MW eve_date eve_MW morn_date morn_MW year'
    recs = []
    for p, text in enumerate(pages):
        for line in clean_lines(text):
            m = re.match(r"^(\d{4}/\d\d/\d\d) (\d{3,4}) (\d{4}/\d\d/\d\d) (\d{3,4}) (\d{4}/\d\d/\d\d) (\d{3,4}) "
                         r"((?:19|20)\d\d)$", line)
            if m:
                y = int(m.group(7))
                f = lambda s: pd.to_datetime(s, format="%Y/%m/%d").strftime("%Y-%m-%d")
                recs += [(y, "Min_date", f(m.group(1)), p + 1, "Arabic report"),
                         (y, "Min_MW", int(m.group(2)), p + 1, "Arabic report"),
                         (y, "Peak_evening_date", f(m.group(3)), p + 1, "Arabic report"),
                         (y, "Peak_evening_MW", int(m.group(4)), p + 1, "Arabic report"),
                         (y, "Peak_morning_date", f(m.group(5)), p + 1, "Arabic report"),
                         (y, "Peak_morning_MW", int(m.group(6)), p + 1, "Arabic report")]
        if recs:
            break
    found["peak"] = recs
    # key figures: '% 2025 2024 2023' then '17.1 4800 4,100 4,240 ) و.م( يئابرهكلا ماظنلل ىصقلأا لمحلا'
    recs = []
    for p, text in enumerate(pages):
        lines = clean_lines(text)
        if not any("ةزيمم تايئاصحإ" in l for l in lines):
            continue
        years = next((header_years(l) for l in lines if header_years(l)), None)
        if not years:
            continue
        for line in lines:
            vals, label = lead_run(line)
            item = ("Key_Peak_MW" if "ىصقلأا لمحلا" in label else "Key_Purchased_GWh" if "ةارتشملا" in label else
                    "Key_Sold_GWh" if "ةعابملا" in label and "روبعلا" not in label else None)
            if item and len(vals) == len(years) + 1:
                for y, t in zip(years, vals[1:]):
                    recs.append((y, item, num(t), p + 1, "Arabic report"))
        break
    found["key"] = recs
    return found


def accept_arabic(found, known):
    """Keep an Arabic table only if its values for years the English reports also give agree (>= 70% of >= 6 shared
    values within 1% or 0.15; dates must match exactly). Returns (records, notes)."""
    out, notes = [], []
    for key, recs in found.items():
        shared = [(r, known[(r[0], r[1])]) for r in recs if (r[0], r[1]) in known]
        ok = 0
        for r, v in shared:
            if isinstance(v, str) or isinstance(r[2], str):
                ok += str(v) == str(r[2])
            else:
                ok += abs(float(v) - float(r[2])) <= max(0.15, 0.01 * abs(float(v)))
        if len(shared) >= 6 and ok / len(shared) >= 0.7:
            out += recs
            notes.append(f"Arabic table '{key}' accepted ({ok}/{len(shared)} shared values agree)")
        else:
            notes.append(f"Arabic table '{key}' REJECTED ({ok}/{len(shared)} shared values agree)")
    return out, notes


# ------------------------------------------------------------------ MEMR prose

def parse_memr(pages, report_year):
    """MEMR annual report -> records for gas for power (MMcf, MMcf/d) and Risha; sentences kept verbatim."""
    t = re.sub(r"\s+", " ", " ".join(pages))
    out = []
    unit = r"(?:million cubic feet|MMscf|MMcf|Mcf)"
    pats = [
        # 'quantities of natural gas consumed ... in electricity generation plants during 2024 amounted to approximately
        #  119,456 million cubic feet, at an average of 327 million cubic feet per day'
        rf"(?:natural gas|LNG)[^.]{{0,160}}?(?:consumed|consumption)[^.]{{0,200}}?(?:power|electric)[^.]{{0,120}}?"
        rf"(?:amounted to|reached|totals)[^.\d]{{0,30}}\(?([\d,]+(?:\.\d+)?)\)? ?{unit}"
        rf"(?:[^.]{{0,60}}?(?:average|rate)[^.\d]{{0,40}}\(?([\d,]+(?:\.\d+)?)\)? ?{unit})?",
        # 'average daily consumption of natural gas to generate electric power amounted to 337 million cubic feet'
        rf"average (?:daily )?(?:amount |consumption |quantity )?of natural gas[^.]{{0,160}}?(?:power|electric)"
        rf"[^.]{{0,80}}?(?:amounted to|reached|totals)[^.\d]{{0,30}}\(?([\d,]+(?:\.\d+)?)\)? ?{unit}()",
        rf"(?:daily average|average daily) (?:natural gas )?consumption[^.]{{0,40}}?natural gas[^.]{{0,80}}?"
        rf"(?:power|electric)[^.]{{0,60}}?(?:amounted to|reached)[^.\d]{{0,20}}(?:around |about )?\(?([\d,]+)\)? ?"
        rf"{unit}()",
        rf"quantity of natural gas consumed for electricity generation reached about \(?([\d,]+)\)? ?{unit} per day()",
    ]
    hits = []
    for k, pat in enumerate(pats):
        for m in re.finditer(pat, t, re.I):
            hits.append((m.start(), k, m))
    hits.sort()
    if hits:
        _, k, m = hits[0]
        sent = m.group(0)
        a = float(m.group(1).replace(",", ""))
        b = float(m.group(2).replace(",", "")) if m.group(2) else None
        after = t[m.end():m.end() + 15].lower()
        daily_only = k > 0 or (b is None and (after.lstrip().startswith(("per day", "/d", "a day")) or
                                              bool(re.search(r"average (?:daily|amount)|daily",
                                                             sent[:sent.find(m.group(1))], re.I))))
        if daily_only:
            mmcfd, mmcf = a, None
        else:
            mmcf, mmcfd = a, b
        basis = "LNG consumed by power plants (report)" if re.search(r"\bLNG total quantities consumed", sent) else \
            "natural gas consumed for electricity generation, all sources"
        if "Mcf" in sent and "MMcf" not in sent and "million" not in sent:
            basis += "; the report prints 'Mcf': value/days = its daily figure, so million cubic feet"
        if mmcf is not None:
            out.append((report_year, "MEMR_Gas_power_MMcf", mmcf, None, basis))
        if mmcfd is not None:
            out.append((report_year, "MEMR_Gas_power_MMcf_d", mmcfd, None, basis))
            out.append((report_year, "MEMR_Gas_power_Bcf_d", round(mmcfd / 1000, 4), None, "MMcf/d / 1000"))
        out.append((report_year, "MEMR_sentence_gas", sent[:600], None, basis))
        others = sorted({float(h[2].group(1).replace(",", "")) for h in hits[1:]} - {a, b})
        if others:
            out.append((report_year, "MEMR_other_gas_figures", ", ".join(f"{x:g}" for x in others), None,
                        "other gas-for-power sentences in the same report: " +
                        " | ".join(h[2].group(0)[:250] for h in hits[1:3])))
    m = re.search(rf"Risha[^.]{{0,200}}?(?:reached|amounted to|supplied|produced)[^.\d]{{0,40}}(?:around |about |"
                  rf"approximately )?([\d,]+(?:\.\d+)?) ?{unit}[^.]{{0,120}}?(?:average|averaging|around)"
                  rf"[^.\d]{{0,20}}(?:of )?([\d.]+) ?{unit}", t, re.I)
    m2 = None if m else re.search(rf"Risha gas field produce an average of ([\d.]+) {unit} per day", t, re.I)
    if m:
        out.append((report_year, "MEMR_Risha_MMcf", float(m.group(1).replace(",", "")), None,
                    "Risha field gas (production / supplied for power, as stated)"))
        out.append((report_year, "MEMR_Risha_MMcf_d", float(m.group(2)), None, "as stated"))
        out.append((report_year, "MEMR_sentence_risha", m.group(0)[:500], None, ""))
    elif m2:
        out.append((report_year, "MEMR_Risha_MMcf_d", float(m2.group(1)), None, "Risha wells' production, as stated"))
        out.append((report_year, "MEMR_sentence_risha", m2.group(0)[:500], None, ""))
    return out


# ------------------------------------------------------------------ downloads

class Net:
    def __init__(self):
        self.s = requests.Session()
        self.requests = {"nepco": 0, "memr": 0}

    def call(self, method, url, timeout=240):
        host = "nepco" if "nepco" in url else "memr"
        for attempt in range(3):
            self.requests[host] += 1
            try:
                r = self.s.request(method, url, headers=H, timeout=(20, timeout), verify=False, allow_redirects=True)
                time.sleep(PAUSE_S)
                return r
            except requests.exceptions.RequestException as e:
                log(f"  {method} {url}: {type(e).__name__} {str(e)[:120]} (attempt {attempt + 1})")
                time.sleep(5 * (attempt + 1))
        return None

    def pdf(self, url):
        r = self.call("GET", url)
        if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
            return r
        return None


def pdf_pages(content):
    import pdfplumber
    pages = []
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for pg in pdf.pages:
            try:
                pages.append(pg.extract_text() or "")
            except Exception as e:  # noqa: BLE001
                pages.append("")
                log(f"   page error {e}")
    return pages


def nepco_url(y, lang):
    if lang == "EN":
        return f"{NEPCO}/store/docs/web/{y}_en.pdf" if y < 2019 else f"{NEPCO}/store/DOCS/web/{y}_EN.pdf"
    return f"{NEPCO}/store/DOCS/web/{y}_AR.pdf"


def memr_links(net):
    r = net.call("GET", MEMR_LIST, timeout=60)
    if r is None or r.status_code != 200:
        return {}
    links = {}
    for l in re.findall(r'[^"\'\s<>()]+\.pdf', r.text, re.I):
        if "/en/" not in l.lower() or "annual" not in l.lower():
            continue
        m = re.search(r"(20\d\d)", l)
        if m:
            links.setdefault(int(m.group(1)), MEMR + l if l.startswith("/") else l)
    return links


# ------------------------------------------------------------------ workbook

def read_sheet(path, name):
    try:
        return pd.read_excel(path, sheet_name=name)
    except Exception:  # noqa: BLE001
        return None


def pick(byrep):
    """Newest report wins per (Year, Item); NEPCO table values before prose within a report."""
    b = byrep.copy()
    # newest report first; for the same report year the English report before the Arabic one
    b["_rank"] = b["Report_year"].astype(int) * 10 + (~b["Report"].astype(str).str.endswith(" AR")).astype(int)
    b = b.sort_values(["Year", "Item", "_rank"], ascending=[True, True, False], kind="stable")
    return b.drop_duplicates(["Year", "Item"], keep="first").drop(columns="_rank")


def build(byrep):
    used = pick(byrep)
    wide = used.pivot(index="Year", columns="Item", values="Value")
    # annual peak = larger of the morning/evening peaks (table), else the report's prose figure
    peak, pdate, src = {}, {}, {}
    for y in wide.index:
        m = pd.to_numeric(wide.at[y, "Peak_morning_MW"], errors="coerce") if "Peak_morning_MW" in wide else None
        e = pd.to_numeric(wide.at[y, "Peak_evening_MW"], errors="coerce") if "Peak_evening_MW" in wide else None
        cand = [(v, k) for v, k in ((m, "morning"), (e, "evening")) if v is not None and pd.notna(v)]
        if cand:
            v, k = max(cand)
            peak[y] = v
            pdate[y] = wide.at[y, f"Peak_{k}_date"] if f"Peak_{k}_date" in wide and pd.notna(
                wide.at[y, f"Peak_{k}_date"]) else None
            src[y] = f"larger of morning/evening peak ({k})"
        elif "Peak_MW_table" in wide and pd.notna(wide.at[y, "Peak_MW_table"]):
            peak[y] = wide.at[y, "Peak_MW_table"]
            src[y] = "older 'System Peak Loads' table, interconnected system"
        elif "Peak_MW_prose" in wide and pd.notna(wide.at[y, "Peak_MW_prose"]):
            peak[y] = wide.at[y, "Peak_MW_prose"]
            src[y] = "report-year peak in the report's text"
    wide["Peak_MW"] = pd.Series(peak)
    wide["Peak_date"] = pd.Series(pdate)
    wide["Peak_source"] = pd.Series(src)
    cols = [c for c in ITEMS if c in wide.columns] + ["Peak_source"]
    ann = wide.reindex(columns=cols)
    for c in ann.columns:
        if ITEMS.get(c) not in ("date", "") and c != "Peak_source":
            ann[c] = pd.to_numeric(ann[c], errors="coerce")
    ann = ann.dropna(how="all")
    ann.index = ann.index.astype(int)
    ann.insert(0, "date", pd.to_datetime(ann.index.astype(str) + "-01-01"))
    ann.index.name = "Year"
    return ann, used


def checks(ann, byrep):
    rows = []

    def chk(y, name, parts, total, tol):
        p = [ann.at[y, c] for c in parts if c in ann.columns]
        t = ann.at[y, total] if total in ann.columns else None
        if t is None or pd.isna(t) or not p or all(pd.isna(x) for x in p):
            return
        s = float(sum(x for x in p if pd.notna(x)))
        ok = abs(s - t) <= tol
        rows.append({"Year": y, "Check": name, "Parts_sum": round(s, 2), "Total": t, "Difference": round(s - t, 2),
                     "Status": "ok" if ok else "FAILED"})

    for y in ann.index:
        chk(y, "Purchases by fuel = total purchased (GWh)",
            ["Purch_Gas_GWh", "Purch_Oil_shale_GWh", "Purch_Heavy_fuel_GWh", "Purch_Diesel_GWh",
             "Purch_Renewables_GWh", "Purch_Imports_GWh", "Purch_Other_GWh"], "Purch_Total_GWh", 3.0)
        chk(y, "Wind + solar + hydro = renewables (GWh)", ["Purch_Wind_GWh", "Purch_Solar_GWh", "Purch_Hydro_GWh"],
            "Purch_Renewables_GWh", 3.0)
        chk(y, "Production by fuel (older basis) = total (GWh)",
            ["Prod_Heavy_fuel_GWh", "Prod_Gas_GWh", "Prod_Diesel_GWh", "Prod_Other_GWh", "Prod_Renewables_GWh"],
            "Prod_Total_GWh", 3.0)
        chk(y, "Exports by counterparty = exported total (GWh)",
            ["Export_Egypt_GWh", "Export_Syria_GWh", "Export_Palestine_Jericho_GWh", "Export_Iraq_Trabeel_GWh",
             "Export_Iraq_transmission_GWh", "Export_Other_GWh"], "Export_Total_GWh", 2.0)
        if "Purch_Imports_GWh" in ann and pd.notna(ann.at[y, "Purch_Imports_GWh"]) and "Import_Total_GWh" in ann:
            chk(y, "Imports in the purchases table = imports in the fuel table (GWh)", ["Import_Total_GWh"],
                "Purch_Imports_GWh", 2.0)
        if "Key_Purchased_GWh" in ann:
            chk(y, "Purchases by fuel total = table 1 purchased energy (GWh)", ["Purch_Total_GWh"],
                "Key_Purchased_GWh", 2.0)
        # fuel ktoe: only when the table has no industrial block
        notes = byrep[(byrep["Year"] == y) & (byrep["Item"] == "Fuel_Total_ktoe")]
        if len(notes) and not str(pick(notes).iloc[0]["Note"]).startswith("table includes industrial"):
            chk(y, "Fuel for power by type = total (ktoe)",
                ["Fuel_Gas_ktoe", "Fuel_Oil_shale_ktoe", "Fuel_Heavy_fuel_ktoe", "Fuel_Diesel_ktoe"],
                "Fuel_Total_ktoe", 1.0)
        # MEMR annual MMcf / days vs stated MMcf/d
        a = ann.at[y, "MEMR_Gas_power_MMcf"] if "MEMR_Gas_power_MMcf" in ann else None
        d = ann.at[y, "MEMR_Gas_power_MMcf_d"] if "MEMR_Gas_power_MMcf_d" in ann else None
        if a is not None and d is not None and pd.notna(a) and pd.notna(d):
            days = 366 if (y % 4 == 0) else 365
            rows.append({"Year": y, "Check": "MEMR MMcf / days in year = MMcf/d stated", "Parts_sum": round(a / days, 1),
                         "Total": d, "Difference": round(a / days - d, 1),
                         "Status": "ok" if abs(a / days - d) <= max(1.0, 0.01 * d) else "FAILED"})
        g = ann.at[y, "Fuel_Gas_ktoe"] if "Fuel_Gas_ktoe" in ann else None
        if d is not None and g is not None and pd.notna(d) and pd.notna(g) and g > 0:
            days = 366 if (y % 4 == 0) else 365
            ratio = d * days / g
            rows.append({"Year": y, "Check": "MEMR gas MMcf per NEPCO gas ktoe (about 37-42 expected at ~1,000 Btu/cf)",
                         "Parts_sum": round(d * days, 0), "Total": g, "Difference": round(ratio, 1),
                         "Status": "ok" if 35 <= ratio <= 45 else "CHECK: bases differ"})
    # values revised between reports (> 2%)
    num_ = byrep[pd.to_numeric(byrep["Value"], errors="coerce").notna()].copy()
    num_["V"] = pd.to_numeric(num_["Value"])
    for (y, it), g in num_.groupby(["Year", "Item"]):
        if g["V"].nunique() > 1:
            lo, hi = g["V"].min(), g["V"].max()
            if hi - lo > max(0.5, 0.02 * abs(hi)):
                rows.append({"Year": y, "Check": f"Revised between reports: {it}",
                             "Parts_sum": None, "Total": None, "Difference": round(hi - lo, 2),
                             "Status": "; ".join(f"{r.Report}: {r.Value}" for r in g.itertuples())[:500]})
    return pd.DataFrame(rows)


def daily_peak_check(ann):
    """NEPCO annual peak vs the daily Max Load Archive workbook; the archive's value on the report's peak dates and
    the day after (the archive is dated one day late from 2023 on)."""
    if not os.path.exists(DAILY_PEAK):
        return pd.DataFrame([{"Year": None, "Note": f"{os.path.basename(DAILY_PEAK)} not found"}])
    d = pd.read_excel(DAILY_PEAK, sheet_name="Daily")
    d["date"] = pd.to_datetime(d["date"])
    x = d.set_index("date")["Peak_Load_MW"].apply(pd.to_numeric, errors="coerce").dropna()
    rows = []
    for y in ann.index:
        xs = x[x.index.year == y]
        if not len(xs):
            continue
        r = {"Year": y, "NEPCO_annual_peak_MW": ann.at[y, "Peak_MW"] if "Peak_MW" in ann else None,
             "NEPCO_peak_date": ann.at[y, "Peak_date"] if "Peak_date" in ann else None,
             "Daily_archive_max_MW": xs.max(), "Daily_archive_max_date": xs.idxmax().strftime("%Y-%m-%d"),
             "Days_posted": len(xs)}
        for k in ("morning", "evening"):
            v = ann.at[y, f"Peak_{k}_MW"] if f"Peak_{k}_MW" in ann else None
            dt = ann.at[y, f"Peak_{k}_date"] if f"Peak_{k}_date" in ann else None
            r[f"NEPCO_{k}_peak_MW"] = v
            r[f"NEPCO_{k}_date"] = dt
            if dt is not None and pd.notna(dt):
                t = pd.Timestamp(dt)
                r[f"Archive_on_{k}_date"] = x.get(t)
                r[f"Archive_day_after_{k}_date"] = x.get(t + pd.Timedelta(days=1))
        r["Archive_median_MW"] = xs.median()
        rows.append(r)
    df = pd.DataFrame(rows)
    return df.set_index("Year") if len(df) else df


PEAK_NOTES = [
    "What the annual reports say about the daily archive's 4,800 MW day and the 2023 step (Oct 2026 reading):",
    "4,800 MW: NEPCO's 2025 report (Arabic, table 1 and table 'Electrical loads of the system') gives the 2025 system "
    "peak as 4,800 MW, an EVENING peak on 13 Aug 2025 (morning peak 4,510 MW the same day), +17.1% on 2024's 4,100 MW; "
    "the winter peak was 4,470 MW on 24 Feb 2025. The archive shows 4,800 on 14 Aug and 4,470 on 25 Feb 2025: the same "
    "figures, dated one day later. So the 4,800 MW day is real (the August 2025 heatwave), not a typo.",
    "2023 step: the reports show no step in demand - purchased energy rose 3.4% in 2023 and 5.8% in 2024, the annual "
    "peak went 4,010 (2022, evening) / 4,240 (2023) / 4,100 (2024) MW and the minimum load stayed 1,115-1,130 MW. In "
    "2021-22 the archive's annual maximum equals the report's MORNING (daytime) peak on the same date (3,540 MW on "
    "20 Dec 2021; 3,900 MW on 24 Jan 2022) and misses the higher evening peaks (3,770 / 4,010 MW); from 2023 it "
    "matches the day's overall peak dated one day later (4,240 MW: report 14 Aug, archive 15 Aug 2023; 4,800 MW: "
    "13/14 Aug 2025; the 31 Dec 2024 archive figure, 4,040 MW, is neither of that day's peaks, 3,715 / 4,100). The "
    "year-on-year rise of the archive's weekly median jumps from about 0-10% to 20-50% after mid-2023 (spring 2024 "
    "+50%). The archive therefore appears to have switched in 2023 from the daytime peak (which distributed solar lowers, most in spring) "
    "to the daily maximum incl. the evening peak, which lifts spring and autumn days most (April median about 1,950 MW "
    "in 2020-23, 3,000 MW in 2024). The reports do not mention the archive; the switch is inferred from these "
    "matches, not stated by NEPCO.",
]


def notes_lines(releases, ann):
    yrs = lambda c: (f"{int(ann[c].dropna().index.min())}-{int(ann[c].dropna().index.max())}"
                     if c in ann and ann[c].notna().any() else "none")
    return [
        "Jordan power annual statistics - NEPCO annual reports (English; Arabic only where no English report exists) "
        "and MEMR annual reports",
        "",
        "UNITS",
        "GWh per year (energy purchased / produced / imported / exported), thousand tonnes of oil equivalent (ktoe, "
        "NEPCO 'T.T.O.E') for fuel burned, MW for loads, MMcf and MMcf/d (million cubic feet) for MEMR gas; "
        "Bcf/d = MMcf/d / 1000. All as published; no energy-content conversion.",
        "",
        "SOURCE AND METHOD",
        "NEPCO annual reports https://www.nepco.com.jo/en/AnnualReports.aspx (/store/docs/web/<yyyy>_en.pdf to 2018, "
        "/store/DOCS/web/<yyyy>_EN.pdf from 2019, Arabic /store/DOCS/web/<yyyy>_AR.pdf); MEMR annual reports "
        "https://www.memr.gov.jo/En/List/Annual_Reports. Text layer only (pdfplumber); image-only PDFs (NEPCO 2007, "
        "2009, 2016) give nothing, their years come from later reports' multi-year tables.",
        "Newest report wins per year and item; every value is on 'By report', the one used on 'Sources used'. "
        "Arabic tables are used only where they agree with the English reports on the years both cover.",
        "",
        "COLUMNS",
        f"Purch_*: NEPCO energy PURCHASED (sent-out, interconnected system) by fuel, {yrs('Purch_Total_GWh')}; wind, "
        "solar and hydro from the 'by type of generation' table; Imports = Egypt interconnection; Other = small "
        "industrial suppliers.",
        f"Prod_*: OLDER BASIS, national gross production by fuel incl. industrial self-generation, imports excluded "
        f"(reports 2008-2017), {yrs('Prod_Total_GWh')}. Never mixed with Purch_*.",
        f"Fuel_*_ktoe: fuel consumption for electricity generation, {yrs('Fuel_Gas_ktoe')} (to the 2017 report the "
        "total includes industrial self-generators; 2024 gas was revised from 3,065.9 (EN 2024, 'primary') to "
        "3,037.2 (AR 2025)).",
        f"MEMR_Gas_power_*: natural gas consumed for electricity generation from MEMR's prose, {yrs('MEMR_Gas_power_MMcf_d')} "
        "(2016 = LNG consumed by power plants only, as worded; 2015 and 2018 reports give no figure); "
        "MEMR_Risha_*: Risha field gas where the report states it.",
        f"Peak_*: interconnected-system peak and minimum load. Peak_MW = larger of the morning and evening peaks of "
        f"the report table (2013 on), else the report-year figure in the report's text; {yrs('Peak_MW')}.",
        f"Import_* / Export_*: NEPCO purchases from / sales to each counterparty, {yrs('Export_Total_GWh')}.",
        "",
        "CHECKS",
        "'Checks': parts against totals (tolerance 1-3 GWh / 1 ktoe for rounding), MEMR MMcf / days against its "
        "MMcf/d, MEMR MMcf per NEPCO ktoe of gas, values revised between reports (>2%). 'Daily vs annual peak': NEPCO's "
        "annual peak against the daily Max Load Archive (jordan_peak_load_daily.xlsx).",
        *PEAK_NOTES,
        "",
        "SCHEDULE",
        "Owner: annual data, so NOT on the 1st/15th schedule: manual runs plus a light check on 1 Sep and 1 Dec "
        "(jordan_power_annual.yml). A check exits without rewriting when no new or changed report is found "
        f"(Last-Modified, then SHA-256 on 'Releases'). {len(releases)} reports on 'Releases'.",
        "Script: middle_east/JORDAN_NEPCO_ANNUAL.py; probe: discovery_archive/middle_east/JORDAN_ANNUAL_REPORTS_PROBE.py.",
    ]


# ------------------------------------------------------------------ main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--reparse", action="store_true", help="download and parse every report again")
    ap.add_argument("--recheck-all", action="store_true", help="HEAD every known report, not only the last 3 years")
    ap.add_argument("--text-dump", help="probe text dump (json.gz): parse it instead of downloading")
    args = ap.parse_args()

    rel = None if args.reparse else read_sheet(args.out, "Releases")
    byrep_old = None if args.reparse else read_sheet(args.out, "By report")
    known = {} if rel is None else {r["Report"]: r for r in rel.to_dict("records")}
    this_year = datetime.now(timezone.utc).year
    todo = {}     # report name -> dict(publisher, year, lang, url, lm, sha, pages)

    def add_todo(name, pub, y, lang, url, lm, content=None, pages=None):
        sha = hashlib.sha256(content).hexdigest()[:16] if content is not None else ""
        if content is not None and name in known and str(known[name].get("SHA256")) == sha:
            return False      # same file (only Last-Modified moved)
        todo[name] = {"Publisher": pub, "Report_year": y, "Language": lang, "URL": url, "Last_Modified": lm or "",
                      "Bytes": len(content) if content is not None else None, "SHA256": sha,
                      "pages": pages if pages is not None else pdf_pages(content)}
        return True

    if args.text_dump:
        dump = json.load(gzip.open(args.text_dump, "rt"))
        for k, v in dump.items():
            m = re.match(r"(nepco|memr)_(\d{4})_(en|ar)$", k)
            if not m:
                continue
            pub, y, lang = ("NEPCO" if m.group(1) == "nepco" else "MEMR"), int(m.group(2)), m.group(3).upper()
            if pub == "MEMR" and y < FIRST_MEMR:
                continue
            name = f"{pub} {y} {lang}"
            todo[name] = {"Publisher": pub, "Report_year": y, "Language": lang, "URL": v["url"],
                          "Last_Modified": v.get("last_modified") or "", "Bytes": v.get("bytes"), "SHA256": "dump",
                          "pages": v["pages"]}
    else:
        net = Net()
        have_en = {int(r["Report_year"]) for r in known.values() if r["Publisher"] == "NEPCO" and r["Language"] == "EN"}
        # known reports: HEAD the recent ones (old reports do not change)
        for name, r in known.items():
            y = int(r["Report_year"])
            if not args.recheck_all and y < this_year - 3:
                continue
            h = net.call("HEAD", r["URL"], timeout=60)
            lm = h.headers.get("Last-Modified", "") if h is not None and h.status_code == 200 else None
            if lm is None or lm == str(r.get("Last_Modified") or ""):
                continue
            g = net.pdf(r["URL"])
            if g is not None and add_todo(name, r["Publisher"], y, r["Language"], r["URL"], lm, g.content):
                log(f"{name}: changed (Last-Modified {r.get('Last_Modified')} -> {lm})")
        # NEPCO reports not yet found: English for every year, Arabic for years without an English report
        for y in range(FIRST_NEPCO, this_year):
            for lang in ("EN", "AR"):
                name = f"NEPCO {y} {lang}"
                if name in known or (lang == "AR" and (y in have_en or f"NEPCO {y} EN" in todo)):
                    continue
                if lang == "AR" and y < this_year - 2:
                    continue
                url = nepco_url(y, lang)
                g = net.pdf(url)
                if g is not None:
                    add_todo(name, "NEPCO", y, lang, url, g.headers.get("Last-Modified", ""), g.content)
                    log(f"{name}: new ({len(g.content):,} bytes)")
                    if lang == "EN":
                        have_en.add(y)
            if net.requests["nepco"] > 55:
                log("NEPCO request budget reached (the server cuts a runner off at ~65); the rest next run")
                break
        # MEMR: the list page names each year's file
        links = memr_links(net)
        for y, url in sorted(links.items()):
            name = f"MEMR {y} EN"
            if y < FIRST_MEMR or (name in known and known[name]["URL"] == url):
                continue
            g = net.pdf(url)
            if g is not None:
                add_todo(name, "MEMR", y, "EN", url, g.headers.get("Last-Modified", ""), g.content)
                log(f"{name}: new or moved ({len(g.content):,} bytes)")
        log(f"requests: {net.requests}")

    if not todo and byrep_old is not None:
        log("No new or changed report; nothing to do.")
        return 0

    recs = [] if byrep_old is None else [r for r in byrep_old.to_dict("records") if r["Report"] not in todo]
    rel_rows = {k: v for k, v in known.items() if k not in todo}
    # English NEPCO and MEMR first, Arabic last (validated against the English values)
    order = sorted(todo, key=lambda k: (todo[k]["Language"] == "AR", todo[k]["Report_year"]))
    for name in order:
        t = todo[name]
        pages, y = t["pages"], t["Report_year"]
        text_pages = sum(1 for p in pages if len(p.strip()) > 50)
        notes = []
        if text_pages < 5:
            r_recs = []
            notes.append("image-only PDF (no text layer): nothing read")
        elif t["Publisher"] == "MEMR":
            r_recs = parse_memr(pages, y)
            if not r_recs:
                notes.append("no gas-for-power sentence found")
        elif t["Language"] == "EN":
            r_recs = parse_en(pages, y)
        else:
            en_years = {int(r["Report_year"]) for r in list(rel_rows.values()) + [todo[k] for k in todo]
                        if r["Publisher"] == "NEPCO" and r["Language"] == "EN"}
            if y in en_years:
                r_recs = []
                notes.append("English report exists: Arabic not used")
            else:
                cur = pick(pd.DataFrame(recs)) if recs else pd.DataFrame(columns=["Year", "Item", "Value"])
                kn = {(int(r.Year), r.Item): r.Value for r in cur.itertuples()}
                r_recs, notes = accept_arabic(parse_ar(pages, y), kn)
        notes += [r[2] for r in r_recs if r[1] == "_note"]
        r_recs = [r for r in r_recs if r[1] != "_note"]
        for yy, item, v, pg, note in r_recs:
            recs.append({"Report": name, "Publisher": t["Publisher"], "Report_year": y, "Year": int(yy), "Item": item,
                         "Value": v, "Page": pg, "Note": note})
        log(f"{name}: {len(r_recs)} values; {'; '.join(notes)}")
        rel_rows[name] = {"Report": name, "Publisher": t["Publisher"], "Report_year": y, "Language": t["Language"],
                          "URL": t["URL"], "Last_Modified": t["Last_Modified"], "Bytes": t["Bytes"],
                          "SHA256": t["SHA256"], "Pages": len(pages), "Text_pages": text_pages,
                          "Read_utc": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M"),
                          "Values": len(r_recs), "Notes": "; ".join(notes)[:1500]}
    byrep = pd.DataFrame(recs, columns=["Report", "Publisher", "Report_year", "Year", "Item", "Value", "Page", "Note"])
    data = byrep[~byrep["Item"].astype(str).str.startswith("MEMR_sentence") &
                 ~byrep["Item"].astype(str).str.startswith("MEMR_other")]
    ann, used = build(data)
    chk = checks(ann, data)
    dp = daily_peak_check(ann)
    sent = byrep[byrep["Item"].astype(str).str.startswith(("MEMR_sentence", "MEMR_other"))]
    releases = pd.DataFrame(sorted(rel_rows.values(), key=lambda r: (r["Publisher"], int(r["Report_year"]),
                                                                    r["Language"])))
    sheets = {
        "Annual": ann,
        "Checks": chk.set_index("Year") if len(chk) else pd.DataFrame({"Year": []}).set_index("Year"),
        "Daily vs annual peak": dp,
        "MEMR sentences": sent[["Report", "Year", "Item", "Value", "Note"]].set_index("Report"),
        "Sources used": used.set_index("Year"),
        "By report": byrep.sort_values(["Report_year", "Report", "Year", "Item"]).set_index("Report"),
        "Releases": releases.set_index("Report"),
    }
    lines = notes_lines(releases, ann)
    titles = [l for l in lines if l.isupper() and l]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, lines, titles + [lines[0]])
    nfail = int(chk["Status"].astype(str).str.startswith("FAILED").sum()) if len(chk) else 0
    log(f"Saved {args.out}: {len(ann)} years ({ann.index.min()}-{ann.index.max()}), {nfail} failed checks")
    gh = os.environ.get("GITHUB_OUTPUT")
    if gh:
        with open(gh, "a") as f:
            f.write("updated=1\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
