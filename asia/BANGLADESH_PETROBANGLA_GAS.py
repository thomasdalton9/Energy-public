"""
Bangladesh natural gas, daily, from Petrobangla's "Daily Gas & Condensate Production and Distribution Report"
(দৈনিক গ্যাস ইনটেক-অফটেক প্রতিবেদন), Production & Marketing Division:
https://petrobangla.org.bd/pages/reports (report type "daily gas report"). Found via
discovery_archive/asia/BD_TH_DISCOVERY*.py.

One PDF per gas day (08:00 to 08:00), listed newest first, 10 per page. Each report gives, in MMCFD:
  I.  Production: state companies BGFCL, SGFL, BAPEX (by field), IOCs (Chevron: Bibiyana, Jalalabad,
      Moulavibazar; Tullow: Bangora), RPGCL R-LNG (regasified LNG from the FSRUs), Grand Total
  II. Distribution: gas to power plants (demand and supply), fertiliser (max demand and supply), others
      (industry, captive power, CNG, households via the distribution companies), total

Writes output/Data and Chart Outputs/bangladesh_gas.xlsx:
  Daily   date (end of the gas day, as Petrobangla labels it), MMCFD: Prod_BGFCL, Prod_SGFL, Prod_BAPEX,
          Prod_state (1+2+3), Prod_IOC, Bibiyana, Jalalabad, Moulavibazar, Bangora, RLNG, Total_supply,
          Power_demand, Power_supply, Fertiliser_demand, Fertiliser_supply, Others_supply, Total_distribution

Incremental: the Daily sheet is the history store; the listing is read newest first only as far back as the
earliest missing day (from DATA_START), and only reports not saved yet are downloaded (plus REVISION_DAYS).
A checkpoint is written every BATCH reports. Runs on the 1st and 15th.

    python3 asia/BANGLADESH_PETROBANGLA_GAS.py
"""
import argparse
import io
import json
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor
from datetime import date, timedelta

import pandas as pd
import pdfplumber
import requests
import urllib3

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

urllib3.disable_warnings()
LIST = "https://petrobangla.org.bd/pages/reports"
FILTER = json.dumps({"reports_type": "6922d2b181fc96cef9e99f16"})
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 120)
DATA_START = date(2021, 1, 1)
REVISION_DAYS = 3
BATCH = 150
MAX_PAGES = 400
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "bangladesh_gas.xlsx")
BN = str.maketrans("০১২৩৪৫৬৭৮৯", "0123456789")
MON = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov",
                                   "dec"], start=1)}
COLS = ["Prod_BGFCL", "Prod_SGFL", "Prod_BAPEX", "Prod_state", "Prod_IOC", "Bibiyana", "Jalalabad", "Moulavibazar",
        "Bangora", "RLNG", "Total_supply", "Power_demand", "Power_supply", "Fertiliser_demand", "Fertiliser_supply",
        "Others_supply", "Total_distribution", "Power_nongrid"]


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    for i in range(4):
        try:
            r = requests.get(url, headers=H, timeout=T, verify=False, **kw)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            if i == 3:
                raise
            out(f"  retry {url[:90]}: {e}")


def listing_page(page):
    """(date label, pdf url) for one listing page, newest first. The label is the report date (dd.mm.yy, often
    in Bengali digits); HTML comments holding the upload date are stripped first."""
    r = get(LIST, params={"filters": FILTER, "page": page})
    html = re.sub(r"<!--.*?-->", " ", r.text, flags=re.S)
    rows = []
    for chunk in re.split(r"<tr[\s>]", html)[1:]:
        m = re.search(r'https://objectstorage[^"\']+\.pdf', chunk)
        if not m:
            continue
        text = re.sub(r"<[^>]+>|\s+", " ", chunk).translate(BN)
        d = re.search(r"\b(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})\b", text)
        when = None
        if d:
            y = int(d.group(3))
            y = y + 2000 if y < 100 else y
            try:
                when = date(y, int(d.group(2)), int(d.group(1)))
            except ValueError:
                when = None
        rows.append((when, m.group(0)))
    # labels are typed by hand: one far from the rest of its page is a typo - treat it as unknown (the PDF's own
    # date is used once downloaded)
    dates = sorted(d for d, _ in rows if d)
    if dates:
        mid = dates[len(dates) // 2]
        rows = [(d if d and abs((d - mid).days) <= 20 else None, u) for d, u in rows]
    return rows


def report_date(text):
    """'Date : 01-02 Oct, 2026' / '30 Sep- 1 Oct 2026' / '29 - 30 Sep, 2026' -> the end of the gas day."""
    m = re.search(r"Date\s*:\s*(.{4,40}?)\s+From", text)
    if not m:
        return None
    s = m.group(1)
    yr = re.findall(r"(20\d\d)", s)
    mons = [MON[x.lower()[:3]] for x in re.findall(r"[A-Za-z]{3,9}", s) if x.lower()[:3] in MON]
    days = [int(x) for x in re.findall(r"\b(\d{1,2})\b", s)]
    if not (yr and mons and days):
        return None
    try:
        return date(int(yr[-1]), mons[-1], days[-1])
    except ValueError:
        return None


def nums_after(line_pat, text, k):
    """k-th number (0-based) after the first match of line_pat on its line."""
    m = re.search(line_pat + r"([^\n]*)", text)
    if not m:
        return None
    v = re.findall(r"-?\d+(?:\.\d+)?", m.group(m.lastindex))
    return float(v[k]) if len(v) > k else None


def _doubled(tok):
    return len(tok) >= 2 and len(tok) % 2 == 0 and tok[::2] == tok[1::2]


def undouble(text):
    """Bold rows in the older (2021 - early 2022) reports come out of the PDF with every character doubled
    ('SSuubb--TToottaall 4444 885511 666600..55' = 'Sub-Total 44 851 660.5'). On a line where most tokens are
    doubled, collapse them; other lines (where '44' is a real number) are left alone."""
    lines = []
    for line in text.split("\n"):
        toks = line.split(" ")
        long = [x for x in toks if len(x) >= 2]
        if long and sum(_doubled(x) for x in long) >= 0.6 * len(long):
            line = " ".join(x[::2] if _doubled(x) else x for x in toks)
        lines.append(line)
    return "\n".join(lines)


def parse(text):
    text = undouble(text)
    prod, _, dist = text.partition("II. Distribution")
    row = {}
    subs = [re.findall(r"-?\d+(?:\.\d+)?", s) for s in re.findall(r"Sub-Total(?!\s*\()([^\n]*)", prod)]
    # Sub-Total lines in order: BGFCL, SGFL, BAPEX, IOCs, RPGCL; columns wells, capacity, production, ...
    for key, s in zip(("Prod_BGFCL", "Prod_SGFL", "Prod_BAPEX", "Prod_IOC", "RLNG_sub"), subs):
        row[key] = float(s[2]) if len(s) > 2 else None
    row["Prod_state"] = nums_after(r"Sub-Total\s*\(1\+2\+3\)", prod, 2)
    row["RLNG"] = nums_after(r"R-?LNG\)", prod, 1)
    if row["RLNG"] is None:
        row["RLNG"] = row.get("RLNG_sub")
    row.pop("RLNG_sub", None)
    for f, pat in (("Bibiyana", r"Bibiyana"), ("Jalalabad", r"Jalalabad"), ("Moulavibazar", r"Ma?o?u?lavibazar"),
                   ("Bangora", r"Bangora")):
        row[f] = nums_after(pat, prod, 2)
    row["Total_supply"] = nums_after(r"Grand Total[^:\n]*:", prod, 2)
    tot = re.findall(r"Total\s*:\s*([^\n]*)", dist)
    v = re.findall(r"-?\d+(?:\.\d+)?", tot[-1]) if tot else []
    if len(v) >= 6:
        (row["Power_demand"], row["Power_supply"], row["Fertiliser_demand"], row["Fertiliser_supply"],
         row["Others_supply"], row["Total_distribution"]) = map(float, v[:6])
    # older reports list gas to non-grid (captive / off-grid) power separately, outside 'Total :'
    row["Power_nongrid"] = nums_after(r"Total Non-Grid Power", dist, 0)
    # the state companies add up to (1+2+3); state + IOCs + R-LNG add up to the grand total: drop what doesn't
    comp = [row.get(k) for k in ("Prod_BGFCL", "Prod_SGFL", "Prod_BAPEX")]
    if row.get("Prod_state") and all(c is not None for c in comp) and abs(sum(comp) - row["Prod_state"]) > 2:
        row["Prod_BGFCL"] = row["Prod_SGFL"] = row["Prod_BAPEX"] = None
    parts = [row.get(k) for k in ("Prod_state", "Prod_IOC", "RLNG")]
    if row.get("Total_supply") and all(p is not None for p in parts) and abs(sum(parts) - row["Total_supply"]) > 3:
        row["Prod_state"] = row["Prod_IOC"] = row["Prod_BGFCL"] = row["Prod_SGFL"] = row["Prod_BAPEX"] = None
    return row


def fetch(item):
    label, url = item
    try:
        r = get(url)
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            text = "\n".join((p.extract_text() or "") for p in pdf.pages[:2])
        when = report_date(text) or label
        row = parse(text)
        # sanity: the three supply blocks add up to the grand total
        parts = [row.get(k) for k in ("Prod_state", "Prod_IOC", "RLNG")]
        if row.get("Total_supply") and all(p is not None for p in parts):
            if abs(sum(parts) - row["Total_supply"]) > 0.03 * row["Total_supply"]:
                out(f"  {when}: parts {sum(parts):.1f} vs total {row['Total_supply']}: kept, check")
        return when, row
    except Exception as e:  # noqa: BLE001
        out(f"  {label} {url[-40:]}: {type(e).__name__}: {e}")
        return label, None


def read_saved(path):
    try:
        d = pd.read_excel(path, sheet_name="Daily", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    d.index = pd.to_datetime(d.index)
    return d.sort_index()


def save(path, new_rows):
    old = read_saved(path)
    new = pd.DataFrame.from_dict(new_rows, orient="index")
    if not new.empty:
        new.index = pd.to_datetime(new.index)
    d = new if old.empty else (old if new.empty else pd.concat([old[~old.index.isin(new.index)], new]))
    d = d.sort_index()
    d = d[[c for c in COLS if c in d] + [c for c in d if c not in COLS]].round(1)
    d.index.name = "date"
    notes = [
        "UNITS",
        "MMCFD = million cubic feet per day over the gas day (08:00 to 08:00); the date is the end of the gas day, as "
        "Petrobangla labels its report. 1 MMCFD = 0.0283 million m3 per day (the charts show mcm/d).",
        "Supply: Prod_BGFCL / Prod_SGFL / Prod_BAPEX = state companies' production, Prod_state = their sum; Prod_IOC = "
        "international oil companies (Chevron: Bibiyana, Jalalabad, Moulavibazar; Tullow: Bangora - each also shown); "
        "RLNG = regasified LNG delivered by RPGCL from the FSRUs (imports); Total_supply = Petrobangla's grand total.",
        "Distribution: Power_demand / Power_supply = gas demanded by and supplied to power plants; Fertiliser_demand "
        "(maximum) / Fertiliser_supply; Others_supply = everything else supplied by the distribution companies "
        "(industry, captive power, CNG, commercial, households); Total_distribution. Power_nongrid = gas to non-grid "
        "power listed separately in the older reports (2021 - early 2022), not in Power_supply.",
        "Checks: the state companies must add up to their (1+2+3) sub-total and state + IOCs + R-LNG to the grand total "
        "(within a few MMCFD); a report that fails keeps its totals but its company split is left blank.",
        "",
        "COVERAGE",
        f"Daily from {d.index.min():%Y-%m-%d} to {d.index.max():%Y-%m-%d} ({len(d)} days; Petrobangla occasionally "
        "skips a day). Pulled from " + f"{DATA_START:%Y-%m-%d}.",
        "",
        "SOURCE",
        "Petrobangla (Bangladesh Oil, Gas and Mineral Corporation), Production & Marketing Division, Daily Gas & "
        f"Condensate Production and Distribution Report: {LIST}",
    ]
    xlsx_notes.write_workbook(path, {"Daily": d}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {path}: {len(d)} days {d.index.min():%Y-%m-%d}..{d.index.max():%Y-%m-%d}")
    return d


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    saved = read_saved(args.out)
    have = set(saved.dropna(subset=["Total_supply"]).index.date) if "Total_supply" in saved else set()
    today = date.today()
    revise = {today - timedelta(days=k) for k in range(REVISION_DAYS + 1)}
    missing = {DATA_START + timedelta(days=k) for k in range((today - DATA_START).days + 1)} - (have - revise)
    earliest = min(missing) if missing else today
    out(f"{len(have)} days saved; {len(missing)} candidate days, earliest {earliest}")
    todo = []
    for page in range(1, MAX_PAGES + 1):
        rows = listing_page(page)
        if not rows:
            break
        todo += [(d, u) for d, u in rows if d is None or (d >= DATA_START and d in missing)]
        dates = [d for d, _ in rows if d]
        if dates and max(dates) < earliest:   # the whole page is older than anything still missing
            break
    out(f"{len(todo)} reports to download (listing read to page {page})")
    done = {}
    for b in range(0, len(todo), BATCH):
        with ThreadPoolExecutor(max_workers=6) as ex:
            for when, row in ex.map(fetch, todo[b:b + BATCH]):
                if when and row and row.get("Total_supply"):
                    done[when] = row
        save(args.out, done)
    if not todo:
        save(args.out, done)
    d = read_saved(args.out)
    if d.empty:
        raise SystemExit("No Petrobangla reports parsed")
    out(d.tail(4).to_string())


if __name__ == "__main__":
    main()
