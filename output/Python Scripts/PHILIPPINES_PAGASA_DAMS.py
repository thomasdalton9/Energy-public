"""
Philippines dam water levels from PAGASA's daily dam information table (DOST-PAGASA Hydro-Meteorology
Division), https://www.pagasa.dost.gov.ph/flood. Found via discovery_archive/asia/SEA_DISCOVERY_TH_VN_PH.py
and SEA_DISCOVERY2.py.

The flood page carries one table, refreshed each morning, with the 08:00 reading for today and yesterday
for the Luzon reservoirs: Angat (Manila's water supply and 218 MW hydro), Ipo, La Mesa, Ambuklao, Binga,
San Roque, Pantabangan, Magat and Caliraya - reservoir water level (m above sea level), normal high water
level (NHWL), rule curve elevation. PAGASA keeps no history, so this runs DAILY and the workbook is the
history store (both readings on the page are saved, so one missed run leaves no gap).

Writes output/Data and Chart Outputs/philippines_dam_levels.xlsx:
  Daily     date x dam: water level (m)  - <Dam>_m columns
  Limits    per dam: latest NHWL and rule-curve elevation (m)
  Water year charts (Oct-Sep) of the Angat, San Roque, Magat and Pantabangan levels (add_charts.py)

    python3 asia/PHILIPPINES_PAGASA_DAMS.py
"""
import argparse
import html as htmllib
import os
import re
import sys
from datetime import date

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

URL = "https://www.pagasa.dost.gov.ph/flood"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 90)
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "philippines_dam_levels.xlsx")
DAMS = ["Angat", "Ipo", "La Mesa", "Ambuklao", "Binga", "San Roque", "Pantabangan", "Magat", "Caliraya"]
DAY = re.compile(r"^([A-Z][a-z]{2})-(\d{1,2})$")
TIME = re.compile(r"^\d{1,2}:\d\d\s*[ap]m$", re.I)


def out(*a):
    print(*a, flush=True)


def num(x):
    try:
        return float(x.replace(",", ""))
    except ValueError:
        return None


def day_of(m, today):
    d = pd.Timestamp(f"{m.group(1)} {int(m.group(2))} {today.year}")
    return d - pd.DateOffset(years=1) if d.date() > today else d   # Dec readings seen in early January


def parse(page, today):
    """-> (levels: {(date, dam): m}, limits: {dam: (NHWL, rule curve)})"""
    tab = next((t for t in re.findall(r"(?is)<table.*?</table>", page) if "dam name" in t.lower()), None)
    if tab is None:
        raise SystemExit("PAGASA dam table not found")
    cells = [re.sub(r"\s+", " ", htmllib.unescape(re.sub(r"<[^>]+>", " ", c))).strip()
             for c in re.findall(r"(?is)<t[dh][^>]*>(.*?)</t[dh]>", tab)]
    cells = [c for c in cells if c]
    starts = [(i, d) for i, c in enumerate(cells) for d in DAMS if c.lower().startswith(d.lower())]
    levels, limits = {}, {}
    for k, (i, dam) in enumerate(starts):
        seg = cells[i + 1:(starts[k + 1][0] if k + 1 < len(starts) else len(cells))]
        rows, cur = [], None
        for c in seg:   # a reading starts at its time cell and ends at its date cell
            if TIME.match(c):
                cur = []
            elif DAY.match(c) and cur is not None:
                rows.append((day_of(DAY.match(c), today), cur))
                cur = None
            elif cur is not None and num(c) is not None:
                cur.append(num(c))
        for n, (d, vals) in enumerate(rows):
            if vals:
                levels[(d, dam)] = vals[0]
            if n == 0 and len(vals) >= 6:   # today's row: RWL, hours, change, NHWL, dev, rule curve, ...
                limits[dam] = (vals[3], vals[5])
    return levels, limits


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    r = requests.get(URL, headers=H, timeout=T)
    r.raise_for_status()
    today = date.today()
    levels, limits = parse(r.text, today)
    if not levels:
        raise SystemExit("No dam readings parsed")
    new = pd.Series(levels).unstack().reindex(columns=DAMS)
    new.columns = [f"{c.replace(' ', '_')}_m" for c in new.columns]
    out(new.to_string())
    try:
        old = pd.read_excel(args.out, sheet_name="Daily", index_col=0)
        old.index = pd.to_datetime(old.index)
    except (FileNotFoundError, ValueError):
        old = pd.DataFrame()
    # a dam missing from today's table keeps its stored reading: new values win only where present
    daily = new if old.empty else new.combine_first(old).sort_index()
    daily = daily[sorted(daily.columns, key=lambda c: [f"{d.replace(' ', '_')}_m" for d in DAMS].index(c)
                         if c in [f"{d.replace(' ', '_')}_m" for d in DAMS] else 99)]
    daily.index.name = "date"
    lim = pd.DataFrame([{"dam": d, "NHWL_m": v[0], "Rule_curve_m": v[1], "as_of": today.isoformat()}
                        for d, v in limits.items()]).set_index("dam") if limits else pd.DataFrame()
    try:   # a dam whose limits were not parsed today keeps its saved ones
        saved = pd.read_excel(args.out, sheet_name="Limits", index_col=0)
        lim = saved if lim.empty else pd.concat([lim, saved[~saved.index.isin(lim.index)]])
    except (FileNotFoundError, ValueError):
        pass
    notes = [
        "UNITS",
        "Daily: reservoir water level at 08:00, metres above mean sea level, one column per dam. Limits: normal high "
        "water level (NHWL, full) and the day's rule-curve elevation (the operating guide level) from the latest run.",
        "",
        "COVERAGE",
        f"Daily from {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d} - PAGASA posts only today's and "
        "yesterday's readings, so the history here starts with this pull's first run. Luzon reservoirs: Angat (Metro "
        "Manila water supply, 218 MW hydro), Ipo and La Mesa (water supply), Ambuklao, Binga, San Roque (hydro, Agno "
        "river), Pantabangan and Magat (irrigation + hydro), Caliraya (pumped storage).",
        "",
        "SOURCE",
        "DOST-PAGASA, Dam Information (flood page): https://www.pagasa.dost.gov.ph/flood",
    ]
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "Limits": lim}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}: {len(daily)} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d}")


if __name__ == "__main__":
    main()
