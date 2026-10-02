"""
Guyana natural gas (Stabroek block, all associated gas): daily production and what happens to it - reinjected,
used as fuel offshore, flared, and (once the Gas-to-Energy pipeline runs) anything else the Ministry reports.

Source: Ministry of Natural Resources, Petroleum Management Programme - Data Centre chart pages
  https://petroleum.gov.gy/data-chart/gas-injected-flared-and-used/   (and /gas-produced/)
Each page embeds the whole daily series (from 20-Dec-2019) as JSON for its amCharts chart: 'Daily Reported Produced
Gas (kscf)', 'Daily Reported Injected Gas (kscf)', 'Gas Flared (kscf)', 'Gas Used for Fuel (kscf)'. The Ministry
publishes with about a month's lag, after auditing Esso E&P Guyana's monthly production reports. kscf = thousand
standard cubic feet; converted at 1 scf = 0.0283168 m3. Found via discovery_archive/south_america/GUYANA_GAS_DISCOVERY*.py.

Incremental: the page always carries the full history in one file, so it is downloaded and parsed only when its
stated period end ('20 December 2019 - <end>') is later than the workbook's last day (recorded on the Units sheet).

Usage: python3 GUYANA_GAS.py [--out PATH] [--force]
"""
import argparse
import json
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import xlsx_notes  # noqa: E402

OUT = os.path.join("output", "Data and Chart Outputs", "guyana_gas.xlsx")
PAGES = ["https://petroleum.gov.gy/data-chart/gas-injected-flared-and-used/",
         "https://petroleum.gov.gy/data-chart/gas-produced/"]
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                    "Safari/537.36"}
KSCF_TO_MCM = 1000 * 0.0283168 / 1e6      # thousand scf -> million m3
NAMES = {"Daily Reported Produced Gas (kscf)": "Produced", "Daily Reported Injected Gas (kscf)": "Reinjected",
         "Gas Used for Fuel (kscf)": "Used_as_fuel", "Gas Flared (kscf)": "Flared"}


def page_end(html):
    m = re.search(r"20 December 2019\s*(?:&#8211;|–|-)\s*(\d{1,2} \w+ \d{4})", html)
    return pd.Timestamp(m.group(1)) if m else None


def parse(html):
    """Every embedded data array: var _data<n> = '[{...}]' -> one frame of the kscf columns, daily."""
    frames = []
    for m in re.finditer(r"var\s+_data\d*\s*=\s*'(\[.*?\])'\s*;", html, re.S):
        rows = json.loads(m.group(1))
        df = pd.DataFrame(rows)
        if "Date" not in df:
            continue
        df["date"] = pd.to_datetime(df["Date"], format="%d-%b-%Y", errors="coerce")
        cols = [c for c in df.columns if str(c).endswith("(kscf)")]
        if cols:
            frames.append(df.dropna(subset=["date"]).set_index("date")[cols].apply(pd.to_numeric, errors="coerce"))
    if not frames:
        raise RuntimeError("no embedded data array found - the page layout changed")
    return combine(frames)


def combine(frames):
    """Arrays split the history by period and share column names: merge them cell by cell (first non-blank)."""
    out = frames[0]
    for f in frames[1:]:
        out = out.combine_first(f)
    return out.groupby(level=0).first().sort_index()


def to_mcm(kscf):
    d = kscf.rename(columns=lambda c: NAMES.get(c, re.sub(r"\W+", "_", c.replace("(kscf)", "")).strip("_"))) * KSCF_TO_MCM
    d.columns = [f"{c}_mcm_per_day" for c in d.columns]
    disp = [c for c in d.columns if not c.startswith("Produced")]
    if "Produced_mcm_per_day" in d and disp:
        # produced minus the reported uses: gas sent onshore (Gas-to-Energy) once that starts, plus rounding
        d["Other_or_unreported_mcm_per_day"] = d["Produced_mcm_per_day"] - d[disp].sum(axis=1, min_count=1)
    return d.round(4)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--force", action="store_true", help="re-read the page even if it shows no newer period")
    args = ap.parse_args()
    old = pd.DataFrame()
    if os.path.exists(args.out):
        old = pd.read_excel(args.out, sheet_name="Daily", index_col=0)
        old.index = pd.to_datetime(old.index)
    s = requests.Session()
    s.headers.update(UA)
    raws, end = [], None
    for url in PAGES:
        r = s.get(url, timeout=120)
        r.raise_for_status()
        end = end or page_end(r.text)
        if not old.empty and end is not None and end <= old.index.max() and not args.force:
            print(f"page period ends {end:%Y-%m-%d}, workbook already to {old.index.max():%Y-%m-%d}: nothing new",
                  flush=True)
            return
        raws.append(parse(r.text))
    kscf = combine(raws)
    daily = to_mcm(kscf)
    if not old.empty:   # keep any day the page no longer carries
        daily = daily.combine_first(old)
    daily.index.name = "date"
    monthly = daily.resample("MS").mean().round(3)
    monthly["Reinjected_share_pct"] = (100 * monthly["Reinjected_mcm_per_day"] / monthly["Produced_mcm_per_day"]).round(1)
    monthly["Days"] = daily["Produced_mcm_per_day"].resample("MS").count()
    lines = ["GUYANA NATURAL GAS (Stabroek block, associated gas)",
             "Source: Ministry of Natural Resources, Petroleum Management Programme - Data Centre "
             "(https://petroleum.gov.gy/data-chart/gas-injected-flared-and-used/), daily, from the audited monthly "
             "production reports of Esso Exploration and Production Guyana Ltd; published about a month in arrears.",
             f"Period on the page: 20-Dec-2019 to {end:%d-%b-%Y}" if end is not None else "Period: see Daily",
             "",
             "UNITS",
             "million m3/day (from the Ministry's kscf/day at 1 scf = 0.0283168 m3; 1 mcm/d = 35.3 MMscf/d). "
             "Produced = gross associated gas. Reinjected = gas returned to the reservoir (pressure support / storage) - "
             "the largest use; shown as its own series. Used_as_fuel = offshore power and compression. Flared. "
             "Other_or_unreported = produced minus those uses: gas sent onshore through the Gas-to-Energy pipeline "
             "once it flows (the Ministry adds a column for it when reported), otherwise rounding.",
             "Monthly = mean of the days published in the month; 'Days' counts them.",
             "",
             "UPDATES",
             "The page carries the whole history; it is read only when its period end moves past the workbook's last "
             "day (guyana_gas.yml, 1st and 15th of the month)."]
    xlsx_notes.write_workbook(args.out, {"Monthly": monthly, "Daily": daily}, lines, ["UNITS", "UPDATES"])
    last = monthly.dropna(subset=["Produced_mcm_per_day"]).tail(1)
    print(f"saved {args.out}: {len(daily)} days to {daily.index.max():%Y-%m-%d}; columns {list(daily.columns)}", flush=True)
    print(last.round(2).to_string(), flush=True)


if __name__ == "__main__":
    main()
