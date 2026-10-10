"""
Pull China's monthly ENERGY production from the National Bureau of
Statistics (NBS), industrial enterprises above designated size, from
2021 on: raw coal, coke, crude oil, crude oil processed (refinery runs),
natural gas, and electricity generation - total and by source (thermal,
hydro, nuclear, wind, solar).

Previously parsed from the English "Energy Production in {Month}
{Year}" press release (prose, rounded figures, English list only reaches
back to Apr 2024, no generation by source). Now read from the Chinese
monthly industrial production release's product-output table - same
NBS numbers, unrounded, with generation by source and history back to
late 2021. NBS's Chinese "能源生产情况" release itself has no table (prose
with the same monthly figures plus daily averages = monthly / days in
month), so it adds nothing pullable. See china_nbs_common.py and
china_nbs_output.py for the source and parsing.

    python3 asia/CHINA_NBS_ENERGY_PRODUCTION.py --out "output/Data and Chart Outputs/china_nbs_energy_production_monthly.xlsx"
"""

import argparse
import os
import re
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))                   # asia/, for china_nbs_*
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import china_nbs_common as nbs  # noqa: E402
import china_nbs_output as out  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "china_nbs_energy_production_monthly.xlsx")
COLUMNS = ["Raw_Coal_Mt", "Coke_Mt", "Crude_Oil_Mt", "Crude_Oil_Processing_Mt", "Natural_Gas_Bcm",
           "Total_Generation_TWh", "Thermal_Generation_TWh", "Hydro_Generation_TWh", "Nuclear_Generation_TWh",
           "Wind_Generation_TWh", "Solar_Generation_TWh"]

NOTES_LINES = [
    "UNITS",
    "Every value is that calendar month's output. " + out.units_line(COLUMNS) + ". "
    "Daily averages (as quoted in NBS's energy production release) = monthly value / days in the month.",
    "",
    "IMPORTS (sheets 'Imports', 'Imports Jan-Feb', 'Imports checked')",
    "Coal, crude oil and natural gas imports (Mt, weight; natural gas is NOT converted to volume) as quoted in the prose of "
    "NBS's monthly 能源生产情况 release from customs flash data, with their y/y %. NBS quoted them from the March 2022 "
    "release to December 2024 (not in the September 2022 release) and dropped them from the 2025 releases; the GACC "
    "customs workbook is the charted import series. 'Imports checked' lists the releases already read.",
    "",
] + out.COMMON_NOTES
NOTES_SECTION_TITLES = {"UNITS", "IMPORTS (sheets 'Imports', 'Imports Jan-Feb', 'Imports checked')", "SCOPE", "JANUARY AND FEBRUARY", "Y/Y AND YEAR-TO-DATE COLUMNS", "SOURCE", "UPDATES"}


# Imports quoted in the prose of the 能源生产情况 release (customs flash data), Mar 2022 - Dec 2024 (NBS dropped them from the
# 2025 releases and left them out of Sep 2022): month value (万吨 -> Mt) and its y/y. Data only; the master charts the
# GACC customs workbook instead.
ENERGY_TITLE_RE = r"\d{4}年((\d{1,2}|1[—\-－～~]\d{1,2})月份?|上半年|前三季度)能源生产情况"
IMPORTS = [("Coal_Imports_Mt", r"进口煤炭([\d.]+)(万吨|亿吨)，同比(增长|下降)([\d.]+)%", "Coal imports"),
           ("Crude_Oil_Imports_Mt", r"进口原油([\d.]+)(万吨|亿吨)，同比(增长|下降)([\d.]+)%", "Crude oil imports"),
           ("Natural_Gas_Imports_Mt", r"进口天然气([\d.]+)(万吨|亿吨)，同比(增长|下降)([\d.]+)%", "Natural gas imports (weight)")]
IMPORT_COLS = [c for c, *_ in IMPORTS] + [f"{c[:-3]}_YoY_pct" for c, *_ in IMPORTS]


def parse_imports(html):
    """{column: value} from the prose: the FIRST mention of each import (the month, or Jan-Feb in the combined release)."""
    text = re.sub(r"<[^>]+>", "", html)
    row = {}
    for col, rx, _label in IMPORTS:
        m = re.search(rx, text)
        if not m:
            continue
        v = float(m.group(1)) * (0.01 if m.group(2) == "万吨" else 100.0)
        row[col] = round(v, 4)
        row[f"{col[:-3]}_YoY_pct"] = float(m.group(4)) * (1 if m.group(3) == "增长" else -1)
    return row


def imports_extra(out_path):
    """Hook for out.pull: reads the held 'Imports' sheets, fetches only energy releases not checked yet."""
    def hook(_data, _janfeb):
        imp = nbs.read_sheet(out_path, "Imports") if nbs.has_sheet(out_path, "Imports") else None
        imp_jf = nbs.read_sheet(out_path, "Imports Jan-Feb") if nbs.has_sheet(out_path, "Imports Jan-Feb") else None
        chk = nbs.read_sheet(out_path, "Imports checked") if nbs.has_sheet(out_path, "Imports checked") else None
        imp = pd.DataFrame(columns=IMPORT_COLS, dtype=float) if imp is None else imp.reindex(columns=IMPORT_COLS)
        imp_jf = pd.DataFrame(columns=IMPORT_COLS, dtype=float) if imp_jf is None else imp_jf.reindex(columns=IMPORT_COLS)
        chk = pd.DataFrame(columns=["imports_found"], dtype=float) if chk is None else chk.reindex(columns=["imports_found"])
        checked_ym = {(p.year, p.month) for p in chk.index}

        def wanted(title):
            period, is_jf = nbs.month_period(title)
            if period is None or period < pd.Timestamp(out.HISTORY_START):
                return False
            return (period.year, period.month) not in checked_ym

        deep = len(chk) == 0
        releases = nbs.crawl_index(ENERGY_TITLE_RE, wanted, deep, out.ENERGY_EXTRA, stop_after_known=3)
        nbs.log(f"  energy-production releases to read for imports: {len(releases)}")
        new, new_jf, new_chk = {}, {}, {}
        for title, url in releases:
            period, is_jf = nbs.month_period(title)
            try:
                html = nbs.fetch(url)
            except Exception as e:  # noqa: BLE001 - next run retries
                nbs.log(f"  [{period:%Y-%m}] energy release fetch failed ({type(e).__name__}) - skipped")
                continue
            row = parse_imports(html or "")
            (new_jf if is_jf else new)[period] = row
            new_chk[period] = 1.0 if row else 0.0
            nbs.log(f"  [{period:%Y-%m}{' Jan-Feb' if is_jf else ''}] imports {'found' if row else 'not in this release'}")
        def merge(old, add):
            add = pd.DataFrame.from_dict({k: v for k, v in add.items() if v}, orient="index")
            if add.empty:
                return old
            return add.combine_first(old) if len(old) else add
        imp, imp_jf = merge(imp, new), merge(imp_jf, new_jf)
        if new_chk:
            c2 = pd.DataFrame({"imports_found": pd.Series(new_chk)})
            chk = c2.combine_first(chk) if len(chk) else c2
        for df in (imp, imp_jf, chk):
            df.index.name = "month"
        return {"Imports": imp.reindex(columns=IMPORT_COLS).sort_index(),
                "Imports Jan-Feb": imp_jf.reindex(columns=IMPORT_COLS).sort_index(),
                "Imports checked": chk.sort_index()}
    return hook


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()
    out.pull(args.out, COLUMNS, NOTES_LINES, NOTES_SECTION_TITLES, with_va=out.ENERGY_VA,
             extra=imports_extra(args.out))


if __name__ == "__main__":
    main()
