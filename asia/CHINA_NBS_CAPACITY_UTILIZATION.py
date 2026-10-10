"""
Pull China NBS's quarterly industrial capacity utilisation rate
(全国规模以上工业产能利用率) by industry, from 2021 on: total industry,
mining, manufacturing, utilities, and ~15 industries incl. coal mining,
oil and gas extraction, chemicals, non-metallic minerals (cement/glass),
ferrous (steel) and non-ferrous smelting.

Each quarterly release has a table [industry | quarter rate % | change
vs a year earlier (pp) | (year-to-date rate % | change pp)]; only the
quarter's rate is kept. Chinese release list www.stats.gov.cn/sj/zxfb/
(see china_nbs_common.py); 2021 Q1-Q3 are no longer on the list but
still online (EXTRA_RELEASES, found by CHINA_NBS_DISCOVERY5.py).

Incremental: quarters already in the workbook are not re-fetched.

    python3 asia/CHINA_NBS_CAPACITY_UTILIZATION.py --out "output/Data and Chart Outputs/china_nbs_capacity_utilization_quarterly.xlsx"
"""

import argparse
import os
import re
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))                   # asia/, for china_nbs_common
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import china_nbs_archive_ids as arch  # noqa: E402
import china_nbs_common as nbs  # noqa: E402
import xlsx_notes  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "china_nbs_capacity_utilization_quarterly.xlsx")
TITLE_RE = r"\d{4}年[一二三四]季度全国(规模以上)?工业产能利用率"
HISTORY_START = "2013-01-01"
EXTRA_RELEASES = arch.CAPACITY   # releases no longer on the list, from the ID scan
QUARTER = {"一": 1, "二": 4, "三": 7, "四": 10}

# (Chinese industry name, column, English label, chart group)
INDUSTRIES = [
    ("工业", "Industry_Total", "Industry (total)", "industrial capacity utilisation"),
    ("采矿业", "Mining", "Mining", "industrial capacity utilisation"),
    ("制造业", "Manufacturing", "Manufacturing", "industrial capacity utilisation"),
    ("电力、热力、燃气及水生产和供应业", "Utilities", "Power, heat, gas and water supply", "industrial capacity utilisation"),
    ("煤炭开采和洗选业", "Coal_Mining", "Coal mining and washing", "energy and heavy industry capacity utilisation"),
    ("石油和天然气开采业", "Oil_Gas_Extraction", "Oil and gas extraction", "energy and heavy industry capacity utilisation"),
    ("化学原料和化学制品制造业", "Chemicals", "Chemical raw materials and products", "energy and heavy industry capacity utilisation"),
    ("化学纤维制造业", "Chemical_Fibres", "Chemical fibres", "energy and heavy industry capacity utilisation"),
    ("非金属矿物制品业", "Non_Metallic_Minerals", "Non-metallic mineral products (cement, glass)",
     "energy and heavy industry capacity utilisation"),
    ("黑色金属冶炼和压延加工业", "Ferrous_Metals", "Ferrous metal smelting and rolling (steel)", "energy and heavy industry capacity utilisation"),
    ("有色金属冶炼和压延加工业", "Non_Ferrous_Metals", "Non-ferrous metal smelting and rolling",
     "energy and heavy industry capacity utilisation"),
    ("食品制造业", "Food", "Food manufacturing", "other manufacturing capacity utilisation"),
    ("纺织业", "Textiles", "Textiles", "other manufacturing capacity utilisation"),
    ("医药制造业", "Pharmaceuticals", "Pharmaceuticals", "other manufacturing capacity utilisation"),
    ("通用设备制造业", "General_Equipment", "General-purpose equipment", "other manufacturing capacity utilisation"),
    ("专用设备制造业", "Special_Equipment", "Special-purpose equipment", "other manufacturing capacity utilisation"),
    ("汽车制造业", "Automobiles", "Automobiles", "other manufacturing capacity utilisation"),
    ("电气机械和器材制造业", "Electrical_Machinery", "Electrical machinery and equipment", "other manufacturing capacity utilisation"),
    ("计算机、通信和其他电子设备制造业", "Electronics", "Computers, communication and electronic equipment",
     "other manufacturing capacity utilisation"),
]
BY_NAME = {cn: col for cn, col, *_ in INDUSTRIES}
BY_NAME.update({"规模以上工业": "Industry_Total", "全国规模以上工业": "Industry_Total", "全国工业": "Industry_Total"})


def period_of(title):
    m = re.search(r"(\d{4})年([一二三四])季度", title)
    return pd.Timestamp(int(m.group(1)), QUARTER[m.group(2)], 1) if m else None


def parse_release(html):
    row, unmatched = {}, []
    for cells in nbs.table_rows(html):
        if len(cells) < 2:
            continue
        name = re.sub(r"^其中[:：]", "", cells[0])
        value = nbs.num(cells[1])
        if value is None:
            continue
        col = BY_NAME.get(name)
        if col is None:
            unmatched.append(name)
        elif col not in row:
            row[col] = value
            for suf, i in (("YoY_pp", 2), ("YTD", 3), ("YTD_YoY_pp", 4)):
                v = nbs.num(cells[i]) if len(cells) > i else None
                if v is not None:
                    row[f"{col}_{suf}"] = v
    return row, unmatched


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()
    base = [i[1] for i in INDUSTRIES]
    cols = base + [f"{c}_{s}" for c in base for s in ("YoY_pp", "YTD", "YTD_YoY_pp")]
    data = nbs.read_sheet(args.out, "Data") if nbs.has_sheet(args.out, "Series") else None
    data = pd.DataFrame(columns=cols, dtype=float) if data is None else data.reindex(columns=cols)
    held = set(data.index[data["Industry_Total_YoY_pp"].notna()]) if len(data) else set()   # held once the y/y columns are in
    gaps = nbs.recent_gaps(held, "QS")
    deep = (not held) or bool(gaps)

    def wanted(title):
        p = period_of(title)
        return p is not None and p >= pd.Timestamp(HISTORY_START) and p not in held

    releases = nbs.crawl_index(TITLE_RE, wanted, deep, EXTRA_RELEASES, stop_after_known=2, near_re=r"产能利用")
    nbs.log(f"  held {len(held)} quarters; {len(releases)} release(s) to fetch")
    new, unmatched_all = {}, set()
    for title, url in releases:
        p = period_of(title)
        try:
            html = nbs.fetch(url)
        except Exception as e:  # noqa: BLE001 - next run retries
            nbs.log(f"  [{p:%Y-%m}] fetch failed ({type(e).__name__}) - skipped")
            continue
        row, unmatched = parse_release(html or "")
        unmatched_all.update(unmatched)
        m = re.search(r"产能利用率为([\d.]+)%", title)
        if row and "Industry_Total" not in row and m:   # some releases label the total row differently
            row["Industry_Total"] = float(m.group(1))
        if not row:
            nbs.log(f"  [{p:%Y-%m}] no rates parsed - skipped ({url})")
            continue
        new[p] = row
        nbs.log(f"  [{p:%Y} Q{(p.month - 1) // 3 + 1}] {len(row)} industries, total {row.get('Industry_Total')}")
    if unmatched_all:
        nbs.log(f"  rows not mapped (left out): {sorted(unmatched_all)}")
    if new:
        add = pd.DataFrame.from_dict(new, orient="index")
        data = data.combine_first(add) if len(data) else add
    if data.dropna(how="all").empty:
        raise SystemExit("No data at all - nothing to save.")
    data = data.reindex(columns=cols).sort_index()
    data.index.name = "quarter_start"
    rows = [(col, label, "%", group, "line") for _cn, col, label, group in INDUSTRIES]
    for _cn, col, label, _g in INDUSTRIES:
        rows += [(f"{col}_YoY_pp", f"{label}, change vs a year earlier", "percentage points", "", "line"),
                 (f"{col}_YTD", f"{label}, year-to-date rate", "%", "", "line"),
                 (f"{col}_YTD_YoY_pp", f"{label}, year-to-date change vs a year earlier", "percentage points", "", "line")]
    series = nbs.series_sheet(rows)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": data, "Series": series}, NOTES_LINES, NOTES_SECTION_TITLES)
    print(f"Saved {len(data)} quarter(s) ({data.index.min():%Y-%m}..{data.index.max():%Y-%m}), {len(new)} new, "
          f"to {args.out}")


NOTES_LINES = [
    "UNITS",
    "Capacity utilisation rate, % of capacity, for the quarter (rows dated the first day of the quarter). "
    "Full industry names are on the 'Series' sheet.",
    "",
    "SCOPE",
    "Industrial enterprises above designated size (annual main business revenue of RMB 20 million or more). "
    "NBS's releases for 2021 are titled '全国工业产能利用率' (same survey).",
    "",
    "SOURCE",
    "National Bureau of Statistics of China, quarterly release 'YYYY年X季度全国规模以上工业产能利用率为X%', "
    "Chinese release list https://www.stats.gov.cn/sj/zxfb/. Beside the quarter's rate: <name>_YoY_pp (change vs the "
    "same quarter a year earlier, percentage points), <name>_YTD (rate for the year to date: first half, first three "
    "quarters, full year) and <name>_YTD_YoY_pp, all as published. NBS's quarterly releases found online start with "
    "2017 Q4.",
    "",
    "UPDATES",
    "Incremental: quarters already held are not re-fetched.",
]
NOTES_SECTION_TITLES = {"UNITS", "SCOPE", "SOURCE", "UPDATES"}


if __name__ == "__main__":
    main()
