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

import china_nbs_common as nbs  # noqa: E402
import xlsx_notes  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "china_nbs_capacity_utilization_quarterly.xlsx")
TITLE_RE = r"\d{4}年[一二三四]季度全国(规模以上)?工业产能利用率"
HISTORY_START = "2021-01-01"
_OLD = "https://www.stats.gov.cn/sj/zxfb/202302/t20230203_{}.html"
EXTRA_RELEASES = [(t, _OLD.format(i)) for t, i in [
    ("2021年一季度全国工业产能利用率为77.2%", 1901053),
    ("2021年二季度全国工业产能利用率为78.4%", 1901161),
    ("2021年三季度全国工业产能利用率为77.1%", 1901247),
]]
QUARTER = {"一": 1, "二": 4, "三": 7, "四": 10}

# (Chinese industry name, column, English label, chart group)
INDUSTRIES = [
    ("工业", "Industry_Total", "Industry (total)", "industrial capacity utilisation"),
    ("采矿业", "Mining", "Mining", "industrial capacity utilisation"),
    ("制造业", "Manufacturing", "Manufacturing", "industrial capacity utilisation"),
    ("电力、热力、燃气及水生产和供应业", "Utilities", "Power, heat, gas and water supply", "industrial capacity utilisation"),
    ("煤炭开采和洗选业", "Coal_Mining", "Coal mining and washing", "capacity utilisation - energy and heavy industry"),
    ("石油和天然气开采业", "Oil_Gas_Extraction", "Oil and gas extraction", "capacity utilisation - energy and heavy industry"),
    ("化学原料和化学制品制造业", "Chemicals", "Chemical raw materials and products", "capacity utilisation - energy and heavy industry"),
    ("化学纤维制造业", "Chemical_Fibres", "Chemical fibres", "capacity utilisation - energy and heavy industry"),
    ("非金属矿物制品业", "Non_Metallic_Minerals", "Non-metallic mineral products (cement, glass)",
     "capacity utilisation - energy and heavy industry"),
    ("黑色金属冶炼和压延加工业", "Ferrous_Metals", "Ferrous metal smelting and rolling (steel)", "capacity utilisation - energy and heavy industry"),
    ("有色金属冶炼和压延加工业", "Non_Ferrous_Metals", "Non-ferrous metal smelting and rolling",
     "capacity utilisation - energy and heavy industry"),
    ("食品制造业", "Food", "Food manufacturing", "capacity utilisation - other manufacturing"),
    ("纺织业", "Textiles", "Textiles", "capacity utilisation - other manufacturing"),
    ("医药制造业", "Pharmaceuticals", "Pharmaceuticals", "capacity utilisation - other manufacturing"),
    ("通用设备制造业", "General_Equipment", "General-purpose equipment", "capacity utilisation - other manufacturing"),
    ("专用设备制造业", "Special_Equipment", "Special-purpose equipment", "capacity utilisation - other manufacturing"),
    ("汽车制造业", "Automobiles", "Automobiles", "capacity utilisation - other manufacturing"),
    ("电气机械和器材制造业", "Electrical_Machinery", "Electrical machinery and equipment", "capacity utilisation - other manufacturing"),
    ("计算机、通信和其他电子设备制造业", "Electronics", "Computers, communication and electronic equipment",
     "capacity utilisation - other manufacturing"),
]
BY_NAME = {cn: col for cn, col, *_ in INDUSTRIES}


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
    return row, unmatched


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()
    cols = [i[1] for i in INDUSTRIES]
    data = nbs.read_sheet(args.out, "Data") if nbs.has_sheet(args.out, "Series") else None
    data = pd.DataFrame(columns=cols, dtype=float) if data is None else data.reindex(columns=cols)
    held = set(data.index[data.notna().any(axis=1)]) if len(data) else set()
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
    series = nbs.series_sheet([(col, label, "%", group, "line") for _cn, col, label, group in INDUSTRIES])
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
    "Chinese release list https://www.stats.gov.cn/sj/zxfb/. Only the quarter's rate is kept (the release also "
    "gives the change vs a year earlier and the year-to-date rate).",
    "",
    "UPDATES",
    "Incremental: quarters already held are not re-fetched.",
]
NOTES_SECTION_TITLES = {"UNITS", "SCOPE", "SOURCE", "UPDATES"}


if __name__ == "__main__":
    main()
