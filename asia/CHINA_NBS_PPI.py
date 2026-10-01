"""
Pull China NBS's monthly industrial producer price indices (PPI) from
2021 on: ex-factory prices (headline, producer/consumer goods, mining,
raw materials, processing), purchaser prices (incl. fuel and power,
ferrous, non-ferrous, chemical raw materials, building materials) and
ex-factory prices by industry (coal mining, oil and gas extraction,
petroleum/coal processing, power and heat, gas supply, steel,
non-ferrous, chemicals, ...).

NBS publishes these only as % changes (no index levels): month-on-month
and year-on-year, both kept. Each monthly release's table is [series |
m/m % | y/y % | YTD y/y %]; the column order is read from the header.
Chinese release list www.stats.gov.cn/sj/zxfb/ (see china_nbs_common.py);
Jan-Sep 2021 are no longer on the list but still online
(EXTRA_RELEASES, found by CHINA_NBS_DISCOVERY5.py).

Incremental: months already in the workbook are not re-fetched.

    python3 asia/CHINA_NBS_PPI.py --out "output/Data and Chart Outputs/china_nbs_ppi_monthly.xlsx"
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
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "china_nbs_ppi_monthly.xlsx")
TITLE_RE = r"\d{4}年\d{1,2}月份工业生产者(出厂)?价格"
HISTORY_START = "2021-01-01"
_OLD = "https://www.stats.gov.cn/sj/zxfb/202302/t20230203_{}.html"
EXTRA_RELEASES = [(t, _OLD.format(i)) for t, i in [
    ("2021年4月份工业生产者出厂价格同比上涨6.8% 环比上涨0.9%", 1901078),
]]

# (Chinese series name, column stem, English label, chart group or None = data only)
SERIES = [
    ("工业生产者出厂价格", "PPI", "PPI: ex-factory prices (all industrial products)", "PPI headline y/y"),
    ("生产资料", "PPI_Producer_Goods", "Ex-factory: producer goods", "PPI headline y/y"),
    ("采掘", "PPI_Mining", "Ex-factory: mining (extractive) goods", "PPI headline y/y"),
    ("原材料", "PPI_Raw_Materials", "Ex-factory: raw materials", "PPI headline y/y"),
    ("加工", "PPI_Processing", "Ex-factory: processed goods", "PPI headline y/y"),
    ("生活资料", "PPI_Consumer_Goods", "Ex-factory: consumer goods", "PPI headline y/y"),
    ("工业生产者购进价格", "PPIRM", "Purchaser prices (all)", "PPI headline y/y"),
    ("食品", "PPI_Food", "Ex-factory: food", None),
    ("衣着", "PPI_Clothing", "Ex-factory: clothing", None),
    ("一般日用品", "PPI_Daily_Goods", "Ex-factory: general daily goods", None),
    ("耐用消费品", "PPI_Durables", "Ex-factory: durable consumer goods", None),
    ("燃料、动力类", "PPIRM_Fuel_Power", "Purchaser: fuel and power", "PPI purchaser prices y/y"),
    ("黑色金属材料类", "PPIRM_Ferrous", "Purchaser: ferrous metal materials", "PPI purchaser prices y/y"),
    ("有色金属材料及电线类", "PPIRM_Non_Ferrous", "Purchaser: non-ferrous metals and wire", "PPI purchaser prices y/y"),
    ("化工原料类", "PPIRM_Chemicals", "Purchaser: chemical raw materials", "PPI purchaser prices y/y"),
    ("建筑材料及非金属类", "PPIRM_Building_Materials", "Purchaser: building materials and non-metals",
     "PPI purchaser prices y/y"),
    ("木材及纸浆类", "PPIRM_Timber_Pulp", "Purchaser: timber and pulp", None),
    ("其他工业原材料及半成品类", "PPIRM_Other_Materials", "Purchaser: other industrial materials", None),
    ("农副产品类", "PPIRM_Farm_Products", "Purchaser: farm and sideline products", None),
    ("纺织原料类", "PPIRM_Textile_Materials", "Purchaser: textile raw materials", None),
    ("煤炭开采和洗选业", "Coal_Mining", "Coal mining and washing", "PPI energy industries y/y"),
    ("石油和天然气开采业", "Oil_Gas_Extraction", "Oil and gas extraction", "PPI energy industries y/y"),
    ("石油、煤炭及其他燃料加工业", "Petroleum_Coal_Processing", "Petroleum, coal and other fuel processing",
     "PPI energy industries y/y"),
    ("电力、热力生产和供应业", "Power_Heat", "Power and heat production and supply", "PPI energy industries y/y"),
    ("燃气生产和供应业", "Gas_Supply", "Gas production and supply", "PPI energy industries y/y"),
    ("黑色金属矿采选业", "Ferrous_Ore_Mining", "Ferrous metal ore mining", "PPI metals and materials y/y"),
    ("有色金属矿采选业", "Non_Ferrous_Ore_Mining", "Non-ferrous metal ore mining", "PPI metals and materials y/y"),
    ("黑色金属冶炼和压延加工业", "Ferrous_Smelting", "Ferrous metal smelting and rolling (steel)",
     "PPI metals and materials y/y"),
    ("有色金属冶炼和压延加工业", "Non_Ferrous_Smelting", "Non-ferrous metal smelting and rolling",
     "PPI metals and materials y/y"),
    ("非金属矿物制品业", "Non_Metallic_Minerals", "Non-metallic mineral products (cement, glass)",
     "PPI metals and materials y/y"),
    ("化学原料和化学制品制造业", "Chemicals", "Chemical raw materials and products", "PPI metals and materials y/y"),
    ("化学纤维制造业", "Chemical_Fibres", "Chemical fibres", "PPI metals and materials y/y"),
    ("非金属矿采选业", "Non_Metallic_Mining", "Non-metallic mineral mining", None),
    ("农副食品加工业", "Agri_Food_Processing", "Agricultural food processing", None),
    ("食品制造业", "Food_Manufacturing", "Food manufacturing", None),
    ("酒、饮料和精制茶制造业", "Beverages", "Liquor, beverages and tea", None),
    ("烟草制品业", "Tobacco", "Tobacco", None),
    ("纺织业", "Textiles", "Textiles", None),
    ("纺织服装、服饰业", "Apparel", "Apparel", None),
    ("木材加工和木、竹、藤、棕、草制品业", "Wood_Products", "Wood products", None),
    ("造纸和纸制品业", "Paper", "Paper", None),
    ("印刷和记录媒介复制业", "Printing", "Printing", None),
    ("医药制造业", "Pharmaceuticals", "Pharmaceuticals", None),
    ("橡胶和塑料制品业", "Rubber_Plastics", "Rubber and plastics", None),
    ("金属制品业", "Metal_Products", "Metal products", None),
    ("通用设备制造业", "General_Equipment", "General-purpose equipment", None),
    ("专用设备制造业", "Special_Equipment", "Special-purpose equipment", None),
    ("汽车制造业", "Automobiles", "Automobiles", None),
    ("铁路、船舶、航空航天和其他运输设备制造业", "Other_Transport_Equipment", "Rail, ship, aerospace and other transport",
     None),
    ("电气机械和器材制造业", "Electrical_Machinery", "Electrical machinery", None),
    ("计算机、通信和其他电子设备制造业", "Electronics", "Computers, communication and electronics", None),
    ("水的生产和供应业", "Water_Supply", "Water production and supply", None),
]
BY_NAME = {cn: stem for cn, stem, *_ in SERIES}


def columns():
    return [f"{stem}_{k}" for _cn, stem, *_ in SERIES for k in ("YoY_pct", "MoM_pct")]


def parse_release(html):
    """{column: value} using the header row to find the m/m and y/y columns."""
    row, unmatched = {}, []
    yoy_i = mom_i = None
    for cells in nbs.table_rows(html):
        if any("环比" in c for c in cells) and any("同比" in c for c in cells):
            hdr = cells[1:] if not re.search(r"环比|同比", cells[0]) else cells
            mom_i = next(i for i, c in enumerate(hdr) if "环比" in c)
            yoy_i = next(i for i, c in enumerate(hdr) if "同比" in c and not re.search(r"\d月|累计", c))
            continue
        if yoy_i is None or len(cells) < 2:
            continue
        name = re.sub(r"^[一二三四五六七八九十]+、", "", re.sub(r"^其中[:：]", "", cells[0]))
        vals = [nbs.num(c) for c in cells[1:]]
        if not vals or all(v is None for v in vals):
            continue
        stem = BY_NAME.get(name)
        if stem is None:
            unmatched.append(name)
            continue
        if f"{stem}_YoY_pct" in row:
            continue
        if yoy_i < len(vals) and vals[yoy_i] is not None:
            row[f"{stem}_YoY_pct"] = vals[yoy_i]
        if mom_i < len(vals) and vals[mom_i] is not None:
            row[f"{stem}_MoM_pct"] = vals[mom_i]
    return row, unmatched


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()
    cols = columns()
    data = nbs.read_sheet(args.out, "Data") if nbs.has_sheet(args.out, "Series") else None
    data = pd.DataFrame(columns=cols, dtype=float) if data is None else data.reindex(columns=cols)
    held = {(d.year, d.month) for d in data.index[data.notna().any(axis=1)]} if len(data) else set()
    gaps = nbs.recent_gaps(data.index[data.notna().any(axis=1)] if len(data) else [], "MS")
    deep = (not held) or bool(gaps)

    def wanted(title):
        p, _ = nbs.month_period(title)
        return p is not None and p >= pd.Timestamp(HISTORY_START) and (p.year, p.month) not in held

    releases = nbs.crawl_index(TITLE_RE, wanted, deep, EXTRA_RELEASES)
    nbs.log(f"  held {len(held)} months; {len(releases)} release(s) to fetch")
    new, unmatched_all = {}, set()
    for title, url in releases:
        p, _ = nbs.month_period(title)
        try:
            html = nbs.fetch(url)
        except Exception as e:  # noqa: BLE001 - next run retries
            nbs.log(f"  [{p:%Y-%m}] fetch failed ({type(e).__name__}) - skipped")
            continue
        row, unmatched = parse_release(html or "")
        unmatched_all.update(unmatched)
        if not row:
            nbs.log(f"  [{p:%Y-%m}] no table parsed - skipped ({url})")
            continue
        new[p] = row
        nbs.log(f"  [{p:%Y-%m}] {len(row) // 2} series, PPI y/y {row.get('PPI_YoY_pct')}")
    if unmatched_all:
        nbs.log(f"  rows not mapped (left out): {sorted(unmatched_all)}")
    if new:
        add = pd.DataFrame.from_dict(new, orient="index")
        data = data.combine_first(add) if len(data) else add
    if data.dropna(how="all").empty:
        raise SystemExit("No data at all - nothing to save.")
    data = data.reindex(columns=cols).sort_index()
    data.index.name = "month"
    rows = []
    for _cn, stem, label, group in SERIES:
        rows.append((f"{stem}_YoY_pct", f"{label}, y/y", "% y/y", group or "", "line"))
        rows.append((f"{stem}_MoM_pct", f"{label}, m/m", "% m/m", "", "line"))
    series = nbs.series_sheet(rows)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": data, "Series": series}, NOTES_LINES, NOTES_SECTION_TITLES)
    print(f"Saved {len(data)} month(s) ({data.index.min():%Y-%m}..{data.index.max():%Y-%m}), {len(new)} new, "
          f"to {args.out}")


NOTES_LINES = [
    "UNITS",
    "*_YoY_pct: % change vs the same month a year earlier. *_MoM_pct: % change vs the previous month. NBS "
    "publishes no index levels for these series. PPI = ex-factory prices of industrial producers; PPIRM = "
    "purchaser prices of industrial producers (raw materials, fuel, power). Industry series are ex-factory "
    "prices by industry. Full names on the 'Series' sheet; only y/y series in the main groups are charted.",
    "",
    "SOURCE",
    "National Bureau of Statistics of China, monthly release 'YYYY年M月份工业生产者出厂价格...', Chinese release "
    "list https://www.stats.gov.cn/sj/zxfb/. The release also gives year-to-date y/y changes (not kept).",
    "",
    "UPDATES",
    "Incremental: months already held are not re-fetched.",
]
NOTES_SECTION_TITLES = {"UNITS", "SOURCE", "UPDATES"}


if __name__ == "__main__":
    main()
