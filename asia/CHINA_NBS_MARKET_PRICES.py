"""
Pull China NBS's 10-day market prices of important means of production
(流通领域重要生产资料市场价格变动情况): ~50 producer goods monitored
nationwide every 旬 (1st-10th, 11th-20th, 21st-end of month), incl.
coal grades (anthracite, 4500/5000/5500/5800 kcal steam coal, coking
coal), metallurgical coke, LNG, LPG, gasoline 92#/95#, diesel 0#,
steel products, copper/aluminium/lead/zinc, basic chemicals and
polymers, cement, float glass, fertilisers, farm and forest products.

Each release is an HTML table [product (spec) | unit | price (yuan) |
change vs previous period (yuan) | change %]; only the price is kept.
Chinese release list www.stats.gov.cn/sj/zxfb/ (back to Oct 2021, see
china_nbs_common.py). An English version exists only from Apr 2024.

Incremental: periods already in the workbook are not re-fetched.

    python3 asia/CHINA_NBS_MARKET_PRICES.py --out "output/Data and Chart Outputs/china_nbs_market_prices_10day.xlsx"
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
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "china_nbs_market_prices_10day.xlsx")
TITLE_RE = r"\d{4}年\d{1,2}月[上中下]旬流通领域重要生产资料市场价格"
HISTORY_START = "2021-01-01"
# Releases that still exist but are no longer on the (~1000-item) release list, found by
# CHINA_NBS_DISCOVERY5.py's ID scan of the Feb-2023 site migration.
_OLD = "https://www.stats.gov.cn/sj/zxfb/202302/t20230203_{}.html"
EXTRA_RELEASES = [(t + "流通领域重要生产资料市场价格变动情况", _OLD.format(i)) for t, i in [
    ("2021年10月上旬", 1901236),
    ("2021年9月下旬", 1901235),
    ("2021年9月中旬", 1901229),
    ("2021年9月上旬", 1901214),
    ("2021年8月下旬", 1901210),
    ("2021年8月中旬", 1901201),
    ("2021年8月上旬", 1901187),
    ("2021年7月下旬", 1901183),
    ("2021年7月中旬", 1901176),
    ("2021年7月上旬", 1901151),
    ("2021年6月下旬", 1901144),
    ("2021年6月中旬", 1901139),
    ("2021年6月上旬", 1901124),
    ("2021年5月下旬", 1901119),
    ("2021年5月中旬", 1901112),
    ("2021年5月上旬", 1901095),
    ("2021年4月下旬", 1901076),
    ("2021年4月中旬", 1901066),
    ("2021年4月上旬", 1901043),
    ("2021年3月下旬", 1901039),
    ("2021年3月中旬", 1901034),
    ("2021年3月上旬", 1901021),
    ("2021年2月下旬", 1901008),
    ("2021年2月上旬", 1900996),
    ("2021年1月下旬", 1900991),
    ("2021年1月中旬", 1900984),
    ("2021年1月上旬", 1900961),
]]

# (regex on "name(spec)" as printed, column, English label, chart group). First match wins.
PRODUCTS = [
    (r"^无烟煤", "Anthracite", "Anthracite (washed medium lump)", "coal and coke prices"),
    (r"^普通混煤", "Steam_Coal_4500", "Steam coal, ordinary blend (4500 kcal)", "coal and coke prices"),
    (r"^山西大混", "Steam_Coal_Shanxi_5000", "Steam coal, Shanxi blend (5000 kcal)", "coal and coke prices"),
    (r"^山西优混", "Steam_Coal_Shanxi_5500", "Steam coal, Shanxi premium blend (5500 kcal)", "coal and coke prices"),
    (r"^大同混煤", "Steam_Coal_Datong_5800", "Steam coal, Datong blend (5800 kcal)", "coal and coke prices"),
    (r"^焦煤", "Coking_Coal", "Coking coal (prime)", "coal and coke prices"),
    (r"^焦炭", "Coke", "Metallurgical coke (quasi grade 1)", "coal and coke prices"),
    (r"液化天然气|LNG", "LNG", "LNG", "oil and gas product prices"),
    (r"液化石油气|LPG", "LPG", "LPG", "oil and gas product prices"),
    (r"^汽油.*95", "Gasoline_95", "Gasoline 95#", "oil and gas product prices"),
    (r"^汽油.*92", "Gasoline_92", "Gasoline 92#", "oil and gas product prices"),
    (r"^柴油", "Diesel_0", "Diesel 0#", "oil and gas product prices"),
    (r"^石蜡", "Paraffin_Wax", "Paraffin wax", "oil and gas product prices"),
    (r"^螺纹钢", "Rebar", "Rebar (HRB400E, 16-25mm)", "steel prices"),
    (r"^线材", "Wire_Rod", "Wire rod (HPB300, 6.5mm)", "steel prices"),
    (r"^普通中板", "Medium_Plate", "Medium plate (Q235, 20mm)", "steel prices"),
    (r"^热轧普通板卷", "Hot_Rolled_Coil", "Hot-rolled coil (Q235)", "steel prices"),
    (r"^无缝钢管", "Seamless_Pipe", "Seamless steel pipe", "steel prices"),
    (r"^角钢", "Angle_Steel", "Angle steel", "steel prices"),
    (r"^电解铜", "Copper", "Electrolytic copper", "non-ferrous metal prices"),
    (r"^铝锭", "Aluminium", "Aluminium ingot (A00)", "non-ferrous metal prices"),
    (r"^铅锭", "Lead", "Lead ingot", "non-ferrous metal prices"),
    (r"^锌锭", "Zinc", "Zinc ingot", "non-ferrous metal prices"),
    (r"^硫酸", "Sulfuric_Acid", "Sulfuric acid (98%)", "basic chemical prices"),
    (r"^烧碱", "Caustic_Soda", "Caustic soda (liquid, 32%)", "basic chemical prices"),
    (r"^甲醇", "Methanol", "Methanol", "basic chemical prices"),
    (r"^纯苯", "Benzene", "Benzene", "basic chemical prices"),
    (r"^苯乙烯", "Styrene", "Styrene", "basic chemical prices"),
    (r"^聚乙烯", "Polyethylene", "Polyethylene (LLDPE film)", "polymer and fibre prices"),
    (r"^聚丙烯", "Polypropylene", "Polypropylene (drawing grade)", "polymer and fibre prices"),
    (r"^聚氯乙烯", "PVC", "PVC (SG5)", "polymer and fibre prices"),
    (r"^顺丁胶", "Butadiene_Rubber", "Butadiene rubber (BR9000)", "polymer and fibre prices"),
    (r"^涤纶长丝", "Polyester_Filament", "Polyester filament (POY)", "polymer and fibre prices"),
    (r"^普通硅酸盐水泥.*袋装", "Cement_Bagged", "Cement P.O 42.5, bagged", "building material prices"),
    (r"^普通硅酸盐水泥.*散装", "Cement_Bulk", "Cement P.O 42.5, bulk", "building material prices"),
    (r"^浮法平板玻璃", "Float_Glass", "Float glass (4.8/5mm)", "building material prices"),
    (r"^尿素", "Urea", "Urea", "fertiliser and agrochemical prices"),
    (r"^复合肥", "Compound_Fertiliser", "Compound fertiliser (potassium sulphate, 45%)", "fertiliser and agrochemical prices"),
    (r"^农药|草甘膦", "Glyphosate", "Pesticide (glyphosate, 95%)", "fertiliser and agrochemical prices"),
    (r"^稻米", "Rice", "Rice (japonica)", "farm product prices"),
    (r"^小麦", "Wheat", "Wheat", "farm product prices"),
    (r"^玉米", "Corn", "Corn", "farm product prices"),
    (r"^棉花", "Cotton", "Cotton (lint)", "farm product prices"),
    (r"^大豆", "Soybeans", "Soybeans", "farm product prices"),
    (r"^豆粕", "Soybean_Meal", "Soybean meal", "farm product prices"),
    (r"^花生", "Peanuts", "Peanuts", "farm product prices"),
    (r"^生猪", "Live_Hogs", "Live hogs", "live hog prices"),
    (r"^天然橡胶", "Natural_Rubber", "Natural rubber", "forest product prices"),
    (r"^纸浆", "Pulp", "Pulp (imported softwood)", "forest product prices"),
    (r"^瓦楞纸", "Corrugated_Paper", "Corrugated paper", "forest product prices"),
]
UNITS = {"吨": "yuan/tonne", "千克": "yuan/kg", "公斤": "yuan/kg"}
XUN_DAY = {"上": 1, "中": 11, "下": 21}


def period_of(title):
    m = re.search(r"(\d{4})年(\d{1,2})月([上中下])旬", title)
    return pd.Timestamp(int(m.group(1)), int(m.group(2)), XUN_DAY[m.group(3)]) if m else None


def parse_release(html):
    """({column: price}, {column: unit}, [unmatched product names])."""
    prices, units, unmatched = {}, {}, []
    for cells in nbs.table_rows(html):
        if len(cells) < 3 or cells[0] in ("产品名称",) or cells[0].isdigit():
            continue
        unit, price = cells[1], nbs.num(cells[2])
        if unit not in UNITS or price is None:
            continue
        name = cells[0].replace("（", "(").replace("）", ")")
        for rx, col, *_ in PRODUCTS:
            if re.search(rx, name):
                if col not in prices:
                    prices[col], units[col] = price, UNITS[unit]
                break
        else:
            unmatched.append(name)
    return prices, units, unmatched


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()
    cols = [p[1] for p in PRODUCTS]

    data = nbs.read_sheet(args.out, "Data") if nbs.has_sheet(args.out, "Series") else None
    data = pd.DataFrame(columns=cols, dtype=float) if data is None else data.reindex(columns=cols)
    held = set(data.index[data.notna().any(axis=1)]) if len(data) else set()
    gaps = nbs.recent_gaps(held, "10D")
    deep = (not held) or bool(gaps)
    nbs.log(f"  held {len(held)} periods, {len(gaps)} gaps -> {'full' if deep else 'shallow'} crawl")

    def wanted(title):
        p = period_of(title)
        return p is not None and p >= pd.Timestamp(HISTORY_START) and p not in held

    releases = nbs.crawl_index(TITLE_RE, wanted, deep, EXTRA_RELEASES, stop_after_known=6)
    nbs.log(f"  {len(releases)} release(s) to fetch")
    new, unit_of, unmatched_all = {}, {}, set()
    for title, url in releases:
        p = period_of(title)
        try:
            html = nbs.fetch(url)
        except Exception as e:  # noqa: BLE001 - next run retries
            nbs.log(f"  [{p:%Y-%m-%d}] fetch failed ({type(e).__name__}) - skipped")
            continue
        prices, units, unmatched = parse_release(html or "")
        if not prices:
            nbs.log(f"  [{p:%Y-%m-%d}] no prices parsed - skipped ({url})")
            continue
        new[p] = prices
        unit_of.update(units)
        unmatched_all.update(unmatched)
        nbs.log(f"  [{p:%Y-%m-%d}] {len(prices)} prices")
    if unmatched_all:
        nbs.log(f"  products not mapped (left out): {sorted(unmatched_all)}")

    if new:
        add = pd.DataFrame.from_dict(new, orient="index")
        data = data.combine_first(add) if len(data) else add
    if data.dropna(how="all").empty:
        raise SystemExit("No prices at all - nothing to save.")
    data = data.reindex(columns=cols).sort_index()
    data.index.name = "period_start"
    old_series = pd.read_excel(args.out, sheet_name="Series", index_col=0) if nbs.has_sheet(args.out, "Series") else None
    unit_rows = []
    for _rx, col, label, group in PRODUCTS:
        u = unit_of.get(col) or (old_series.loc[col, "unit"] if old_series is not None and col in old_series.index
                                 else ("yuan/kg" if col == "Live_Hogs" else "yuan/tonne"))
        unit_rows.append((col, label, u, group, "line"))
    series = nbs.series_sheet(unit_rows)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": data, "Series": series}, NOTES_LINES, NOTES_SECTION_TITLES)
    print(f"Saved {len(data)} 10-day period(s) ({data.index.min():%Y-%m-%d}..{data.index.max():%Y-%m-%d}), "
          f"{len(new)} new, to {args.out}")


NOTES_LINES = [
    "UNITS",
    "Prices in yuan per tonne (live hogs: yuan per kg), average market transaction price in circulation "
    "for that 10-day period, nationwide (NBS's monitoring survey of the circulation sector). "
    "Column names are on the 'Series' sheet with the full product description and unit.",
    "",
    "PERIODS",
    "Each row is one 旬 (10-day period): dated the 1st (1st-10th), 11th (11th-20th) or 21st (21st-end of month).",
    "",
    "SOURCE",
    "National Bureau of Statistics of China, 'YYYY年M月X旬流通领域重要生产资料市场价格变动情况' (Market Prices of "
    "Important Means of Production in Circulation), Chinese release list https://www.stats.gov.cn/sj/zxfb/ "
    "(reaches back to Oct 2021; English version only from Apr 2024). Only the price column is kept (the "
    "release also gives the change vs the previous period).",
    "",
    "UPDATES",
    "Incremental: periods already held are not re-fetched; each run adds new periods and fills gaps still "
    "on the release list.",
]
NOTES_SECTION_TITLES = {"UNITS", "PERIODS", "SOURCE", "UPDATES"}


if __name__ == "__main__":
    main()
