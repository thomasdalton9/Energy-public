"""
Pull China NBS's monthly consumer price index (CPI) from the Chinese release '2026年8月份居民消费价格...': month-on-month,
year-on-year and year-to-date y/y % for the headline, urban/rural, food/non-food, goods/services, CPI excluding food
and energy, the eight categories and their sub-items - incl. the energy-related ones: housing utilities
(水电燃料: water, electricity and fuel) and fuel for vehicles (交通工具用燃料, renamed 交通工具用能源 from the
January 2026 release; one column). NBS publishes only % changes, no index levels.

January and February are separate CPI releases (no combined release), so every month is present from Oct 2021 on the
release list, Jan-Sep 2021 from the migrated archive (EXTRA_RELEASES).

    python3 asia/CHINA_NBS_CPI.py --out "output/Data and Chart Outputs/china_nbs_cpi_monthly.xlsx"
"""

import argparse
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))                   # asia/
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root

import china_nbs_common as nbs  # noqa: E402
import china_nbs_archive_ids as arch  # noqa: E402
import china_nbs_tables as tbl  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "china_nbs_cpi_monthly.xlsx")
TITLE_RE = r"\d{4}年\d{1,2}月份?居民消费价格"
HISTORY_START = "2013-01-01"
EXTRA_RELEASES = arch.CPI   # releases no longer on the list, from the ID scan

# (cleaned Chinese row name(s), column stem, English label, chart group for the y/y column)
H = "CPI headline y/y"
E = "CPI energy-related items y/y"
SERIES = [
    (("居民消费价格",), "CPI", "CPI (all items)", H),
    (("城市",), "CPI_Urban", "Urban", ""),
    (("农村",), "CPI_Rural", "Rural", ""),
    (("食品",), "CPI_Food", "Food", H),
    (("非食品",), "CPI_Non_Food", "Non-food", H),
    (("消费品",), "CPI_Goods", "Consumer goods", H),
    (("服务",), "CPI_Services", "Services", H),
    (("不包括食品和能源",), "CPI_Core_Ex_Food_Energy", "Excluding food and energy", H),
    (("食品烟酒及在外餐饮", "食品烟酒"), "CPI_Food_Tobacco_Liquor", "Food, tobacco and liquor (incl. dining out)", ""),
    (("粮食",), "CPI_Grain", "Grain", ""),
    (("食用油",), "CPI_Edible_Oil", "Edible oil", ""),
    (("鲜菜",), "CPI_Fresh_Vegetables", "Fresh vegetables", ""),
    (("畜肉类",), "CPI_Meat", "Meat", ""),
    (("猪肉",), "CPI_Pork", "Pork", ""),
    (("牛肉",), "CPI_Beef", "Beef", ""),
    (("羊肉",), "CPI_Mutton", "Mutton", ""),
    (("水产品",), "CPI_Aquatic_Products", "Aquatic products", ""),
    (("蛋类",), "CPI_Eggs", "Eggs", ""),
    (("奶类",), "CPI_Dairy", "Dairy", ""),
    (("鲜果",), "CPI_Fresh_Fruit", "Fresh fruit", ""),
    (("卷烟",), "CPI_Cigarettes", "Cigarettes", ""),
    (("酒类",), "CPI_Liquor", "Liquor", ""),
    (("衣着",), "CPI_Clothing", "Clothing", ""),
    (("服装",), "CPI_Garments", "Garments", ""),
    (("鞋类",), "CPI_Footwear", "Footwear", ""),
    (("居住",), "CPI_Housing", "Housing", E),
    (("租赁房房租",), "CPI_Rent", "Rent", ""),
    (("水电燃料",), "CPI_Utilities_Fuel", "Housing: water, electricity and fuel", E),
    (("生活用品及服务",), "CPI_Household_Goods_Services", "Household goods and services", ""),
    (("家用器具",), "CPI_Appliances", "Household appliances", ""),
    (("家庭服务",), "CPI_Household_Services", "Household services", ""),
    (("交通通信",), "CPI_Transport_Comms", "Transport and communication", E),
    (("交通工具", "小汽车"), "CPI_Vehicles", "Vehicles (cars)", ""),
    (("交通工具用燃料", "交通工具用能源"), "CPI_Vehicle_Fuel", "Fuel for vehicles (renamed 'energy for vehicles' in 2026)", E),
    (("交通工具使用和维修",), "CPI_Vehicle_Use_Repair", "Vehicle use and repair", ""),
    (("通信工具",), "CPI_Communication_Devices", "Communication devices", ""),
    (("通信服务",), "CPI_Communication_Services", "Communication services", ""),
    (("邮递服务",), "CPI_Postal_Services", "Postal services", ""),
    (("教育文化娱乐",), "CPI_Education_Culture", "Education, culture and recreation", ""),
    (("教育服务",), "CPI_Education_Services", "Education services", ""),
    (("旅游", "旅行社及其他旅游服务"), "CPI_Tourism", "Tourism", ""),
    (("医疗保健",), "CPI_Health", "Health care", ""),
    (("中药",), "CPI_Chinese_Medicine", "Chinese medicine", ""),
    (("西药",), "CPI_Western_Medicine", "Western medicine", ""),
    (("医疗服务",), "CPI_Medical_Services", "Medical services", ""),
    (("其他用品及服务",), "CPI_Other", "Other goods and services", ""),
]
BY_NAME = {n: stem for names, stem, *_ in SERIES for n in names}
COLUMNS = [f"{stem}_{k}" for _n, stem, *_ in SERIES for k in ("YoY_pct", "MoM_pct", "YTD_YoY_pct")]


def parse_release(html, _is_jf=False):
    """{column: value} from the 'main data' table: [name, m/m, y/y, YTD y/y] (no YTD column in January's release)."""
    row, unmatched = {}, []
    for cells in nbs.table_rows(html):
        if len(cells) < 3:
            continue
        name = tbl.clean_name(cells[0])
        vals = [nbs.num(c) for c in cells[1:]]
        if all(v is None for v in vals):
            continue
        stem = BY_NAME.get(name)
        if stem is None:
            unmatched.append(name)
            continue
        if f"{stem}_YoY_pct" in row:
            continue
        for k, i in (("MoM_pct", 0), ("YoY_pct", 1), ("YTD_YoY_pct", 2)):
            if i < len(vals) and vals[i] is not None:
                row[f"{stem}_{k}"] = vals[i]
    if unmatched:
        nbs.log(f"    rows not mapped: {sorted(set(unmatched))[:12]}")
    return row


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()
    rows = []
    for _n, stem, label, group in SERIES:
        rows.append((f"{stem}_YoY_pct", f"{label}, y/y", "% y/y", group, "line"))
        rows.append((f"{stem}_MoM_pct", f"{label}, m/m", "% m/m", "", "line"))
        rows.append((f"{stem}_YTD_YoY_pct", f"{label}, year-to-date y/y", "% y/y (year to date)", "", "line"))
    tbl.pull(args.out, title_re=TITLE_RE, near_re=r"居民消费价格", period_fn=nbs.month_period, extra=EXTRA_RELEASES,
             parse_fn=parse_release, columns=COLUMNS, series_rows=rows, notes_lines=NOTES_LINES,
             notes_titles=NOTES_SECTION_TITLES, history_start=HISTORY_START, probe_col="CPI_YoY_pct")


NOTES_LINES = [
    "UNITS",
    "*_YoY_pct: % change vs the same month a year earlier. *_MoM_pct: % change vs the previous month. *_YTD_YoY_pct: "
    "average change of the year to date vs the same period of the previous year (none in January's release). NBS publishes "
    "no index levels. Full names on the 'Series' sheet. CPI_Vehicle_Fuel is 交通工具用燃料 until Dec 2025 and 交通工具"
    "用能源 from the January 2026 release (the same row, renamed); CPI_Vehicles is 交通工具 until Dec 2025 and 小汽车 "
    "from 2026; CPI_Tourism is 旅游 until Dec 2025 and 旅行社及其他旅游服务 from 2026.",
    "",
    "SOURCE",
    "National Bureau of Statistics of China, monthly release 'YYYY年M月份居民消费价格...', table 'main data', Chinese "
    "release list https://www.stats.gov.cn/sj/zxfb/.",
    "",
    "UPDATES",
    "Incremental: months already held are not re-fetched.",
]
NOTES_SECTION_TITLES = {"UNITS", "SOURCE", "UPDATES"}


if __name__ == "__main__":
    main()
