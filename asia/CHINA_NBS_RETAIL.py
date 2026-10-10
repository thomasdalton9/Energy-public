"""
Pull China NBS's monthly retail-sales release ('2026年1—8月份社会消费品零售总额增长X%'): total retail sales of consumer
goods and, among the above-quota goods categories, petroleum and products (石油及制品类 - fuel retail), automobiles,
household appliances, building and decoration materials ..., as the month value (bn yuan), month y/y, year-to-date value
and year-to-date y/y (nominal). January and February come only combined (Jan-Feb sheet; nothing is split).

    python3 asia/CHINA_NBS_RETAIL.py --out "output/Data and Chart Outputs/china_nbs_retail_sales_monthly.xlsx"
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))                   # asia/
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root

import china_nbs_common as nbs  # noqa: E402
import china_nbs_archive_ids as arch  # noqa: E402
import china_nbs_tables as tbl  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "china_nbs_retail_sales_monthly.xlsx")
TITLE_RE = r"\d{4}年(\d{1,2}月份?|1[—\-－～~]\d{1,2}月份?|上半年|前三季度|一季度)?社会消费品零售总额"
HISTORY_START = "2013-01-01"
EXTRA_RELEASES = arch.RETAIL   # releases no longer on the list, from the ID scan

F = "retail sales y/y selected goods"
T = "retail sales y/y total"
SERIES = [
    (("社会消费品零售总额",), "Retail_Total", "Total retail sales of consumer goods", T),
    (("除汽车以外的消费品零售额",), "Retail_Ex_Autos", "Retail sales excluding automobiles", T),
    (("限额以上单位消费品零售额",), "Retail_Above_Quota", "Above-quota units' retail sales", ""),
    (("网上商品零售额", "实物商品网上零售额"), "Retail_Online_Goods", "Online retail sales of goods", ""),
    (("城镇",), "Retail_Urban", "Urban", ""),
    (("乡村",), "Retail_Rural", "Rural", ""),
    (("餐饮收入",), "Retail_Catering", "Catering revenue", T),
    (("限额以上单位餐饮收入",), "Retail_Catering_Above_Quota", "Catering revenue, above-quota units", ""),
    (("商品零售额", "商品零售"), "Retail_Goods", "Retail sales of goods", T),
    (("限额以上单位商品零售额", "限额以上单位商品零售"), "Retail_Goods_Above_Quota", "Goods retail sales, above-quota units", ""),
    (("粮油、食品类",), "Retail_Grain_Food", "Grain, oil and food", ""),
    (("饮料类",), "Retail_Beverages", "Beverages", ""),
    (("烟酒类",), "Retail_Tobacco_Liquor", "Tobacco and liquor", ""),
    (("服装鞋帽、针纺织品类", "服装、鞋帽、针纺织品类"), "Retail_Clothing", "Clothing, footwear and textiles", ""),
    (("化妆品类",), "Retail_Cosmetics", "Cosmetics", ""),
    (("金银珠宝类",), "Retail_Jewellery", "Gold, silver and jewellery", ""),
    (("日用品类",), "Retail_Daily_Goods", "Daily necessities", ""),
    (("体育、娱乐用品类",), "Retail_Sports_Goods", "Sports and recreation goods", ""),
    (("家用电器和音像器材类",), "Retail_Appliances", "Household appliances and AV equipment", F),
    (("中西药品类",), "Retail_Medicines", "Medicines", ""),
    (("文化办公用品类",), "Retail_Office_Goods", "Culture and office supplies", ""),
    (("家具类",), "Retail_Furniture", "Furniture", ""),
    (("通讯器材类",), "Retail_Telecom_Equipment", "Communication equipment", ""),
    (("石油及制品类",), "Retail_Petroleum_Products", "Petroleum and products (fuel retail)", F),
    (("汽车类",), "Retail_Autos", "Automobiles", F),
    (("建筑及装潢材料类",), "Retail_Building_Materials", "Building and decoration materials", F),
]
BY_NAME = {n: stem for names, stem, *_ in SERIES for n in names}
COLUMNS = [f"{s}_{k}" for _n, s, *_ in SERIES for k in ("bn_yuan", "YoY_pct", "YTD_bn_yuan", "YTD_YoY_pct")]


def parse_release(html, is_jf=False):
    """Rows [name, month value, month y/y, YTD value, YTD y/y]; the Jan-Feb release has [name, value, y/y] (stored as the
    month columns of the Jan-Feb sheet). Values in 100 million yuan -> bn yuan (x 0.1)."""
    row, unmatched = {}, []
    for cells in nbs.table_rows(html):
        if len(cells) < 3:
            continue
        name = tbl.clean_name(cells[0])
        stem = BY_NAME.get(name)
        vals = [nbs.num(c) for c in cells[1:]]
        if stem is None:
            if any(v is not None for v in vals) and not name.startswith(("环比", "注")) and not name[:1].isdigit() and \
                    "年" not in name and "月" not in name:
                unmatched.append(name)
            continue
        if f"{stem}_YoY_pct" in row or f"{stem}_YTD_YoY_pct" in row:
            continue
        pairs = [("bn_yuan", 0.1, 0), ("YoY_pct", None, 1), ("YTD_bn_yuan", 0.1, 2), ("YTD_YoY_pct", None, 3)]
        for k, f, i in pairs:
            if i < len(vals) and vals[i] is not None:
                row[f"{stem}_{k}"] = round(vals[i] * f, 3) if f else vals[i]
    if unmatched:
        nbs.log(f"    rows not mapped: {sorted(set(unmatched))[:10]}")
    return row


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()
    rows = []
    for _n, s, lab, grp in SERIES:
        rows.append((f"{s}_bn_yuan", f"{lab}, month", "bn yuan per month (nominal)", "", "line"))
        rows.append((f"{s}_YoY_pct", f"{lab}, y/y", "% y/y (nominal)", grp, "line"))
        rows.append((f"{s}_YTD_bn_yuan", f"{lab}, year to date", "bn yuan (year to date)", "", "line"))
        rows.append((f"{s}_YTD_YoY_pct", f"{lab}, year-to-date y/y", "% y/y (year to date, nominal)", "", "line"))
    tbl.pull(args.out, title_re=TITLE_RE, near_re=r"社会消费品零售总额", period_fn=tbl.ytd_period, extra=EXTRA_RELEASES,
             parse_fn=parse_release, columns=COLUMNS, series_rows=rows, notes_lines=NOTES_LINES,
             notes_titles=NOTES_SECTION_TITLES, history_start=HISTORY_START, probe_col="Retail_Total_YoY_pct",
             jan_feb_sheet=True, skip_gap_months=(1, 2))


NOTES_LINES = [
    "UNITS",
    "Retail sales in billion yuan (NBS prints 100 million yuan; x 0.1), nominal (price effects not removed). *_bn_yuan and "
    "*_YoY_pct are the month's value and its y/y; *_YTD_* the year to date. Category rows (food ... building materials, "
    "petroleum and products, automobiles) are above-quota units only. January and February are only published combined: "
    "the Jan-Feb sheet holds the two-month value and y/y in the month columns (dated 1 Feb), and the monthly Data sheet is "
    "blank for January and February. Releases for some months are titled by quarter or half-year (March, June, September) "
    "or by the year (December) but carry that month's table.",
    "",
    "SOURCE",
    "National Bureau of Statistics of China, monthly release '...社会消费品零售总额...', Chinese release list "
    "https://www.stats.gov.cn/sj/zxfb/.",
    "",
    "UPDATES",
    "Incremental: months already held are not re-fetched.",
]
NOTES_SECTION_TITLES = {"UNITS", "SOURCE", "UPDATES"}


if __name__ == "__main__":
    main()
