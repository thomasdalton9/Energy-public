"""
Pull China NBS's monthly industrial-profit release ('2026年1—8月份全国规模以上工业企业利润增长X%'): year-to-date operating
revenue, operating cost and total profit (bn yuan) with their y/y growth for industrial enterprises above designated size,
by sector, ownership and industry (41 industries incl. coal mining, oil and gas extraction, petroleum/coal/fuel
processing, power and heat, gas supply, steel, non-ferrous, chemicals), the sector financial ratios (profit margin, cost and
expense per 100 yuan of revenue, asset turnover, revenue per head, debt ratio, inventory and receivable days) and the
single-month profit growth quoted in the text.

Every figure is YEAR TO DATE as NBS publishes it (the row of month N covers January to N; the January-February
release is the first). Nothing is differenced into monthly values. The y/y growth is NBS's like-for-like figure; where
the previous year was a loss NBS prints a note instead of a rate and the cell is blank.

    python3 asia/CHINA_NBS_PROFITS.py --out "output/Data and Chart Outputs/china_nbs_industrial_profits_monthly.xlsx"
"""

import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))                   # asia/
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root

import china_nbs_common as nbs  # noqa: E402
import china_nbs_archive_ids as arch  # noqa: E402
import china_nbs_tables as tbl  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "china_nbs_industrial_profits_monthly.xlsx")
TITLE_RE = r"\d{4}年(1[—\-－～~]\d{1,2}月份?|上半年|前三季度|一季度|全国|全年)(全国)?规模以上工业企业利润"
HISTORY_START = "2013-01-01"
EXTRA_RELEASES = arch.PROFITS   # releases no longer on the list, from the ID scan

P = "profit YTD y/y energy industries"
# (cleaned Chinese name(s), stem, English label, chart group of the profit y/y column)
GROUPS = [
    (("总计",), "All", "All industry", "profit YTD y/y by sector"),
    (("采矿业",), "Mining", "Mining", "profit YTD y/y by sector"),
    (("制造业",), "Manufacturing", "Manufacturing", "profit YTD y/y by sector"),
    (("电力、热力、燃气及水生产和供应业",), "Utilities", "Power, heat, gas and water supply", "profit YTD y/y by sector"),
    (("国有控股企业", "国有及国有控股企业"), "State_Controlled", "State-controlled enterprises", ""),
    (("集体企业",), "Collective", "Collective enterprises (to 2018)", ""),
    (("股份制企业",), "Joint_Stock", "Joint-stock enterprises", ""),
    (("外商及港澳台投资企业", "外商及港澳台商投资企业"), "Foreign_HMT", "Foreign, Hong Kong, Macao and Taiwan invested", ""),
    (("私营企业",), "Private", "Private enterprises", ""),
]
INDUSTRIES = [
    (("煤炭开采和洗选业",), "Coal_Mining", "Coal mining and washing", P),
    (("石油和天然气开采业",), "Oil_Gas_Extraction", "Oil and gas extraction", P),
    (("黑色金属矿采选业",), "Ferrous_Ore_Mining", "Ferrous metal ore mining", ""),
    (("有色金属矿采选业",), "Non_Ferrous_Ore_Mining", "Non-ferrous metal ore mining", ""),
    (("非金属矿采选业",), "Non_Metallic_Mining", "Non-metallic mineral mining", ""),
    (("开采专业及辅助性活动", "开采辅助活动"), "Mining_Support", "Mining support activities", ""),
    (("其他采矿业",), "Other_Mining", "Other mining", ""),
    (("农副食品加工业",), "Agri_Food_Processing", "Agricultural food processing", ""),
    (("食品制造业",), "Food_Manufacturing", "Food manufacturing", ""),
    (("酒、饮料和精制茶制造业",), "Beverages", "Liquor, beverages and tea", ""),
    (("烟草制品业",), "Tobacco", "Tobacco", ""),
    (("纺织业",), "Textiles", "Textiles", ""),
    (("纺织服装、服饰业",), "Apparel", "Apparel", ""),
    (("皮革、毛皮、羽毛及其制品和制鞋业",), "Leather_Footwear", "Leather, fur, feathers and footwear", ""),
    (("木材加工和木、竹、藤、棕、草制品业",), "Wood_Products", "Wood products", ""),
    (("家具制造业",), "Furniture", "Furniture", ""),
    (("造纸和纸制品业",), "Paper", "Paper", ""),
    (("印刷和记录媒介复制业",), "Printing", "Printing", ""),
    (("文教、工美、体育和娱乐用品制造业",), "Culture_Sports_Goods", "Culture, education, sports and recreation goods", ""),
    (("石油、煤炭及其他燃料加工业", "石油加工、炼焦和核燃料加工业"), "Petroleum_Coal_Processing",
     "Petroleum, coal and other fuel processing", P),
    (("化学原料和化学制品制造业",), "Chemicals", "Chemical raw materials and products", ""),
    (("医药制造业",), "Pharmaceuticals", "Pharmaceuticals", ""),
    (("化学纤维制造业",), "Chemical_Fibres", "Chemical fibres", ""),
    (("橡胶和塑料制品业",), "Rubber_Plastics", "Rubber and plastics", ""),
    (("非金属矿物制品业",), "Non_Metallic_Minerals", "Non-metallic mineral products (cement, glass)", ""),
    (("黑色金属冶炼和压延加工业",), "Ferrous_Smelting", "Ferrous metal smelting and rolling (steel)", ""),
    (("有色金属冶炼和压延加工业",), "Non_Ferrous_Smelting", "Non-ferrous metal smelting and rolling", ""),
    (("金属制品业",), "Metal_Products", "Metal products", ""),
    (("通用设备制造业",), "General_Equipment", "General-purpose equipment", ""),
    (("专用设备制造业",), "Special_Equipment", "Special-purpose equipment", ""),
    (("汽车制造业",), "Automobiles", "Automobiles", ""),
    (("铁路、船舶、航空航天和其他运输设备制造业",), "Other_Transport_Equipment", "Rail, ship, aerospace and other transport", ""),
    (("电气机械和器材制造业",), "Electrical_Machinery", "Electrical machinery and equipment", ""),
    (("计算机、通信和其他电子设备制造业",), "Electronics", "Computers, communication and electronics", ""),
    (("仪器仪表制造业",), "Instruments", "Instruments and meters", ""),
    (("其他制造业",), "Other_Manufacturing", "Other manufacturing", ""),
    (("废弃资源综合利用业",), "Waste_Recycling", "Waste resource recycling", ""),
    (("金属制品、机械和设备修理业",), "Repair", "Repair of metal products, machinery and equipment", ""),
    (("电力、热力生产和供应业",), "Power_Heat", "Power and heat production and supply", P),
    (("燃气生产和供应业",), "Gas_Supply", "Gas production and supply", P),
    (("水的生产和供应业",), "Water_Supply", "Water production and supply", ""),
]
GROUP_BY = {n: stem for names, stem, *_ in GROUPS for n in names}
IND_BY = {n: stem for names, stem, *_ in INDUSTRIES for n in names}
MEASURES = [("Revenue_bn_yuan", 0.1), ("Revenue_YoY_pct", None), ("Cost_bn_yuan", 0.1), ("Cost_YoY_pct", None),
            ("Profit_bn_yuan", 0.1), ("Profit_YoY_pct", None)]
RATIOS = ["Margin_pct", "Cost_per_100_yuan", "Expense_per_100_yuan", "Revenue_per_100_yuan_assets",
          "Revenue_per_head_10k_yuan", "Debt_ratio_pct", "Inventory_days", "Receivable_days"]
STEMS = [s for _n, s, *_ in GROUPS + INDUSTRIES]
COLUMNS = ([f"{s}_{m}" for s in STEMS for m, _f in MEASURES] + [f"{s}_{r}" for _n, s, *_ in GROUPS for r in RATIOS]
           + ["All_Month_Profit_YoY_pct"])


def parse_release(html, _is_jf=False):
    row, unmatched, mode = {}, [], None
    for cells in nbs.table_rows(html):
        head = "".join(cells)
        if cells[0] in ("分组", "行业", "分 组", "行 业") or cells[0].replace(" ", "") in ("分组", "行业"):
            mode = "ratios" if "营业收入利润率" in head else ("industry" if cells[0].replace(" ", "") == "行业" else "group")
            continue
        if mode is None or len(cells) < 4:
            continue
        name = tbl.clean_name(cells[0])
        vals = [nbs.num(c) for c in cells[1:]]
        if mode == "ratios":
            stem = GROUP_BY.get(name)
            if stem and len(vals) >= 8 and f"{stem}_Margin_pct" not in row:
                for r, v in zip(RATIOS, vals):
                    if v is not None:
                        row[f"{stem}_{r}"] = v
            continue
        stem = (GROUP_BY if mode == "group" else IND_BY).get(name)
        if stem is None:
            if re.search(r"[一-鿿]", name) and not name.startswith("注"):
                unmatched.append(name)
            continue
        if len(vals) < 6 or f"{stem}_Revenue_bn_yuan" in row:
            continue
        for (m, f), v in zip(MEASURES, vals):
            if v is not None:
                row[f"{stem}_{m}"] = round(v * f, 4) if f else v
    text = re.sub(r"<[^>]+>", "", html)
    m = re.search(r"\d{1,2}月份，规模以上工业企业利润同比(增长|下降)([\d.]+)%", text)
    if m:
        row["All_Month_Profit_YoY_pct"] = float(m.group(2)) * (1 if m.group(1) == "增长" else -1)
    if unmatched:
        nbs.log(f"    rows not mapped: {sorted(set(unmatched))[:10]}")
    return row


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()
    rows = []
    labels = {s: (lab, grp) for _n, s, lab, grp in GROUPS + INDUSTRIES}
    for s in STEMS:
        lab, grp = labels[s]
        rows.append((f"{s}_Revenue_bn_yuan", f"{lab}: operating revenue, year to date", "bn yuan (year to date)", "", "line"))
        rows.append((f"{s}_Revenue_YoY_pct", f"{lab}: operating revenue, y/y", "% y/y (year to date)", "", "line"))
        rows.append((f"{s}_Cost_bn_yuan", f"{lab}: operating cost, year to date", "bn yuan (year to date)", "", "line"))
        rows.append((f"{s}_Cost_YoY_pct", f"{lab}: operating cost, y/y", "% y/y (year to date)", "", "line"))
        rows.append((f"{s}_Profit_bn_yuan", f"{lab}: total profit, year to date", "bn yuan (year to date)", "", "line"))
        rows.append((f"{s}_Profit_YoY_pct", f"{lab}: total profit, y/y", "% y/y (year to date)", grp, "line"))
    for _n, s, lab, _g in GROUPS:
        for r, unit in zip(RATIOS, ["%", "yuan", "yuan", "yuan", "10,000 yuan per person", "%", "days", "days"]):
            rows.append((f"{s}_{r}", f"{lab}: {r.replace('_', ' ').lower()}", unit, "", "line"))
    rows.append(("All_Month_Profit_YoY_pct", "All industry: profit growth in the single month (text of the release)",
                 "% y/y (month)", "", "line"))
    tbl.pull(args.out, title_re=TITLE_RE, near_re=r"规模以上工业企业利润", period_fn=tbl.ytd_period, extra=EXTRA_RELEASES,
             parse_fn=parse_release, columns=COLUMNS, series_rows=rows, notes_lines=NOTES_LINES,
             notes_titles=NOTES_SECTION_TITLES, history_start=HISTORY_START, probe_col="All_Profit_bn_yuan",
             skip_gap_months=(1,), index_name="period_end_month")


NOTES_LINES = [
    "UNITS",
    "Operating revenue, operating cost and total profit in billion yuan (NBS prints 100 million yuan; x 0.1), YEAR TO "
    "DATE: the row dated month N covers January to N (the row dated February is January-February, the first release of "
    "the year; there is no January row; the December row is the full year). *_YoY_pct: NBS's y/y growth of that year-to-date "
    "total on a like-for-like basis (blank where the previous year was a loss). Ratios (sector rows only): profit margin "
    "% of revenue, cost and expense per 100 yuan of revenue, revenue per 100 yuan of assets, revenue per head (10,000 yuan), "
    "debt ratio %, finished-goods inventory days and receivable days (the last four are end-of-month stocks). "
    "All_Month_Profit_YoY_pct is the single-month profit growth quoted in the text.",
    "",
    "SCOPE",
    "Industrial enterprises above designated size (annual main business revenue of RMB 20 million or more). Ownership groups "
    "overlap, so they do not add up to the total. Nothing is differenced into monthly values.",
    "",
    "SOURCE",
    "National Bureau of Statistics of China, monthly release '...规模以上工业企业利润...' tables 1-3, Chinese release list "
    "https://www.stats.gov.cn/sj/zxfb/.",
    "",
    "UPDATES",
    "Incremental: periods already held are not re-fetched.",
]
NOTES_SECTION_TITLES = {"UNITS", "SCOPE", "SOURCE", "UPDATES"}


if __name__ == "__main__":
    main()
