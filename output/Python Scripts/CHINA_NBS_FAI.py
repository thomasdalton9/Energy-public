"""
Pull China NBS's monthly fixed-asset-investment release ('2026年1—8月份全国固定资产投资...'): year-to-date y/y growth
(nominal, like-for-like) of investment (excl. rural households) in total, by ownership, by component (construction,
equipment), by sector and by industry - incl. mining, manufacturing and its industries (chemicals, non-ferrous,
automobiles ...) and power, heat, gas and water supply - plus, from the text, the year-to-date totals (bn yuan) for the
three sectors, industrial and infrastructure investment growth and growth by region. All year to date as published (the
row dated month N covers January to N; the January-February release is the first).

    python3 asia/CHINA_NBS_FAI.py --out "output/Data and Chart Outputs/china_nbs_fixed_asset_investment_monthly.xlsx"
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
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "china_nbs_fixed_asset_investment_monthly.xlsx")
TITLE_RE = r"\d{4}年(1[—\-－～~]\d{1,2}月份?|上半年|前三季度|一季度|全国)(全国)?固定资产投资"
HISTORY_START = "2013-01-01"
EXTRA_RELEASES = arch.FAI   # releases no longer on the list, from the ID scan

G = "fixed investment by sector y/y"
I = "fixed investment by industry y/y"
SERIES = [
    (("固定资产投资(不含农户)", "固定资产投资"), "FAI_Total", "Fixed-asset investment (excl. rural households)", G),
    (("国有控股",), "FAI_State_Controlled", "State-controlled", ""),
    (("民间投资",), "FAI_Private", "Private investment", ""),
    (("建筑安装工程",), "FAI_Construction", "Construction and installation", ""),
    (("设备工器具购置",), "FAI_Equipment", "Equipment and tools purchased", ""),
    (("其他费用",), "FAI_Other_Costs", "Other expenses", ""),
    (("第一产业",), "FAI_Primary", "Primary sector", G),
    (("第二产业",), "FAI_Secondary", "Secondary sector (industry and construction)", G),
    (("第三产业",), "FAI_Tertiary", "Tertiary sector (services)", G),
    (("农林牧渔业",), "FAI_Agriculture", "Agriculture, forestry, animal husbandry, fishery", ""),
    (("采矿业",), "FAI_Mining", "Mining", I),
    (("制造业",), "FAI_Manufacturing", "Manufacturing", I),
    (("农副食品加工业",), "FAI_Agri_Food_Processing", "Agricultural food processing", ""),
    (("食品制造业",), "FAI_Food_Manufacturing", "Food manufacturing", ""),
    (("纺织业",), "FAI_Textiles", "Textiles", ""),
    (("化学原料和化学制品制造业",), "FAI_Chemicals", "Chemical raw materials and products", ""),
    (("医药制造业",), "FAI_Pharmaceuticals", "Pharmaceuticals", ""),
    (("有色金属冶炼和压延加工业",), "FAI_Non_Ferrous_Smelting", "Non-ferrous metal smelting and rolling", ""),
    (("金属制品业",), "FAI_Metal_Products", "Metal products", ""),
    (("通用设备制造业",), "FAI_General_Equipment", "General-purpose equipment", ""),
    (("专用设备制造业",), "FAI_Special_Equipment", "Special-purpose equipment", ""),
    (("汽车制造业",), "FAI_Automobiles", "Automobiles", ""),
    (("铁路、船舶、航空航天和其他运输设备制造业",), "FAI_Other_Transport_Equipment", "Rail, ship, aerospace and other transport", ""),
    (("电气机械和器材制造业",), "FAI_Electrical_Machinery", "Electrical machinery and equipment", ""),
    (("计算机、通信和其他电子设备制造业",), "FAI_Electronics", "Computers, communication and electronics", ""),
    (("电力、热力、燃气及水生产和供应业",), "FAI_Utilities", "Power, heat, gas and water supply", I),
    (("交通运输、仓储和邮政业",), "FAI_Transport", "Transport, storage and post", ""),
    (("铁路运输业",), "FAI_Rail", "Railways", ""),
    (("道路运输业",), "FAI_Road", "Road transport", ""),
    (("水利、环境和公共设施管理业",), "FAI_Water_Environment", "Water, environment and public utilities", ""),
    (("水利管理业",), "FAI_Water_Management", "Water management", ""),
    (("公共设施管理业",), "FAI_Public_Utilities", "Public utilities management", ""),
    (("教育",), "FAI_Education", "Education", ""),
    (("卫生和社会工作",), "FAI_Health", "Health and social work", ""),
    (("文化、体育和娱乐业",), "FAI_Culture_Sports", "Culture, sports and entertainment", ""),
    (("内资企业",), "FAI_Domestic", "Domestic-funded enterprises", ""),
    (("港澳台商投资企业", "港澳台投资企业"), "FAI_HMT", "Hong Kong, Macao and Taiwan invested", ""),
    (("外商投资企业",), "FAI_Foreign", "Foreign invested", ""),
]
BY_NAME = {n: stem for names, stem, *_ in SERIES for n in names}
# text: "第二产业投资113094亿元，下降2.9%" / "工业投资同比下降2.9%" / "基础设施投资(...)同比下降4.0%" / "东部地区投资同比下降9.4%"
LEVELS = [("FAI_Total", r"全国固定资产投资\(不含农户\)(\d+)亿元"), ("FAI_Primary", r"第一产业投资(\d+)亿元"),
          ("FAI_Secondary", r"第二产业投资(\d+)亿元"), ("FAI_Tertiary", r"第三产业投资(\d+)亿元")]
TEXT_YOY = [("FAI_Industry", r"工业投资同比(增长|下降)([\d.]+)%"),
            ("FAI_Infrastructure", r"基础设施投资(?:\([^)]*\))?同比(增长|下降)([\d.]+)%"),
            ("FAI_East", r"东部地区投资(?:同比)?(增长|下降)([\d.]+)%"), ("FAI_Central", r"中部地区投资(?:同比)?(增长|下降)([\d.]+)%"),
            ("FAI_West", r"西部地区投资(?:同比)?(增长|下降)([\d.]+)%"), ("FAI_Northeast", r"东北地区投资(?:同比)?(增长|下降)([\d.]+)%")]
COLUMNS = ([f"{s}_YTD_YoY_pct" for _n, s, *_ in SERIES] + [f"{s}_YTD_bn_yuan" for s, _ in LEVELS]
           + [f"{s}_YTD_YoY_pct" for s, _ in TEXT_YOY] + ["FAI_Total_MoM_pct"])


def parse_release(html, _is_jf=False):
    row, unmatched = {}, []
    for cells in nbs.table_rows(html):
        if len(cells) != 2:
            continue
        v = nbs.num(cells[1])
        if v is None:
            continue
        stem = BY_NAME.get(tbl.clean_name(cells[0]))
        if stem is None:
            if not re.fullmatch(r"\d{1,2}月", tbl.clean_name(cells[0])):   # the table of revised m/m rates is not read
                unmatched.append(tbl.clean_name(cells[0]))
        elif f"{stem}_YTD_YoY_pct" not in row:
            row[f"{stem}_YTD_YoY_pct"] = v
    text = re.sub(r"\s+", "", re.sub(r"<[^>]+>", "", html)).replace("（", "(").replace("）", ")")
    for stem, rx in LEVELS:
        m = re.search(rx, text)
        if m:
            row[f"{stem}_YTD_bn_yuan"] = round(float(m.group(1)) * 0.1, 2)
    for stem, rx in TEXT_YOY:
        m = re.search(rx, text)
        if m and f"{stem}_YTD_YoY_pct" not in row:
            row[f"{stem}_YTD_YoY_pct"] = float(m.group(2)) * (1 if m.group(1) == "增长" else -1)
    m = re.search(r"\d{1,2}月份固定资产投资\(不含农户\)(增长|下降)([\d.]+)%", text)
    if m:
        row["FAI_Total_MoM_pct"] = float(m.group(2)) * (1 if m.group(1) == "增长" else -1)
    if unmatched:
        nbs.log(f"    rows not mapped: {sorted(set(unmatched))[:10]}")
    return row


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()
    rows = [(f"{s}_YTD_YoY_pct", f"{lab}: investment, year-to-date y/y", "% y/y (year to date, nominal)", grp, "line")
            for _n, s, lab, grp in SERIES]
    rows += [(f"{s}_YTD_bn_yuan", f"{s[4:].replace('_', ' ')} investment, year to date", "bn yuan (year to date)", "", "line")
             for s, _ in LEVELS]
    rows += [(f"{s}_YTD_YoY_pct", f"{s[4:].replace('_', ' ')} investment, year-to-date y/y (text of the release)",
              "% y/y (year to date, nominal)", "", "line") for s, _ in TEXT_YOY]
    rows.append(("FAI_Total_MoM_pct", "Total investment, month-on-month (seasonally adjusted, text of the release)",
                 "% m/m", "", "line"))
    tbl.pull(args.out, title_re=TITLE_RE, near_re=r"固定资产投资", period_fn=tbl.ytd_period, extra=EXTRA_RELEASES,
             parse_fn=parse_release, columns=COLUMNS, series_rows=rows, notes_lines=NOTES_LINES,
             notes_titles=NOTES_SECTION_TITLES, history_start=HISTORY_START, probe_col="FAI_Total_YTD_YoY_pct",
             skip_gap_months=(1,), index_name="period_end_month")


NOTES_LINES = [
    "UNITS",
    "*_YTD_YoY_pct: y/y growth of investment for the year to date, nominal (price effects not removed), on NBS's "
    "like-for-like basis. The row dated month N covers January to N (February = January-February, the first release of "
    "the year; no January row; December = full year). *_YTD_bn_yuan: year-to-date investment in billion yuan (NBS prints "
    "100 million yuan; x 0.1), from the text. FAI_Total_MoM_pct: the month's seasonally adjusted change quoted in the text. "
    "Industry rows appear as NBS lists them (the list changes over time). Real-estate development investment is a separate "
    "NBS release and is not pulled.",
    "",
    "SOURCE",
    "National Bureau of Statistics of China, monthly release '...全国固定资产投资...', Chinese release list "
    "https://www.stats.gov.cn/sj/zxfb/.",
    "",
    "UPDATES",
    "Incremental: periods already held are not re-fetched.",
]
NOTES_SECTION_TITLES = {"UNITS", "SOURCE", "UPDATES"}


if __name__ == "__main__":
    main()
