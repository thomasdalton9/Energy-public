"""
Shared pull logic for NBS's monthly industrial-production release table
"规模以上工业主要产品产量" (output of major industrial products, industrial
enterprises above designated size). Used by CHINA_NBS_INDUSTRIAL_OUTPUT.py
(every product), CHINA_NBS_ENERGY_PRODUCTION.py and
CHINA_NBS_CLEAN_ENERGY_PRODUCTS.py (subsets). Not a pull script itself.

Each table row is [product (unit), month value, month y/y %, YTD value,
YTD y/y %]; the combined January-February release has [product (unit),
Jan-Feb value, y/y %]. Only the first value is kept. NBS publishes no
separate January or February figures: the Jan-Feb total goes to its own
sheet, and Jan/Feb stay blank in the monthly Data sheet.

Incremental: months already in the workbook are not re-fetched; the
release list is only crawled deep when the workbook has gaps.
"""

import os
import re

import pandas as pd

import china_nbs_common as nbs
import xlsx_notes

TITLE_RE = r"\d{4}年(\d{1,2}|1[—\-－～~]2)月份?(规模以上|规上)工业(增加值|生产)"
NEAR_RE = r"\d{1,2}月份?.*工业(增加值|生产(?!者))"
HISTORY_START = "2021-01-01"

# Releases that still exist but are no longer on the (~1000-item) release list,
# found by CHINA_NBS_DISCOVERY5.py's ID scan of the Feb-2023 site migration.
_OLD = "https://www.stats.gov.cn/sj/zxfb/202302/t20230203_{}.html"
EXTRA_RELEASES = [(t, _OLD.format(i)) for t, i in [
    ("2021年1—2月份规模以上工业增加值增长35.1%", 1901025),
    ("2021年3月份规模以上工业增加值增长14.1%", 1901047),
    ("2021年4月份规模以上工业增加值增长9.8%", 1901099),
    ("2021年5月份规模以上工业增加值增长8.8%", 1901126),
    ("2021年6月份规模以上工业增加值增长8.3%", 1901155),
    ("2021年7月份规模以上工业增加值增长6.4%", 1901191),
    ("2021年8月份规模以上工业增加值增长5.3%", 1901218),
    ("2021年9月份规模以上工业增加值增长3.1%", 1901241),
]]

# (regex on the product name with any "其中：" prefix and unit removed, column, English label,
#  output unit, source unit, factor source->output, chart group)
PRODUCTS = [
    (r"^火力?发电量$|^火电$", "Thermal_Generation_TWh", "Thermal", "TWh", "亿千瓦时", 0.1, "power generation by source"),
    (r"^水力?发电量$|^水电$", "Hydro_Generation_TWh", "Hydro", "TWh", "亿千瓦时", 0.1, "power generation by source"),
    (r"^核能?发电量$|^核电$", "Nuclear_Generation_TWh", "Nuclear", "TWh", "亿千瓦时", 0.1, "power generation by source"),
    (r"^风力?发电量$|^风电$", "Wind_Generation_TWh", "Wind", "TWh", "亿千瓦时", 0.1, "power generation by source"),
    (r"^太阳能发电量$", "Solar_Generation_TWh", "Solar", "TWh", "亿千瓦时", 0.1, "power generation by source"),
    (r"^原煤$", "Raw_Coal_Mt", "Raw coal", "Mt", "万吨", 0.01, "raw coal output"),
    (r"^焦炭$", "Coke_Mt", "Coke", "Mt", "万吨", 0.01, "coke output"),
    (r"^原油$", "Crude_Oil_Mt", "Crude oil", "Mt", "万吨", 0.01, "crude oil output and refinery runs"),
    (r"^原油加工量$", "Crude_Oil_Processing_Mt", "Crude oil processed (refinery runs)", "Mt", "万吨", 0.01,
     "crude oil output and refinery runs"),
    (r"^天然气$", "Natural_Gas_Bcm", "Natural gas", "bcm", "亿立方米", 0.1, "natural gas output"),
    (r"^(规模以上工业)?发电量$", "Total_Generation_TWh", "Electricity generation (total)", "TWh", "亿千瓦时", 0.1,
     "electricity generation (total)"),
    (r"^太阳能电池", "Solar_Cells_GW", "Solar cells (PV cells)", "GW", "万千瓦", 0.01,
     "solar cell and power equipment output"),
    (r"^发电机组", "Power_Generation_Equipment_GW", "Power generation equipment (generator sets)", "GW", "万千瓦", 0.01,
     "solar cell and power equipment output"),
    (r"^汽车$", "Motor_Vehicles_k_units", "Motor vehicles", "thousand units", "万辆", 10, "vehicle output"),
    (r"^轿车$", "Cars_k_units", "Cars (sedans)", "thousand units", "万辆", 10, "vehicle output"),
    (r"SUV", "SUVs_k_units", "SUVs", "thousand units", "万辆", 10, "vehicle output"),
    (r"^新能源汽车$", "New_Energy_Vehicles_k_units", "New energy vehicles", "thousand units", "万辆", 10, "vehicle output"),
    (r"^粗钢$", "Crude_Steel_Mt", "Crude steel", "Mt", "万吨", 0.01, "iron and steel output"),
    (r"^生铁$", "Pig_Iron_Mt", "Pig iron", "Mt", "万吨", 0.01, "iron and steel output"),
    (r"^钢材$", "Finished_Steel_Mt", "Finished steel products", "Mt", "万吨", 0.01, "iron and steel output"),
    (r"^水泥$", "Cement_Mt", "Cement", "Mt", "万吨", 0.01, "cement output"),
    (r"^平板玻璃$", "Plate_Glass_M_cases", "Plate glass", "million weight cases", "万重量箱", 0.01, "plate glass output"),
    (r"^十种有色金属$", "Ten_Nonferrous_Metals_Mt", "Ten non-ferrous metals", "Mt", "万吨", 0.01, "non-ferrous metals output"),
    (r"^原铝|^电解铝", "Primary_Aluminium_Mt", "Primary aluminium (electrolytic)", "Mt", "万吨", 0.01,
     "non-ferrous metals output"),
    (r"^乙烯$", "Ethylene_Mt", "Ethylene", "Mt", "万吨", 0.01, "chemicals output"),
    (r"^硫酸", "Sulfuric_Acid_Mt", "Sulfuric acid (100%)", "Mt", "万吨", 0.01, "chemicals output"),
    (r"^烧碱", "Caustic_Soda_Mt", "Caustic soda (100%)", "Mt", "万吨", 0.01, "chemicals output"),
    (r"^化学纤维$", "Chemical_Fibre_Mt", "Chemical fibre", "Mt", "万吨", 0.01, "chemicals output"),
    (r"^布$", "Cloth_bn_m", "Cloth", "billion metres", "亿米", 0.1, "cloth output"),
    (r"^金属切削机床$", "Metal_Cutting_Machine_Tools_k_units", "Metal-cutting machine tools", "thousand units", "万台",
     10, "machine tool output"),
    (r"^工业机器人$", "Industrial_Robots_k_sets", "Industrial robots", "thousand sets", "套", 0.001, "industrial robot output"),
    (r"^服务机器人$", "Service_Robots_k_sets", "Service robots", "thousand sets", "套", 0.001, "service robot output"),
    (r"^微型计算机", "Microcomputers_M_units", "Microcomputers", "million units", "万台", 0.01, "electronics output"),
    (r"^移动通信手持机$|^手机$", "Mobile_Phones_M_units", "Mobile phones", "million units", "万台", 0.01, "electronics output"),
    (r"^智能手机$", "Smartphones_M_units", "Smartphones", "million units", "万台", 0.01, "electronics output"),
    (r"^集成电路$", "Integrated_Circuits_bn_units", "Integrated circuits", "billion units", "亿块", 0.1,
     "integrated circuit output"),
]
STACKED_GROUPS = {"power generation by source"}


def parse_release(html):
    """{column: value} from the product-output table of one release."""
    row, started = {}, False
    for cells in nbs.table_rows(html):
        if any("主要产品产量" in c for c in cells):
            started = True
            continue
        if not started or len(cells) < 2:
            continue
        name, unit = nbs.norm_label(cells[0])
        value = nbs.num(cells[1])
        if value is None:
            continue
        for rx, col, _label, _u, src_unit, factor, _g in PRODUCTS:
            if col in row or not re.search(rx, name):
                continue
            if unit and unit != src_unit:
                nbs.log(f"    {col}: unexpected unit {unit!r} (expected {src_unit!r}) - skipped")
                break
            row[col] = round(value * factor, 6)
            break
    return row


def pull(out_path, columns, notes_lines, notes_titles):
    cols = [p[1] for p in PRODUCTS if p[1] in columns]
    data = janfeb = None
    if nbs.has_sheet(out_path, "Series"):   # earlier (English-release) layout -> rebuild from scratch
        data = nbs.read_sheet(out_path, "Data")
        janfeb = nbs.read_sheet(out_path, "Jan-Feb")
    # reindex: a workbook written with older column names is simply re-pulled in full
    data = pd.DataFrame(columns=cols, dtype=float) if data is None else data.reindex(columns=cols)
    janfeb = pd.DataFrame(columns=cols, dtype=float) if janfeb is None else janfeb.reindex(columns=cols)
    held = {(p.year, p.month) for p in data.index[data.notna().any(axis=1)]} if len(data) else set()
    held_jf = {p.year for p in janfeb.index[janfeb.notna().any(axis=1)]} if len(janfeb) else set()
    gaps = [g for g in nbs.recent_gaps(data.index[data.notna().any(axis=1)], "MS") if g.month not in (1, 2)]
    deep = (not held) or bool(gaps)
    nbs.log(f"  held {len(held)} months + {len(held_jf)} Jan-Feb; gaps {len(gaps)} -> {'full' if deep else 'shallow'} crawl")

    def wanted(title):
        period, is_jf = nbs.month_period(title)
        if period is None or period < pd.Timestamp(HISTORY_START):
            return False
        return (period.year not in held_jf) if is_jf else ((period.year, period.month) not in held)

    releases = nbs.crawl_index(TITLE_RE, wanted, deep, EXTRA_RELEASES, near_re=NEAR_RE)
    nbs.log(f"  {len(releases)} release(s) to fetch")
    new_rows, new_jf = {}, {}
    for title, url in releases:
        period, is_jf = nbs.month_period(title)
        try:
            html = nbs.fetch(url)
        except Exception as e:  # noqa: BLE001 - keep going, the next run retries
            nbs.log(f"  [{period:%Y-%m}] fetch failed ({type(e).__name__}) - skipped")
            continue
        row = parse_release(html or "")
        row = {k: v for k, v in row.items() if k in cols}
        if not row:
            nbs.log(f"  [{period:%Y-%m}] no product rows matched - skipped ({url})")
            continue
        nbs.log(f"  [{period:%Y-%m}{' Jan-Feb' if is_jf else ''}] {len(row)}/{len(cols)} products")
        (new_jf if is_jf else new_rows)[period] = row

    if new_rows:
        data = data.combine_first(pd.DataFrame.from_dict(new_rows, orient="index")) if len(data) else \
            pd.DataFrame.from_dict(new_rows, orient="index")
    if new_jf:
        janfeb = janfeb.combine_first(pd.DataFrame.from_dict(new_jf, orient="index")) if len(janfeb) else \
            pd.DataFrame.from_dict(new_jf, orient="index")
    if data.dropna(how="all").empty:
        raise SystemExit("No data at all - nothing to save.")
    data = data.reindex(columns=cols).sort_index()
    data = data[data.index >= pd.Timestamp(HISTORY_START)]
    janfeb = janfeb.reindex(columns=cols).sort_index()
    data.index.name = janfeb.index.name = "month"
    series = nbs.series_sheet([(p[1], p[2], f"{p[3]} per month", p[6], "stacked_bar" if p[6] in STACKED_GROUPS else "line")
                               for p in PRODUCTS if p[1] in cols])
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    xlsx_notes.write_workbook(out_path, {"Data": data, "Jan-Feb": janfeb, "Series": series},
                              notes_lines, notes_titles)
    print(f"Saved {len(data)} month(s) ({data.index.min():%Y-%m}..{data.index.max():%Y-%m}), "
          f"{len(janfeb)} Jan-Feb total(s), {len(new_rows) + len(new_jf)} new, to {out_path}")


def units_line(columns):
    return "; ".join(f"{p[1]}: {p[2]}, {p[3]} per month (source unit {p[4]})" for p in PRODUCTS if p[1] in columns)


COMMON_NOTES = [
    "SCOPE",
    "Industrial enterprises above designated size (annual main business revenue of RMB 20 million or more), "
    "NBS's own basis for these figures - not total national output. The sample of enterprises changes each "
    "year, so NBS's published year-on-year growth rates are computed on a like-for-like basis and will not "
    "always match the change between the levels shown here.",
    "",
    "JANUARY AND FEBRUARY",
    "NBS does not publish separate January or February output: it releases one combined January-February "
    "figure (around mid-March). The monthly Data sheet therefore leaves January and February blank; the "
    "combined totals are on the 'Jan-Feb' sheet (dated 1 Feb of each year).",
    "",
    "SOURCE",
    "National Bureau of Statistics of China, monthly industrial production release ('YYYY年M月份规模以上工业增加值"
    "增长X%'), table 'Output of Major Industrial Products' (规模以上工业主要产品产量), Chinese release list "
    "https://www.stats.gov.cn/sj/zxfb/ (the list reaches back to late 2021; data.stats.gov.cn is blocked to "
    "automated requests - 403 UrlACL). Only the monthly value is kept (the release also gives y/y % and "
    "year-to-date figures).",
    "",
    "UPDATES",
    "Incremental: months already held are not re-fetched; each run adds newly published months and fills "
    "any gaps still available on the release list.",
]
