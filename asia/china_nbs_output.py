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

import china_nbs_archive_ids as arch
import china_nbs_common as nbs
import xlsx_notes

TITLE_RE = r"\d{4}年((\d{1,2}|1[—\-－～~]\d{1,2})月份?|上半年|前三季度)(规模以上|规上)工业(增加值|生产)"
NEAR_RE = r"(月份?|半年|季度).*工业(增加值|生产(?!者))"
HISTORY_START = "2013-01-01"
EXTRA_RELEASES = arch.INDUSTRIAL   # releases no longer on the list, from the ID scan
ENERGY_EXTRA = arch.ENERGY   # energy-production (能源生产情况) releases no longer on the list

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
    (r"^(天然)?原油$", "Crude_Oil_Mt", "Crude oil", "Mt", "万吨", 0.01, "crude oil output and refinery runs"),
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
    (r"^氧化铝", "Alumina_Mt", "Alumina (to 2015)", "Mt", "万吨", 0.01, "non-ferrous metals output"),
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
    (r"^产品销售率", "Sales_Rate_pct", "Product sales rate", "%", "%", 1, "product sales rate"),
    (r"^出口交货值", "Export_Delivery_Value_bn_yuan", "Export delivery value", "bn yuan", "亿元", 0.1,
     "export delivery value"),
]



STACKED_GROUPS = {"power generation by source"}


VA_ROWS = [
    # (Chinese name exactly as in the release, column stem, English label, chart group)
    ("规模以上工业增加值", "VA_Industry", "Industry (all, above designated size)", "industrial value added y/y"),
    ("采矿业", "VA_Mining", "Mining", "industrial value added y/y"),
    ("制造业", "VA_Manufacturing", "Manufacturing", "industrial value added y/y"),
    ("高技术制造业", "VA_High_Tech_Manufacturing", "High-tech manufacturing", ""),
    ("电力、热力、燃气及水生产和供应业", "VA_Utilities", "Power, heat, gas and water supply", "industrial value added y/y"),
    ("国有控股企业", "VA_State_Controlled", "State-controlled enterprises", ""),
    ("股份制企业", "VA_Joint_Stock", "Joint-stock enterprises", ""),
    ("外商及港澳台商投资企业", "VA_Foreign_HMT", "Foreign, Hong Kong, Macao and Taiwan invested", ""),
    ("外商及港澳台投资企业", "VA_Foreign_HMT", "Foreign, Hong Kong, Macao and Taiwan invested", ""),
    ("私营企业", "VA_Private", "Private enterprises", ""),
    ("集体企业", "VA_Collective", "Collective enterprises (to 2015)", ""),
    ("煤炭开采和洗选业", "VA_Coal_Mining", "Coal mining and washing", "energy industries value added y/y"),
    ("石油和天然气开采业", "VA_Oil_Gas_Extraction", "Oil and gas extraction", "energy industries value added y/y"),
    ("电力、热力生产和供应业", "VA_Power_Heat", "Power and heat production and supply", "energy industries value added y/y"),
    ("农副食品加工业", "VA_Agri_Food_Processing", "Agricultural food processing", ""),
    ("食品制造业", "VA_Food_Manufacturing", "Food manufacturing", ""),
    ("酒、饮料和精制茶制造业", "VA_Beverages", "Liquor, beverages and tea", ""),
    ("纺织业", "VA_Textiles", "Textiles", ""),
    ("化学原料和化学制品制造业", "VA_Chemicals", "Chemical raw materials and products", "heavy industry value added y/y"),
    ("医药制造业", "VA_Pharmaceuticals", "Pharmaceuticals", ""),
    ("橡胶和塑料制品业", "VA_Rubber_Plastics", "Rubber and plastics", ""),
    ("非金属矿物制品业", "VA_Non_Metallic_Minerals", "Non-metallic mineral products", "heavy industry value added y/y"),
    ("黑色金属冶炼和压延加工业", "VA_Ferrous_Smelting", "Ferrous metal smelting and rolling (steel)",
     "heavy industry value added y/y"),
    ("有色金属冶炼和压延加工业", "VA_Non_Ferrous_Smelting", "Non-ferrous metal smelting and rolling",
     "heavy industry value added y/y"),
    ("金属制品业", "VA_Metal_Products", "Metal products", ""),
    ("通用设备制造业", "VA_General_Equipment", "General-purpose equipment", ""),
    ("专用设备制造业", "VA_Special_Equipment", "Special-purpose equipment", ""),
    ("汽车制造业", "VA_Automobiles", "Automobiles", ""),
    ("铁路、船舶、航空航天和其他运输设备制造业", "VA_Other_Transport_Equipment", "Rail, ship, aerospace and other transport", ""),
    ("电气机械和器材制造业", "VA_Electrical_Machinery", "Electrical machinery and equipment", ""),
    ("计算机、通信和其他电子设备制造业", "VA_Electronics", "Computers, communication and electronics", ""),
]
VA_BY_NAME = {cn: stem for cn, stem, *_ in VA_ROWS}
VA_LABEL = {stem: (label, group) for _cn, stem, label, group in VA_ROWS}
VA_STEMS = list(dict.fromkeys(stem for _cn, stem, *_ in VA_ROWS))
ENERGY_VA = ["VA_Industry", "VA_Mining", "VA_Manufacturing", "VA_Utilities", "VA_Coal_Mining", "VA_Oil_Gas_Extraction",
             "VA_Power_Heat"]
# product groups of the y/y charts (products not listed here have y/y data but no chart)
YOY_GROUPS = {
    "power generation y/y by source": ["Total_Generation_TWh", "Thermal_Generation_TWh", "Hydro_Generation_TWh",
                                       "Nuclear_Generation_TWh", "Wind_Generation_TWh", "Solar_Generation_TWh"],
    "fuel output y/y": ["Raw_Coal_Mt", "Coke_Mt", "Crude_Oil_Mt", "Crude_Oil_Processing_Mt", "Natural_Gas_Bcm"],
    "heavy industry output y/y": ["Crude_Steel_Mt", "Pig_Iron_Mt", "Finished_Steel_Mt", "Cement_Mt", "Plate_Glass_M_cases",
                                  "Primary_Aluminium_Mt", "Ten_Nonferrous_Metals_Mt", "Ethylene_Mt"],
    "equipment output y/y": ["Solar_Cells_GW", "Power_Generation_Equipment_GW", "Motor_Vehicles_k_units",
                             "New_Energy_Vehicles_k_units", "Integrated_Circuits_bn_units", "Industrial_Robots_k_sets"],
}
SUFFIXES = ("YoY_pct", "YTD", "YTD_YoY_pct")
UNIT_ALIAS = {"台/套": "套"}   # industrial robots were reported in 台/套 (units/sets) until 2020


def _pct(s):
    """'-0.1(percentage points)' / '…' / '-' -> float or None."""
    if s is None:
        return None
    s = re.sub(r"[（(].*?[）)]", "", str(s)).strip()
    return nbs.num(s)


def parse_tables(rows):
    """{'prod': {col: (value, yoy, ytd, ytd_yoy)}, 'va': {stem: (yoy, ytd_yoy)}} from the release's table rows.
    Month release rows: [name, month value, month y/y, YTD value, YTD y/y]; January-February release: [name, value, y/y]."""
    prod, va, started = {}, {}, False
    for cells in rows:
        periods = [c for c in cells if re.fullmatch(r"(1[—\-－~～])?\d{1,2}月", c)]
        if not started and periods and re.match(r"1[—\-－~～]([3-9]|1[0-2])月", periods[0]):
            nbs.log(f"    table holds only year-to-date columns ({periods}) - no monthly figures")
            return {"prod": {}, "va": {}}
        if any("主要产品产量" in c for c in cells):
            started = True
            continue
        if len(cells) < 3:
            continue
        name, unit = nbs.norm_label(cells[0])
        name = re.sub(r"^[一二三四五六七八九十]+、", "", name.replace("其中:", ""))
        if not started:
            stem = VA_BY_NAME.get(name)
            if stem and stem not in va:
                va[stem] = (_pct(cells[2]), _pct(cells[4]) if len(cells) > 4 else None)
            continue
        value = nbs.num(cells[1])
        if value is None:
            continue
        for rx, col, _label, _u, src_unit, factor, _g in PRODUCTS:
            if col in prod or not re.search(rx, name):
                continue
            if unit and unit != src_unit and UNIT_ALIAS.get(unit) != src_unit:
                nbs.log(f"    {col}: unexpected unit {unit!r} (expected {src_unit!r}) - skipped")
                break
            ytd = nbs.num(cells[3]) if len(cells) > 3 else None
            prod[col] = (round(value * factor, 6), _pct(cells[2]), None if ytd is None else round(ytd * factor, 6),
                         _pct(cells[4]) if len(cells) > 4 else None)
            break
    return {"prod": prod, "va": va}


def parse_release(html):
    """{column: monthly value} (kept for callers that only need the level)."""
    return {c: v[0] for c, v in parse_tables(nbs.table_rows(html))["prod"].items()}


def flatten(parsed, cols, with_va):
    """One wide row: product value columns as before, plus <col>_YoY_pct / _YTD / _YTD_YoY_pct and VA_<x>_YoY_pct."""
    row = {}
    for col, (v, y, ytd, ytd_y) in parsed["prod"].items():
        if col not in cols:
            continue
        row[col] = v
        if col == "Sales_Rate_pct":   # NBS gives its change in percentage points, not a y/y rate: only the monthly level is kept
            y = ytd = ytd_y = None
        for suf, val in zip(SUFFIXES, (y, ytd, ytd_y)):
            if val is not None:
                row[f"{col}_{suf}"] = val
    if with_va:
        for stem, (y, ytd_y) in parsed["va"].items():
            if stem in with_va:
                if y is not None:
                    row[f"{stem}_YoY_pct"] = y
                if ytd_y is not None:
                    row[f"{stem}_YTD_YoY_pct"] = ytd_y
    return row


def all_columns(cols, with_va):
    out = list(cols)
    for c in cols:
        out += [f"{c}_{s}" for s in SUFFIXES if c != "Sales_Rate_pct"]
    for stem in with_va or ():
        out += [f"{stem}_YoY_pct", f"{stem}_YTD_YoY_pct"]
    return out


def series_rows(cols, with_va):
    """Series sheet rows (column, label, unit, chart group, kind) for the base, y/y, YTD and value-added columns."""
    rows = []
    for p in PRODUCTS:
        if p[1] not in cols:
            continue
        rows.append((p[1], p[2], f"{p[3]} per month", p[6], "stacked_bar" if p[6] in STACKED_GROUPS else "line"))
    for p in PRODUCTS:
        if p[1] not in cols or p[1] == "Sales_Rate_pct":   # NBS gives only a percentage-point change for the sales rate
            continue
        grp = next((g for g, members in YOY_GROUPS.items() if p[1] in members), "")
        rows.append((f"{p[1]}_YoY_pct", f"{p[2]}, y/y", "% y/y", grp, "line"))
        rows.append((f"{p[1]}_YTD", f"{p[2]}, year to date", f"{p[3]} year to date", "", "line"))
        rows.append((f"{p[1]}_YTD_YoY_pct", f"{p[2]}, year-to-date y/y", "% y/y (as published by NBS)", "", "line"))
    for stem in with_va or ():
        label, grp = VA_LABEL[stem]
        rows.append((f"{stem}_YoY_pct", f"Value added: {label}, y/y", "% y/y", grp, "line"))
        rows.append((f"{stem}_YTD_YoY_pct", f"Value added: {label}, year-to-date y/y", "% y/y",
                     "", "line"))
    return rows


def pull(out_path, columns, notes_lines, notes_titles, with_va=None, extra=None):
    """with_va: value-added stems to keep (None = none). extra: optional hook(releases_state) for subclasses (unused)."""
    cols = [p[1] for p in PRODUCTS if p[1] in columns]
    wide = all_columns(cols, with_va)
    data = janfeb = None
    if nbs.has_sheet(out_path, "Series"):   # earlier (English-release) layout -> rebuild from scratch
        data = nbs.read_sheet(out_path, "Data")
        janfeb = nbs.read_sheet(out_path, "Jan-Feb")
    # reindex: a workbook written with older column names is simply re-pulled in full
    data = pd.DataFrame(columns=wide, dtype=float) if data is None else data.reindex(columns=wide)
    janfeb = pd.DataFrame(columns=wide, dtype=float) if janfeb is None else janfeb.reindex(columns=wide)
    probe = [c for c in wide if c.endswith("_YoY_pct") and not c.startswith("VA_")]
    # a month counts as held once any y/y column is in (workbooks written before the y/y columns held levels only)
    held = {(p.year, p.month) for p in data.index[data[probe].notna().any(axis=1)]} if len(data) else set()
    held_jf = {p.year for p in janfeb.index[janfeb[probe].notna().any(axis=1)]} if len(janfeb) else set()
    gaps = [g for g in nbs.recent_gaps(data.index[data.notna().any(axis=1)], "MS") if g.month not in (1, 2)]
    old_levels = len(data) and data.index[data[cols].notna().any(axis=1) & data[probe].isna().all(axis=1)].size > 0
    deep = (not held) or bool(gaps) or bool(old_levels)
    nbs.log(f"  held {len(held)} months + {len(held_jf)} Jan-Feb; gaps {len(gaps)}; levels-only {bool(old_levels)} -> "
            f"{'full' if deep else 'shallow'} crawl")

    def wanted(title):
        period, is_jf = nbs.month_period(title)
        if period is None or period < pd.Timestamp(HISTORY_START):
            return False
        return (period.year not in held_jf) if is_jf else ((period.year, period.month) not in held)

    releases = nbs.crawl_index(TITLE_RE, wanted, deep, EXTRA_RELEASES, near_re=NEAR_RE)
    nbs.log(f"  {len(releases)} release(s) to fetch")
    new_rows, new_jf = {}, {}
    for title, url in releases:
        if nbs.over_budget():
            nbs.log("  time budget reached - the remaining releases are picked up by the next run")
            break
        period, is_jf = nbs.month_period(title)
        try:
            html = nbs.fetch(url)
        except Exception as e:  # noqa: BLE001 - keep going, the next run retries
            nbs.log(f"  [{period:%Y-%m}] fetch failed ({type(e).__name__}) - skipped")
            continue
        parsed = parse_tables(nbs.table_rows(html or ""))
        row = flatten(parsed, cols, set(with_va or ()))
        if not any(k in row for k in cols):
            nbs.log(f"  [{period:%Y-%m}] no product rows matched - skipped ({url})")
            continue
        nbs.log(f"  [{period:%Y-%m}{' Jan-Feb' if is_jf else ''}] {sum(k in row for k in cols)}/{len(cols)} products, "
                f"{sum(k.startswith('VA_') and k.endswith('_YoY_pct') for k in row)} value-added rows")
        (new_jf if is_jf else new_rows)[period] = row

    if new_rows:
        data = data.combine_first(pd.DataFrame.from_dict(new_rows, orient="index")) if len(data) else \
            pd.DataFrame.from_dict(new_rows, orient="index")
    if new_jf:
        janfeb = janfeb.combine_first(pd.DataFrame.from_dict(new_jf, orient="index")) if len(janfeb) else \
            pd.DataFrame.from_dict(new_jf, orient="index")
    if data.dropna(how="all").empty:
        raise SystemExit("No data at all - nothing to save.")
    data = data.reindex(columns=wide).sort_index()
    data = data[data.index >= pd.Timestamp(HISTORY_START)]
    janfeb = janfeb.reindex(columns=wide).sort_index()
    data.index.name = janfeb.index.name = "month"
    series = nbs.series_sheet(series_rows(cols, with_va))
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    sheets = {"Data": data, "Jan-Feb": janfeb, "Series": series}
    if extra:
        sheets.update(extra(data, janfeb))
    xlsx_notes.write_workbook(out_path, sheets, notes_lines, notes_titles)
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
    "Y/Y AND YEAR-TO-DATE COLUMNS",
    "Beside each level the release's own month-on-year growth (<name>_YoY_pct, %), the year-to-date total "
    "(<name>_YTD, same unit as the level) and the year-to-date growth (<name>_YTD_YoY_pct) are kept exactly as NBS "
    "publishes them. NBS computes growth on a like-for-like enterprise sample, so it differs from the change between the "
    "levels shown. VA_* columns (industrial workbook and the energy-production workbook's energy subset) are NBS's "
    "real value-added growth, % y/y, by sector, ownership and industry. The product sales rate is a level in %; NBS gives "
    "its change in percentage points, which is not kept. Industry rows (e.g. coal mining, oil and gas) appear in the "
    "release from the January-February 2022 issue and the solar-cell row from January-February 2023 (service robots: "
    "January-February 2025); earlier months are blank because NBS did not publish those rows.",
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
