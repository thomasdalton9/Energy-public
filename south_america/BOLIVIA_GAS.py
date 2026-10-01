"""
Bolivia natural gas: domestic demand by sector, production, and exports
to Brazil and Argentina, monthly from 2021, all from the national
statistics institute (INE).

1. Demand by sector (domestic market): "Bolivia - Volumen Comercializado
   de Gas Natural segun Red de Distribucion segun Ano y Mes" (sheet H05),
   linked from
   https://www.ine.gob.bo/index.php/estadisticas-economicas/hidrocarburos-mineria/hidrocarburo-cuadros-estadisticos/
   Sectors as published: GNV (vehicle gas), Comercial, Domestico
   (residential), Industrial, Generadoras electricas (power), Total.
   Millions of cubic metres per month (the unit is read from the sheet
   header and printed on every run). Exports are NOT included.
2. Production: "Bolivia - Produccion de Gas Natural por Departamento
   segun Ano y Mes" (sheet H02, same page), million m3 per month by
   department (Chuquisaca, Cochabamba, Tarija, Santa Cruz) and total.
   INE/YPFB note: production = pipeline deliveries + liquids + fuel.
   INE publishes no field-level split (San Alberto, Sabalo,
   Margarita-Huacaya, Incahuasi... are in Tarija, Chuquisaca and Santa
   Cruz).
3. Exports: INE's export micro-data (customs records, one workbook per
   year: "EXPORTACIONES 2021" ... "EXPORTACIONES ENE A AGO 2026p"), from
   https://www.ine.gob.bo/index.php/estadisticas-economicas/comercio-exterior/bases-de-datos-exportaciones/
   Natural gas is NANDINA 2711210000 ("gas natural en estado gaseoso");
   rows are summed by month and destination (BRASIL via GASBOL, ARGENTINA
   via GJA / Yacuiba-Pocitos). The FINO column holds the volume in m3
   (checked on every run: KILNET / FINO must be a natural-gas density,
   0.6-0.9 kg/m3, else that year is left out); VALOR is FOB US$.
   Found via discovery_archive/south_america/BOLIVIA_GAS_EXPORTS_DISCOVERY*.py.
   Incremental: final years already complete in the workbook are kept;
   preliminary (p) years and missing years are downloaded again.

The year row (annual total) in the H tables is skipped; INE publishes with
a lag and marks the latest years preliminary (p).
Sheets: 'Demand by sector', 'Production and exports' (million m3/month
and million m3/day), 'Exports (INE customs)' (raw monthly sums) and
'Export cross-check' (INE exports vs Brazil ANP and Argentina ENARGAS
import data, when those workbooks are present in the repo).
Not reachable from the editing sandbox; runs in GitHub Actions.
"""
import argparse
import datetime
import io
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root
import xlsx_charts
import xlsx_notes

PAGE = ("https://www.ine.gob.bo/index.php/estadisticas-economicas/hidrocarburos-mineria/"
        "hidrocarburo-cuadros-estadisticos/")
EXPORTS_PAGE = ("https://www.ine.gob.bo/index.php/estadisticas-economicas/comercio-exterior/"
                "bases-de-datos-exportaciones/")
FALLBACK_URL = "https://nube.ine.gob.bo/index.php/s/Jsm9uMT7L9ALmWr/download"
PRODUCTION_FALLBACK_URL = "https://nube.ine.gob.bo/index.php/s/QQ1WOj4zNwoeFlx/download"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 180)
DATA_START = 2021
GAS_NANDINA = "2711210000"
DENSITY_RANGE = (0.6, 0.9)   # kg/m3: KILNET / FINO must fall here for FINO to be a volume in m3
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output",
                          "Data and Chart Outputs")
MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
         "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
MES3 = {k[:3]: v for k, v in MESES.items()}
RENAME = {"gas natural vehicular": "Vehicle_CNG", "comercial": "Commercial", "domestico": "Residential",
          "industrial": "Industrial", "generadoras electricas": "Power", "total": "Total"}
DEPARTMENTS = {"tarija": "Tarija", "santa cruz": "Santa_Cruz", "chuquisaca": "Chuquisaca",
               "cochabamba": "Cochabamba", "total": "Total"}
EXPORTS_SHEET = "Exports (INE customs)"


def out(*a):
    print(*a, flush=True)


def strip_accents(s):
    return (str(s).lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o")
            .replace("ú", "u").strip())


def page_links(page):
    r = requests.get(page, headers=H, timeout=T)
    r.raise_for_status()
    return [(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", txt)).replace("&#8211;", "-").strip(),
             href.replace("&amp;", "&"))
            for href, txt in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', r.text, re.S | re.I)]


def source_url(pattern, fallback):
    """The download link whose text matches `pattern` on INE's hydrocarbons page (falls back to the known link)."""
    try:
        for txt, href in page_links(PAGE):
            if re.search(pattern, txt, re.I):
                return href
    except requests.RequestException as e:
        out(f"INE page unavailable ({e}); using known link")
    return fallback


def _cell(v):
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else re.sub(r"\s+", " ", strip_accents(v))


def _sector(name):
    if name.startswith("gnv") or "vehicular" in name:
        return "Vehicle_CNG"
    return next((std for key, std in RENAME.items() if name.startswith(key)), None)


def _department(name):
    return next((std for key, std in DEPARTMENTS.items() if name.startswith(key)), None)


def _year_month(v):
    """(year, month) from a label cell: a year ('2023', '2023(p)'), a month name/abbreviation, or a real date."""
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        return v.year, v.month
    if isinstance(v, (int, float)) and not pd.isna(v):
        return (int(v), None) if float(v).is_integer() and 1990 <= v <= 2100 else (None, None)
    s = _cell(v).replace("(p)", "").replace("p/", "").replace("(*)", "").strip(" .*")
    m = re.fullmatch(r"((?:19|20)\d\d)(?:\.0)?\s*p?", s)
    if m:
        return int(m.group(1)), None
    return None, MES3.get(s[:3]) if len(s) >= 3 and s.isalpha() else None


def dump(raw, name, n=30):
    out(f"--- sheet {name!r} {raw.shape}: first {n} rows ---")
    for i in range(min(n, len(raw))):
        out(i, raw.iloc[i].tolist()[:10])


def parse(content, mapper=_sector, prefer="H05", min_cols=3):
    """Find the sheet and row holding the column headers (sectors / departments), then read the year / month
    rows below it. (The first published layout broke a fixed 'Periodo in column A' assumption, so this searches.)"""
    sheets = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None)
    out(f"sheets: {list(sheets)}")
    for name in sorted(sheets, key=lambda s: 0 if s.strip().upper() == prefer else 1):
        raw = sheets[name]
        for i in range(min(40, len(raw))):
            sec = {}
            for j, v in enumerate(raw.iloc[i]):
                std = mapper(_cell(v))
                if std and std not in sec.values():
                    sec[j] = std
            if len(set(sec.values()) - {"Total"}) < min_cols:
                continue
            title = " | ".join(re.sub(r"\s+", " ", v).strip() for r in range(i) for v in raw.iloc[r]
                               if isinstance(v, str) and v.strip())
            dump(raw, name, i + 4)
            out(f"sheet {name!r}, header row {i}: {sec}")
            out(f"sheet header: {title}")
            first = min(sec)
            rows, year = [], None
            for k in range(i + 1, len(raw)):
                r = raw.iloc[k]
                month = None
                for j in range(first):
                    y, m = _year_month(r.iloc[j])
                    year, month = y or year, m or month
                if month and year:
                    rows.append({"Month": pd.Timestamp(year, month, 1),
                                 **{std: pd.to_numeric(r.iloc[j], errors="coerce") for j, std in sec.items()}})
            if rows:
                return pd.DataFrame(rows).groupby("Month").last().sort_index(), title
    for name, raw in sheets.items():
        dump(raw, name)
    raise SystemExit("could not find the header row (sectors / departments) in any sheet")


def unit_scale(title):
    """Factor to million m3 from the unit stated in the sheet header."""
    t = strip_accents(title)
    if re.search(r"miles de metros|mm?3 ?\(miles\)|miles de m", t):
        return 1e-3, "thousand m3"
    if re.search(r"millones de pies|mmpc|mmcf", t):
        return 0.0283168, "million cubic feet"
    if re.search(r"millones de metros|mmm3|mmmc|millones de m", t):
        return 1.0, "million m3"
    return 1.0, "not stated in header - assumed million m3"


def ine_table(pattern, fallback, mapper, prefer, min_cols):
    url = source_url(pattern, fallback)
    out(f"downloading {url}")
    r = requests.get(url, headers=H, timeout=T)
    r.raise_for_status()
    monthly, title = parse(r.content, mapper, prefer, min_cols)
    scale, unit = unit_scale(title)
    out(f"unit from sheet header: {unit} (x{scale} -> million m3)")
    monthly = monthly * scale
    return monthly[monthly.index.year >= DATA_START].dropna(how="all"), title, unit, scale


# ------------------------------------------------------------------ exports (INE customs micro-data)

def export_year_links():
    """{year: (label, url)} for the yearly export micro-data workbooks from DATA_START on."""
    links = {}
    for txt, href in page_links(EXPORTS_PAGE):
        m = re.fullmatch(r"EXPORTACIONES\b.*?((?:19|20)\d\d)\s*(p?)", txt.strip(), re.I)
        if m and int(m.group(1)) >= DATA_START:
            links[int(m.group(1))] = (txt.strip(), href)
    return links


def export_year(url, year):
    """Monthly natural-gas exports by destination from one yearly micro-data workbook.
    Returns (frame indexed by Month with Country, Volume_mcm, Net_weight_t, Value_USD_m, Density_kg_m3,
    months covered by the file)."""
    r = requests.get(url, headers=H, timeout=T)
    r.raise_for_status()
    df = pd.read_excel(io.BytesIO(r.content), sheet_name=0)
    df.columns = [str(c).strip().upper() for c in df.columns]
    need = {"GESTION", "MES", "NANDINA", "DESPAIS", "KILNET", "FINO", "VALOR"}
    if not need <= set(df.columns):
        raise ValueError(f"{year}: columns {sorted(df.columns)} lack {sorted(need - set(df.columns))}")
    df = df[pd.to_numeric(df["GESTION"], errors="coerce") == year]
    months = sorted(pd.to_numeric(df["MES"], errors="coerce").dropna().astype(int).unique())
    code = df["NANDINA"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(10)
    gas = df[code == GAS_NANDINA].copy()
    gas["Country"] = gas["DESPAIS"].astype(str).str.strip().str.title()
    for c in ("KILNET", "FINO", "VALOR"):
        gas[c] = pd.to_numeric(gas[c], errors="coerce")
    g = gas.groupby([pd.to_numeric(gas["MES"]).astype(int), "Country"])[["FINO", "KILNET", "VALOR"]].sum()
    res = pd.DataFrame({"Volume_mcm": g["FINO"] / 1e6, "Net_weight_t": g["KILNET"] / 1e3,
                        "Value_USD_m": g["VALOR"] / 1e6}).reset_index()
    res["Density_kg_m3"] = (res["Net_weight_t"] * 1e3 / (res["Volume_mcm"] * 1e6)).round(4)
    res["Month"] = [pd.Timestamp(year, m, 1) for m in res["MES"]]
    return res.drop(columns="MES").set_index("Month"), months


def load_existing_exports(path):
    try:
        old = pd.read_excel(path, sheet_name=EXPORTS_SHEET)
    except (FileNotFoundError, ValueError) as e:
        out(f"no existing exports sheet ({e}); full backfill")
        return pd.DataFrame()
    old["Month"] = pd.to_datetime(old["Month"].astype(str))
    return old.set_index("Month")


def update_exports(path):
    """Incremental: keep final, complete years from the workbook; download preliminary (p) and missing years."""
    old = load_existing_exports(path)
    links = export_year_links()
    out(f"export micro-data years on INE's page: {[f'{y}: {lab}' for y, (lab, _) in sorted(links.items())]}")
    frames = []
    for year, (label, url) in sorted(links.items()):
        have = old[old.index.year == year] if not old.empty else old
        complete = not have.empty and have.index.month.nunique() == 12
        prelim = label.lower().endswith("p")
        if complete and not prelim:
            out(f"  {year} ({label}): final and complete in the workbook - kept")
            frames.append(have)
            continue
        out(f"  {year} ({label}): downloading {url}")
        try:
            res, months = export_year(url, year)
        except Exception as e:   # keep what we had for that year rather than lose it
            out(f"  {year}: FAILED {type(e).__name__}: {e}")
            if not have.empty:
                frames.append(have)
            continue
        dens = res["Density_kg_m3"].median() if len(res) else float("nan")
        if not DENSITY_RANGE[0] <= dens <= DENSITY_RANGE[1]:
            out(f"  {year}: KILNET/FINO median {dens} kg/m3 is not a gas density - FINO is not m3; year left out")
            if not have.empty:
                frames.append(have)
            continue
        # a month in the file with no gas row for a destination means no exports there that month
        idx = pd.MultiIndex.from_product([[pd.Timestamp(year, m, 1) for m in months],
                                          sorted(set(res["Country"]) | {"Brasil", "Argentina"})],
                                         names=["Month", "Country"])
        res = res.reset_index().set_index(["Month", "Country"]).reindex(idx)
        res[["Volume_mcm", "Net_weight_t", "Value_USD_m"]] = res[["Volume_mcm", "Net_weight_t",
                                                                  "Value_USD_m"]].fillna(0.0)
        res = res.reset_index().set_index("Month")
        out(f"  {year}: months {months}; density median {dens:.3f} kg/m3; volume by country (million m3): "
            f"{res.groupby('Country')['Volume_mcm'].sum().round(1).to_dict()}")
        frames.append(res)
    years_seen = set(links)
    if not old.empty:   # years no longer linked on the page stay as they were
        frames += [old[old.index.year == y] for y in sorted(set(old.index.year) - years_seen)]
    if not frames:
        raise SystemExit("no export data")
    ex = pd.concat(frames).sort_index()
    return ex[["Country", "Volume_mcm", "Net_weight_t", "Value_USD_m", "Density_kg_m3"]].round(4)


# ------------------------------------------------------------------ cross-check against importers' data

def cross_check(exp_mcm_month):
    """INE exports vs Brazil's ANP pipeline imports from Bolivia (mcm/d) and Argentina's ENARGAS imports from
    Bolivia (mcm/month), when those workbooks are in the repo."""
    days = exp_mcm_month.index.days_in_month
    cc = pd.DataFrame(index=exp_mcm_month.index)
    cc["INE_Exports_Brazil_mcm_per_day"] = (exp_mcm_month["Brasil"] / days).round(3)
    cc["INE_Exports_Argentina_mcm_per_day"] = (exp_mcm_month["Argentina"] / days).round(3)
    try:
        b = pd.read_excel(os.path.join(OUTPUT_DIR, "brazil_gas_monthly.xlsx"), sheet_name="Supply (ANP)")
        b = b.set_index(pd.to_datetime(b["date"]))["Bolivia_Pipeline"]
        cc["ANP_Brazil_imports_from_Bolivia_mcm_per_day"] = b.reindex(cc.index).round(3)
    except Exception as e:
        out(f"Brazil cross-check skipped: {type(e).__name__}: {e}")
    try:
        a = pd.read_excel(os.path.join(OUTPUT_DIR, "argentina_gas_monthly.xlsx"), sheet_name="Supply net")
        a = a.set_index(pd.to_datetime(a["date"]))["Imports_Bolivia"]
        cc["ENARGAS_Argentina_imports_from_Bolivia_mcm_per_day"] = (a.reindex(cc.index) / days).round(3)
    except Exception as e:
        out(f"Argentina cross-check skipped: {type(e).__name__}: {e}")
    for ine, imp, lab in (("INE_Exports_Brazil_mcm_per_day", "ANP_Brazil_imports_from_Bolivia_mcm_per_day", "Brazil"),
                          ("INE_Exports_Argentina_mcm_per_day",
                           "ENARGAS_Argentina_imports_from_Bolivia_mcm_per_day", "Argentina")):
        if imp in cc:
            cc[f"Diff_{lab}_mcm_per_day"] = (cc[ine] - cc[imp]).round(3)
            cc[f"Diff_{lab}_pct"] = (100 * (cc[ine] - cc[imp]) / cc[imp].where(cc[imp] > 0.05)).round(1)
    return cc


# ------------------------------------------------------------------ main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/bolivia_gas_demand_by_sector.xlsx")
    args = ap.parse_args()

    # 1. demand by sector (domestic market)
    monthly, title, unit, scale = ine_table(r"Red de Distribuci", FALLBACK_URL, _sector, "H05", 3)
    sectors = [c for c in ["Power", "Industrial", "Residential", "Commercial", "Vehicle_CNG"] if c in monthly]
    days = monthly.index.days_in_month
    perday = monthly.div(days, axis=0).round(3)
    perday.columns = [f"{c}_mcm_per_day" for c in perday.columns]
    table = pd.concat([monthly.round(2).add_suffix("_mcm_month"), perday], axis=1)
    check = (monthly[sectors].sum(axis=1) - monthly["Total"]).abs().max() if "Total" in monthly else None
    out(f"{len(monthly)} months {monthly.index.min():%Y-%m}..{monthly.index.max():%Y-%m}; "
        f"max |sum of sectors - total| = {check}")
    out(perday.tail(6).to_string())

    # 2. production by department
    prod, ptitle, punit, pscale = ine_table(r"Producci.n de Gas Natural por Departamento", PRODUCTION_FALLBACK_URL,
                                            _department, "H02", 3)
    deps = [c for c in ["Tarija", "Santa_Cruz", "Chuquisaca", "Cochabamba"] if c in prod]
    pcheck = (prod[deps].sum(axis=1) - prod["Total"]).abs().max()
    out(f"production: {len(prod)} months {prod.index.min():%Y-%m}..{prod.index.max():%Y-%m}; "
        f"max |sum of departments - total| = {pcheck}")

    # 3. exports by destination
    ex = update_exports(args.out)
    exp = ex.reset_index().pivot_table(index="Month", columns="Country", values="Volume_mcm", aggfunc="sum")
    for c in ("Brasil", "Argentina"):
        if c not in exp:
            exp[c] = 0.0
    other = [c for c in exp.columns if c not in ("Brasil", "Argentina")]
    exp = exp[exp.index.year >= DATA_START]

    # combined monthly table (million m3/month)
    idx = pd.date_range(f"{DATA_START}-01-01", max(prod.index.max(), exp.index.max(), monthly.index.max()),
                        freq="MS")
    pe = pd.DataFrame(index=idx)
    pe["Production"] = prod["Total"]
    for d in deps:
        pe[f"Production_{d}"] = prod[d]
    pe["Exports_Brazil"] = exp["Brasil"]
    pe["Exports_Argentina"] = exp["Argentina"]
    if other:
        pe["Exports_other"] = exp[other].sum(axis=1)
    pe["Domestic_market"] = monthly["Total"]
    flows = [c for c in ("Exports_Brazil", "Exports_Argentina", "Exports_other", "Domestic_market") if c in pe]
    pe["Total"] = pe[flows].sum(axis=1, min_count=len(flows)).where(pe[flows].notna().all(axis=1))
    pe["Production_less_total"] = pe["Production"] - pe["Total"]
    pe.index.name = "Month"
    d = pe.index.days_in_month
    pe_day = pe.div(d, axis=0).round(3).add_suffix("_mcm_per_day")
    pe_table = pd.concat([pe.round(2).add_suffix("_mcm_month"), pe_day], axis=1)
    out("\nProduction and exports, million m3/day (last 14 months):")
    out(pe.div(d, axis=0).round(2).tail(14).to_string())

    cc = cross_check(exp.reindex(idx))
    out("\nExport cross-check (million m3/day):")
    out(cc.dropna(how="all").to_string())
    for lab in ("Brazil", "Argentina"):
        col = f"Diff_{lab}_mcm_per_day"
        if col in cc:
            v = cc[col].dropna()
            if len(v):
                out(f"{lab}: {len(v)} months; mean diff {v.mean():.3f} mcm/d; mean |diff| {v.abs().mean():.3f} mcm/d; "
                    f"median diff % {cc[f'Diff_{lab}_pct'].median():.1f}")

    latest_prod = prod.index.max().strftime("%b %Y")
    latest_exp = exp.index.max().strftime("%b %Y")
    notes = [
        "UNITS",
        "*_mcm_month: millions of cubic metres per month, as published by INE. *_mcm_per_day: the same divided by "
        "days in the month.",
        f"Demand sheet header: {title}",
        f"Unit read from that header: {unit}" + ("" if scale == 1 else f" - converted to million m3 (x{scale})"),
        f"Production sheet header: {ptitle}",
        f"Unit read from that header: {punit}" + ("" if pscale == 1 else f" - converted to million m3 (x{pscale})"),
        "Exports: the FINO (quantity) column of INE's export micro-data, in m3 (standard conditions, 60F as for "
        "YPFB's export statistics). Checked every run: net weight / FINO is a natural-gas density "
        f"({DENSITY_RANGE[0]}-{DENSITY_RANGE[1]} kg/m3); the Density_kg_m3 column shows it.",
        "",
        "SHEETS",
        "Demand by sector: Power (generadoras electricas), Industrial, Residential (domestico), Commercial, "
        "Vehicle_CNG (GNV). Domestic market only.",
        "Production and exports: Production (total and by department: Tarija, Santa Cruz, Chuquisaca, Cochabamba); "
        "Exports_Brazil (GASBOL), Exports_Argentina (GJA / Yacuiba-Pocitos); Domestic_market (= Demand by sector "
        "Total); Total = exports + domestic market; Production_less_total = production minus that total (liquids "
        "extracted, field fuel and own use, losses, statistical difference - INE's production includes pipeline "
        "deliveries, liquids and fuel).",
        "INE publishes no field-level production split (San Alberto, Sabalo, Margarita-Huacaya, Incahuasi...).",
        f"{EXPORTS_SHEET}: monthly sums of the customs records for NANDINA {GAS_NANDINA} by destination: volume "
        "(million m3), net weight (tonnes), FOB value (US$ million) and the implied density.",
        "Export cross-check: INE exports against the importers' own data - Brazil: ANP pipeline imports from "
        "Bolivia (brazil_gas_monthly.xlsx, Supply (ANP), Bolivia_Pipeline, m3 at 20C); Argentina: ENARGAS imports "
        "from Bolivia (argentina_gas_monthly.xlsx, Supply net, Imports_Bolivia, m3 of 9300 kcal). Differences "
        "come from reference conditions (60F vs 20C is ~1.5%), heating-value normalisation, measurement point "
        "and timing, and preliminary customs months.",
        "",
        "COVERAGE",
        f"Monthly from {DATA_START}. Production and domestic demand to {latest_prod} (INE's hydrocarbons tables are "
        f"annual updates); exports to {latest_exp} (customs micro-data, updated monthly). Latest years are "
        "preliminary (p) and may be revised. Exports are incremental: final years already in the workbook are "
        "kept, preliminary and missing years are downloaded again.",
        "",
        "SOURCE",
        f"Instituto Nacional de Estadistica (INE), 'Volumen Comercializado de Gas Natural segun Red de Distribucion' "
        f"and 'Produccion de Gas Natural por Departamento' (INE / YPFB): {PAGE}",
        f"INE export micro-data ('Base de Datos Exportaciones', yearly workbooks): {EXPORTS_PAGE}",
    ]
    table.index = table.index.strftime("%Y-%m")
    table.index.name = "Month"
    pe_table.index = pe_table.index.strftime("%Y-%m")
    ex_out = ex.copy()
    ex_out.index = ex_out.index.strftime("%Y-%m")
    ex_out.index.name = "Month"
    cc_out = cc.dropna(how="all").copy()
    cc_out.index = cc_out.index.strftime("%Y-%m")
    cc_out.index.name = "Month"
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Demand by sector": table, "Production and exports": pe_table,
                                         EXPORTS_SHEET: ex_out, "Export cross-check": cc_out},
                              notes, {"UNITS", "SHEETS", "COVERAGE", "SOURCE"})
    chart_df = monthly[sectors].div(days, axis=0)
    chart_df.columns = [c.replace("_", " ") for c in sectors]
    xlsx_charts.add_chart_sheet(args.out, chart_df, "Bolivia gas demand by sector", "million m3/day",
                                kind="stacked_bar")
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
