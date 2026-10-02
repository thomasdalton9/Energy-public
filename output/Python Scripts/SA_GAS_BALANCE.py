"""
South America gas balance, monthly from 2021, million m3/day: domestic production, imports (pipeline / LNG) and
exports (pipeline / LNG) for each country, built ONLY from the country workbooks this repo already pulls - no new
downloads, no estimates. Apparent supply = production + imports - exports.

Country sources (each workbook's own official source; see its Units tab):
  Argentina  argentina_gas_monthly.xlsx  'National' produccion_gas_natural (SE, gross, million m3/month);
             'Supply net' LNG_Escobar, LNG_BahiaBlanca, Imports_Bolivia, Imports_Chile; 'Exports by destination'
  Bolivia    bolivia_gas_demand_by_sector.xlsx 'Production and exports' (INE)
  Brazil     brazil_gas_monthly.xlsx 'Supply (ANP)' Domestic_Available (production net of reinjection, flaring,
             E&P use and UPGN), Bolivia_Pipeline, Argentina_Pipeline, LNG_Implied
  Chile      chile_gas_imports.xlsx 'Domestic production' (ENAP+CEOP), 'Imports by use' (pipeline / LNG),
             'Gas imports' total for months after the split ends (counted as unsplit imports)
  Colombia   colombia_gas_demand_by_sector.xlsx 'Supply by source' Domestic_production, LNG_imports_SPEC,
             Venezuela_imports
  Ecuador    ecuador_gas.xlsx 'Gas by use' Amistad production, 'Total demand' LNG imports (MMBtu -> m3 at 1,030 Btu/cf)
  Peru       peru_gas_demand_by_sector.xlsx 'Production by lot' total, 'LNG exports'
  Trinidad   trinidad_gas.xlsx 'Production by company' TOTAL; LNG exports = 'Utilization by sector' LNG (feed gas
             to Atlantic LNG, MEEI)
  Uruguay    uruguay_gas_demand_by_sector.xlsx Imports (from Argentina, pipeline); no production
Production is each country's published basis (gross for Argentina, Bolivia, Peru, Trinidad, Colombia; Brazil
net 'available' production) - noted per country on the Units tab. Venezuela publishes no gas statistics.

Usage: python3 SA_GAS_BALANCE.py [--out PATH]
"""
import argparse
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import xlsx_notes  # noqa: E402

DIR = os.path.join("output", "Data and Chart Outputs")
OUT = os.path.join(DIR, "south_america_gas_balance.xlsx")
START = "2021-01-01"
MMBTU_PER_MCM = 36374.0       # 1,030 Btu/cf (as add_charts.py)
MCM_PER_MMSCF = 0.0283168
FIELDS = ["Production", "Imports_pipeline", "Imports_LNG", "Imports_unsplit", "Exports_pipeline", "Exports_LNG"]


def read(name, sheet, date_col):
    path = os.path.join(DIR, name)
    df = pd.read_excel(path, sheet_name=sheet)
    df = df.drop(columns=[c for c in df.columns if str(c).startswith("Unnamed")])
    df[date_col] = pd.to_datetime(df[date_col].astype(str), errors="coerce")
    return df.dropna(subset=[date_col]).set_index(date_col).sort_index()


def per_day(month_total):
    return month_total / month_total.index.days_in_month


def country_frames():
    out, notes = {}, {}

    def put(country, **cols):
        df = pd.DataFrame({k: v for k, v in cols.items() if v is not None})
        df.index = pd.DatetimeIndex(df.index).to_period("M").to_timestamp()
        out[country] = df.groupby(level=0).sum(min_count=1)

    def safe(country, fn, note):
        try:
            fn()
            notes[country] = note
        except Exception as e:  # noqa: BLE001 - one country's workbook changing must not stop the others
            print(f"  {country}: skipped ({type(e).__name__}: {e})", flush=True)

    def argentina():
        n = read("argentina_gas_monthly.xlsx", "National", "date")
        s = read("argentina_gas_monthly.xlsx", "Supply net", "date")
        e = read("argentina_gas_monthly.xlsx", "Exports by destination", "date")
        put("Argentina", Production=per_day(n["produccion_gas_natural"]),
            Imports_pipeline=per_day(s[["Imports_Bolivia", "Imports_Chile"]].sum(axis=1, min_count=1)),
            Imports_LNG=per_day(s[["LNG_Escobar", "LNG_BahiaBlanca"]].sum(axis=1, min_count=1)),
            Exports_pipeline=per_day(e["Total_exports"]))

    def bolivia():
        b = read("bolivia_gas_demand_by_sector.xlsx", "Production and exports", "Month")
        put("Bolivia", Production=b["Production_mcm_per_day"],
            Exports_pipeline=b[["Exports_Brazil_mcm_per_day", "Exports_Argentina_mcm_per_day"]].sum(axis=1, min_count=1))

    def brazil():
        s = read("brazil_gas_monthly.xlsx", "Supply (ANP)", "date")
        put("Brazil", Production=s["Domestic_Available"],
            Imports_pipeline=s[["Bolivia_Pipeline", "Argentina_Pipeline"]].sum(axis=1, min_count=1),
            Imports_LNG=s["LNG_Implied"])

    def chile():
        p = read("chile_gas_imports.xlsx", "Domestic production", "Month")
        u = read("chile_gas_imports.xlsx", "Imports by use", "Month")
        t = read("chile_gas_imports.xlsx", "Gas imports", "Month")
        pipe = u[[c for c in u.columns if c.startswith("Pipeline_") and c.endswith("_mcm_per_day")]].sum(axis=1, min_count=1)
        lng = u[[c for c in u.columns if c.startswith("LNG_") and c.endswith("_mcm_per_day")]].sum(axis=1, min_count=1)
        unsplit = t["Imports_mcm_per_day_approx"][t.index > u.index.max()] if len(u) else t["Imports_mcm_per_day_approx"]
        put("Chile", Production=p["Total_mcm_per_day"], Imports_pipeline=pipe, Imports_LNG=lng, Imports_unsplit=unsplit)

    def colombia():
        s = read("colombia_gas_demand_by_sector.xlsx", "Supply by source", "Month")
        put("Colombia", Production=s["Domestic_production_mcmd"], Imports_LNG=s["LNG_imports_SPEC_mcmd"],
            Imports_pipeline=s["Venezuela_imports_mcmd"])

    def ecuador():
        g = read("ecuador_gas.xlsx", "Gas by use", "Month")
        t = read("ecuador_gas.xlsx", "Total demand", "Month")
        imp = t[["Imports_power_MMBtu_per_day", "Imports_industry_MMBtu_per_day"]].sum(axis=1, min_count=1) / MMBTU_PER_MCM
        put("Ecuador", Production=g["Amistad_production_MMBtu_per_day"] / MMBTU_PER_MCM, Imports_LNG=imp)

    def peru():
        p = read("peru_gas_demand_by_sector.xlsx", "Production by lot", "Month")
        x = read("peru_gas_demand_by_sector.xlsx", "LNG exports", "Month")
        put("Peru", Production=p["Total_mcm_per_day"], Exports_LNG=x["LNG_exports_mcm_per_day"])

    def trinidad():
        p = read("trinidad_gas.xlsx", "Production by company", "date")
        u = read("trinidad_gas.xlsx", "Utilization by sector", "date")
        put("Trinidad & Tobago", Production=p.loc[p["company"] == "TOTAL", "mmscfd"] * MCM_PER_MMSCF,
            Exports_LNG=u.loc[u["sector"] == "LNG", "mmscfd"] * MCM_PER_MMSCF)

    def uruguay():
        u = read("uruguay_gas_demand_by_sector.xlsx", "Demand by sector", "Month")
        put("Uruguay", Imports_pipeline=u["Imports_mcm_per_day"])

    for c, fn, note in [
            ("Argentina", argentina, "gross production (SE); imports: LNG Escobar/Bahia Blanca, pipeline Bolivia/Chile; "
                                     "exports by destination (ENARGAS)"),
            ("Bolivia", bolivia, "production by department (INE H02); exports to Brazil/Argentina (INE customs)"),
            ("Brazil", brazil, "ANP 'available' domestic production (net of reinjection, flaring, E&P use, UPGN); "
                               "imports Bolivia/Argentina pipeline, LNG regasification (implied)"),
            ("Chile", chile, "ENAP+CEOP production (to the last CNE file); imports by use (pipeline/LNG), the "
                             "monthly report total after the split ends (unsplit)"),
            ("Colombia", colombia, "domestic production, LNG (SPEC) and Venezuela imports (supply by source)"),
            ("Ecuador", ecuador, "Amistad production and LNG imports, MMBtu converted at 1,030 Btu/cf"),
            ("Peru", peru, "production by lot (fiscalised); LNG exports (Peru LNG)"),
            ("Trinidad & Tobago", trinidad, "production (MEEI, all companies); LNG = feed gas to Atlantic LNG (MEEI)"),
            ("Uruguay", uruguay, "imports from Argentina (pipeline); no domestic production")]:
        safe(c, fn, note)
    return out, notes


def build():
    frames, notes = country_frames()
    rows = []
    for c, df in frames.items():
        df = df[df.index >= START].reindex(columns=FIELDS)
        for d, r in df.iterrows():
            rows.append({"date": d, "country": c, **r.to_dict()})
    long = pd.DataFrame(rows)
    long["Imports_total"] = long[["Imports_pipeline", "Imports_LNG", "Imports_unsplit"]].sum(axis=1, min_count=1)
    long["Exports_total"] = long[["Exports_pipeline", "Exports_LNG"]].sum(axis=1, min_count=1)
    long["Apparent_supply"] = long[["Production", "Imports_total"]].sum(axis=1, min_count=1) - long["Exports_total"].fillna(0)
    producers = set(long.loc[long["Production"].notna(), "country"])
    long.loc[long["country"].isin(producers) & long["Production"].isna(), "Apparent_supply"] = float("nan")
    num = long.select_dtypes("number").columns
    long[num] = long[num].round(3)
    long = long.sort_values(["date", "country"])

    def wide(col):
        w = long.pivot(index="date", columns="country", values=col)   # one row per country-month
        w = w.dropna(axis=1, how="all").dropna(axis=0, how="all")
        w.index.name = "date"
        return w.round(3)
    return long, {"Production": wide("Production"), "Imports": wide("Imports_total"), "Exports": wide("Exports_total"),
                  "Apparent supply": wide("Apparent_supply")}, notes


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    long, wides, notes = build()
    last = {k: f"{v.dropna(how='all').index.max():%b-%y}" for k, v in wides.items() if not v.empty}
    lines = ["SOUTH AMERICA GAS BALANCE (monthly, million m3/day)",
             "Built only from this repo's country workbooks (each from its official source; see their Units tabs) - no new "
             "downloads and no estimates. Monthly averages in million m3/day. Apparent supply = production + imports - "
             "exports (for a producing country, only months its production is published). A blank is 'not published yet'. "
             "Charts run to the last month every country that reported in the past year has published.",
             f"Latest month per sheet: {last}", "",
             "SHEETS",
             "Production / Imports / Exports / Apparent supply: one column per country. 'Balance (long)': every field "
             "per country and month (Imports_pipeline, Imports_LNG, Imports_unsplit, Exports_pipeline, Exports_LNG).",
             "Intra-regional trade appears twice (Bolivia's exports are Brazil's and Argentina's imports): sum "
             "production, not imports, for a regional total.", "",
             "COUNTRIES"] + [f"{c}: {n}" for c, n in notes.items()] + [
             "Venezuela: no published gas statistics - not included.", "",
             "UPDATES", "Rebuilt from the country workbooks on the 1st and 15th (south_america_gas_balance.yml), after "
             "their pulls."]
    sheets = {**wides, "Balance (long)": long.set_index("date")}
    xlsx_notes.write_workbook(args.out, sheets, lines, ["SHEETS", "COUNTRIES", "UPDATES"])
    print(f"saved {args.out}: {len(long)} country-months, countries {sorted(long['country'].unique())}, "
          f"latest {last}", flush=True)
    print(wides["Production"].tail(3).round(1).to_string())


if __name__ == "__main__":
    main()
