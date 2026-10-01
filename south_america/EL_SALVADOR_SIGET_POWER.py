"""
El Salvador power generation by technology, MONTHLY, from SIGET.

The grid operator UT (Unidad de Transacciones, ut.com.sv) publishes daily
data, but its name servers and web server do not answer outside El Salvador,
so it cannot be reached from GitHub Actions. The regulator SIGET publishes the
same national statistics in its public Power BI report 'Visualizador dinamico
de Estadisticas Electricas'
  https://www.siget.gob.sv/gerencias/electricidad/informe-de-mercado-y-estadisticas-electricas/estadisticas-electricas-bi/
  report: https://app.powerbi.com/view?r=eyJrIjoiYmM0MDJhMTktNmQ2YS00ZDkxLTgyYWUtMDUxYjQ5MjQxYWY5IiwidCI6IjE1OTg0YzZmLTNiMmMtNDZmMi1iNmU0LTUzNDQ3MGY2MDVmNyJ9
Its model table 'DATOS GENERACION PLANTA (3)' holds MWh per plant per month
by technology, split into gross generation, own use and net generation. This
script asks the report's anonymous query endpoint (the same call the browser
makes) for MWh by year, month, technology, generation type and generator type.

Finest granularity is MONTHLY, so sheet "Daily" holds one row per month dated
the first of the month with that month's NET generation (MWh); the sheet name
"Daily" is kept for the standard layout (power_daily_std.py). SIGET refreshes
the report about once a year (it held Jan-2023 to Dec-2025 in Oct-2026), so
saved months are kept and new ones merged in (the report could drop old years).

Mapping (SIGET 'TECNOLOGIA' -> column):
  HIDRAULICA -> Hydro            GNL -> Gas (LNG, Energia del Pacifico)
  EOLICO -> Wind                 FOTOVOLTAICA -> Solar
  BIOMASA, BIOGAS -> Bioenergy   OTROS COMBUSTIBLES FOSILES -> Oil (bunker/diesel)
  GEOTERMICA -> Other            DISTRIBUIDORAS EN MM, unlabelled ids -> Other
  TRANSACCIONES REGIONALES -> excluded (imports/exports)

Usage: python3 EL_SALVADOR_SIGET_POWER.py [--out PATH]
"""

print("STARTING", flush=True)

import argparse
import base64
import json
import os
import sys
import unicodedata

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import power_daily_std as std  # noqa: E402

OUT = "output/Data and Chart Outputs/el_salvador_power_generation_daily.xlsx"
TOKEN = "eyJrIjoiYmM0MDJhMTktNmQ2YS00ZDkxLTgyYWUtMDUxYjQ5MjQxYWY5IiwidCI6IjE1OTg0YzZmLTNiMmMtNDZmMi1iNmU0LTUzNDQ3MGY2MDVmNyJ9"
REPORT_URL = "https://app.powerbi.com/view?r=" + TOKEN
PAGE_URL = "https://www.siget.gob.sv/gerencias/electricidad/informe-de-mercado-y-estadisticas-electricas/estadisticas-electricas-bi/"
CLUSTER = "https://wabi-paas-1-scus-api.analysis.windows.net"
FACT = "DATOS GENERACIÓN PLANTA (3)"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept": "application/json", "Content-Type": "application/json"}
TECH_MAP = {"HIDRAULICA": "Hydro", "GNL": "Gas", "EOLICO": "Wind", "FOTOVOLTAICA": "Solar", "BIOMASA": "Bioenergy",
            "BIOGAS": "Bioenergy", "OTROS COMBUSTIBLES FOSILES": "Oil", "GEOTERMICA": "Other",
            "DISTRIBUIDORAS EN MM": "Other", "TRANSACCIONES REGIONALES": None}
COLS = [f"{f}_MWh" for f in ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]]


def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return " ".join(s.upper().split())


def decode(resp):
    """Power BI DSR (compressed rows) -> list of rows."""
    ds = resp["results"][0]["result"]["data"]["dsr"]["DS"][0]
    if ds.get("RT"):
        raise RuntimeError("Power BI returned a partial window; raise the row count")
    vd = ds.get("ValueDicts", {})
    rows = ds["PH"][0]["DM0"]
    if not rows:
        return []
    schema = rows[0]["S"]
    out, prev = [], [None] * len(schema)
    for r in rows:
        rep, nul = r.get("R", 0), r.get("Ø", 0)
        c = iter(r.get("C", []))
        cur = []
        for i, s in enumerate(schema):
            if rep >> i & 1:
                v = prev[i]
            elif nul >> i & 1:
                v = None
            else:
                v = next(c)
                if "DN" in s and isinstance(v, int):
                    v = vd[s["DN"]][v]
            cur.append(v)
        out.append(cur)
        prev = cur
    return out


class PowerBI:
    def __init__(self):
        tok = json.loads(base64.b64decode(TOKEN + "=" * (-len(TOKEN) % 4)))
        self.h = dict(H, **{"X-PowerBI-ResourceKey": tok["k"]})
        r = requests.get(f"{CLUSTER}/public/reports/{tok['k']}/modelsAndExploration?preferReadOnlySession=true",
                         headers=self.h, timeout=60)
        r.raise_for_status()
        j = r.json()
        self.model_id, self.db = j["models"][0]["id"], j["models"][0]["dbName"]
        self.report = j["exploration"]["report"]["objectId"]

    def query(self, entity, cols, sum_col=None, count=30000):
        sel = [{"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": c}, "Name": f"t.{c}"} for c in cols]
        if sum_col:
            sel.append({"Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Source": "t"}},
                                                                  "Property": sum_col}}, "Function": 0},
                        "Name": f"Sum(t.{sum_col})"})
        body = {"version": "1.0.0", "cancelQueries": [], "modelId": self.model_id,
                "queries": [{"Query": {"Commands": [{"SemanticQueryDataShapeCommand": {
                    "Query": {"Version": 2, "From": [{"Name": "t", "Entity": entity, "Type": 0}], "Select": sel},
                    "Binding": {"Primary": {"Groupings": [{"Projections": list(range(len(sel)))}]},
                                "DataReduction": {"DataVolume": 4, "Primary": {"Window": {"Count": count}}},
                                "Version": 1}}}]},
                    "QueryId": "", "ApplicationContext": {"DatasetId": self.db,
                                                          "Sources": [{"ReportId": self.report, "VisualId": ""}]}}]}
        r = requests.post(f"{CLUSTER}/public/reports/querydata?synchronous=true", headers=self.h,
                          data=json.dumps(body), timeout=120)
        r.raise_for_status()
        return decode(r.json())

    def lookup(self, entity, id_col, name_col):
        return {row[0]: row[1] for row in self.query(entity, [id_col, name_col])}


def fetch():
    """Long frame: month, technology, generation type, generator type, MWh."""
    pbi = PowerBI()
    years = pbi.lookup("AÑO (2)", "ID_AÑO", "N_AÑO")
    techs = pbi.lookup("TECNOLOGÍA (3)", "ID_TECNOLOGÍA", "N_TECNOLOGÍA")
    gens = pbi.lookup("GENERACIÓN (2)", "ID_GENERACIÓN", "N_GENERACIÓN")
    tipos = pbi.lookup("TIPO DE GENERADOR (2)", "ID_TIPO", "N_TIPO GENERADOR")
    rows = pbi.query(FACT, ["ID_AÑO", "ID_MES", "ID_TECNOLOGÍA", "ID_GENERACIÓN", "ID_TIPO"], "MWh")
    recs = []
    for y, m, t, g, k, v in rows:
        if y not in years or v in (None, ""):
            continue
        recs.append({"month": pd.Timestamp(int(years[y]), int(m), 1),
                     "technology": norm(techs.get(t, f"ID {t}")), "generation": norm(gens.get(g, f"ID {g}")),
                     "generator_type": str(tipos.get(k, f"ID {k}")).strip(), "MWh": float(v)})
    print(f"SIGET Power BI: {len(rows)} rows -> {len(recs)} records", flush=True)
    return pd.DataFrame(recs)


def to_monthly(long):
    """Net generation by standard column, complete months only."""
    net = long[long.generation == "GENERACION NETA"].copy()
    unknown = sorted(set(net.technology) - set(TECH_MAP))
    if unknown:
        print(f"WARNING: unmapped SIGET technologies {unknown} -> Other", flush=True)
    net["col"] = net.technology.map(lambda t: TECH_MAP.get(t, "Other"))
    net = net[net.col.notna()]
    wide = net.pivot_table(index="month", columns="col", values="MWh", aggfunc="sum")
    # a month counts only when both renewables and thermal are in (2022 holds a few renewable plants only)
    thermal = wide.reindex(columns=["Gas", "Oil"]).fillna(0).sum(axis=1)
    keep = thermal > 0
    dropped = wide.index[~keep]
    if len(dropped):
        print(f"Skipping {len(dropped)} incomplete months (no thermal rows): {[f'{d:%Y-%m}' for d in dropped]}", flush=True)
    out = std.standardise(wide[keep])
    out.index.name = "date"
    return out


NOTES = [
    "UNITS",
    "Daily: MWh per MONTH of NET generation (gross minus plant own use), one row per month dated the 1st of the "
    "month (sheet name 'Daily' kept for the standard raw-power layout). Total_MWh = sum of the fuel columns.",
    "By_technology_MWh: SIGET's own technology names, MWh per month, split into GENERACION BRUTA (gross), CONSUMO "
    "PROPIO (own use) and GENERACION NETA (net), and by generator type (wholesale-market generators, distributed "
    "renewable GDR, distributed thermal GDT, distributors in the wholesale market).",
    "",
    "COVERAGE",
    "Monthly, from the first month SIGET's report holds with both renewable and thermal plants (Jan-2023 as of "
    "Oct-2026) to the latest month it holds (Dec-2025 as of Oct-2026). Data from 2021-2022 is not in the report: "
    "SIGET's annual 'Boletin de Estadisticas Electricas' PDFs cover those years only as annual totals. SIGET "
    "refreshes the report about once a year; months already saved are kept if the report later drops them.",
    "UT's daily data would be finer, but UT's servers (ut.com.sv) do not answer outside El Salvador.",
    "",
    "SOURCE",
    "SIGET (Superintendencia General de Electricidad y Telecomunicaciones), 'Visualizador dinamico de Estadisticas "
    f"Electricas' (public Power BI report): {PAGE_URL} ; report {REPORT_URL} . Model table 'DATOS GENERACION "
    "PLANTA (3)', read through the report's public query endpoint. Electricity statistics moved from SIGET to "
    "DGEHM on 17-Jul-2026; the report is still SIGET's.",
    "Updated weekly by GitHub Actions (el_salvador_power_generation_daily.yml).",
    "",
    "MAPPING",
    "HIDRAULICA -> Hydro; GNL (LNG, Energia del Pacifico) -> Gas; EOLICO -> Wind; FOTOVOLTAICA -> Solar; "
    "BIOMASA (sugar-cane bagasse) and BIOGAS -> Bioenergy; OTROS COMBUSTIBLES FOSILES (bunker / diesel engines) -> "
    "Oil; GEOTERMICA (geothermal) -> Other; DISTRIBUIDORAS EN MM and unlabelled technology ids (a few GWh a year) "
    "-> Other; TRANSACCIONES REGIONALES (regional market imports/exports) excluded. El Salvador has no coal or "
    "nuclear plant.",
]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()

    print("MAPPING (SIGET technology -> column):", flush=True)
    for k, v in TECH_MAP.items():
        print(f"  {k} -> {v or 'excluded'}", flush=True)

    long = fetch()
    if long.empty:
        print("SIGET report returned no generation rows.", flush=True)
        sys.exit(1)
    monthly = to_monthly(long)

    saved = std.load_sheet(args.out, "Daily")
    daily = std.merge(monthly, saved)
    daily = daily.reindex(columns=[c for c in COLS if c in daily.columns] + ["Total_MWh"])
    detail = long.pivot_table(index="month", columns=["generation", "generator_type", "technology"], values="MWh",
                              aggfunc="sum").sort_index()
    detail.columns = [" | ".join(c) for c in detail.columns]
    saved_detail = std.load_sheet(args.out, "By_technology_MWh")
    if not saved_detail.empty:
        detail = std.merge(detail, saved_detail)
    detail.index.name = "month"
    std.write(args.out, daily, NOTES, {"By_technology_MWh": detail.round(1)})

    m = daily[daily.index >= daily.index.max() - pd.DateOffset(months=11)] / 1000
    print("Latest 12 months, GWh (net):", flush=True)
    print(m.round(1).to_string(), flush=True)
    y = (daily / 1000).groupby(daily.index.year).sum().round(0)
    print("Annual GWh (net):", flush=True)
    print(y.to_string(), flush=True)
    gross = long[long.generation == "GENERACION BRUTA"].groupby(long.month.dt.year).MWh.sum() / 1000
    print("Annual gross GWh (all technologies incl. imports rows):", gross.round(0).to_dict(), flush=True)


if __name__ == "__main__":
    main()
