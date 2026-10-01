"""
Bolivia installed generation capacity by technology from CNDC (Comite
Nacional de Despacho de Carga, the SIN operator) - free, public, no key.

Source: the WordPress REST API behind cndc.bo's dashboards (found via
discovery_archive/south_america/CAPACITY_PBUE_DISCOVERY3-5.py):
  /wp-json/cndc/v1/historico/potencia?desde=Y1&hasta=Y2
      capacity (MW) by technology per year - Hidroelectrica, Eolica, Solar,
      Biomasa, Termoelectrica - at December of each year (the dashboard
      labels the latest values 'Diciembre-<year>')
  /wp-json/cndc/v1/historico/potencia/detalle?anio=Y
      the same capacity per plant (central, empresa, MW) within each technology
  /wp-json/cndc/v1/estadisticas/documentos?categoria_id=235
      CNDC's annual 'Consumo de Diesel' workbooks (consdiesel_<year>.xlsx):
      the plants that burn diesel, used to split thermal capacity into
      diesel (Oil) and natural gas (Gas)
CNDC's monthly dashboard series (dashboard/potencia-historial) only repeats
each year's value for every month of that year (even future months), so the
real granularity is annual: one row per year, dated 1 January, holding the
capacity at December of that year. 2021 onwards.

Incremental: years already in the workbook are kept; only new years and
the latest year (CNDC may revise it) are re-fetched.

Usage: python3 BOLIVIA_CNDC_CAPACITY.py [--out PATH] [--test]
"""

print("STARTING", flush=True)

import argparse
import io
import os
import sys
import unicodedata
from datetime import date

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import power_capacity_common as pcc  # noqa: E402

API = "https://www.cndc.bo/wp-json/cndc/v1/"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36"}
TIMEOUT = (15, 60)
DEFAULT_OUT = os.path.join("output", "Data and Chart Outputs", "bolivia_power_capacity.xlsx")
DIESEL_CATEGORY = 235  # 'Consumo de Diesel' (annual statistics)

TECH = {"HIDROELECTRICA": "Hydro", "EOLICA": "Wind", "SOLAR": "Solar", "BIOMASA": "Bioenergy",
        "TERMOELECTRICA": "Thermal"}
TEC_CODE = {"hidro": "Hydro", "eolica": "Wind", "solar": "Solar", "biomasa": "Bioenergy", "termo": "Thermal"}


def key(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper()
    return " ".join(s.replace("(*)", "").split())


def get_json(path, **params):
    r = requests.get(API + path, params=params, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    return r.json()


def annual_by_technology(first, last):
    j = get_json("historico/potencia", desde=first, hasta=last)
    out = {}
    for s in j["series"]:
        tech = TECH.get(key(s["name"]))
        if tech is None:
            raise SystemExit(f"Unknown CNDC technology {s['name']!r}")
        out[tech] = pd.Series(s["data"], index=[int(y) for y in j["categorias"]], dtype=float)
    return pd.DataFrame(out)


def plants(year):
    j = get_json("historico/potencia/detalle", anio=year)
    rows = []
    for g in j.get("grupos", []):
        for p in g.get("plantas", []):
            rows.append({"year": year, "technology": TEC_CODE.get(g.get("tec"), g.get("name")),
                         "plant": p.get("central"), "company": p.get("empresa"), "mw": p.get("mw")})
    return pd.DataFrame(rows)


def diesel_plants():
    """{year: set of plant keys} from CNDC's annual 'Consumo de Diesel' workbooks."""
    j = get_json("estadisticas/documentos", categoria_id=DIESEL_CATEGORY, agrupado="true")
    out = {}
    for g in j.get("grupos", []):
        for d in g.get("docs", []):
            url = d.get("archivo_url")
            if not url:
                continue
            try:
                raw = pd.read_excel(io.BytesIO(requests.get(url, headers=HEADERS, timeout=TIMEOUT).content),
                                    header=None)
            except Exception as e:  # noqa: BLE001
                print(f"  diesel workbook {url}: {type(e).__name__}", flush=True)
                continue
            # Layout: a row with 'Ano' and the year, then a row with one plant name per column.
            hit = [i for i in range(len(raw)) if any(key(v) == "ANO" for v in raw.iloc[i])]
            if not hit:
                continue
            names = [v for v in raw.iloc[hit[0] + 1] if isinstance(v, str) and v.strip()]
            out[int(g.get("año") or d.get("año"))] = {key(n) for n in names}
    return out


def split_thermal(annual, detail, diesel):
    """Gas = thermal - diesel plants' capacity; Oil = diesel plants (matched by name to the plant detail)."""
    oil, matched = {}, {}
    for year in annual.index:
        names = diesel.get(year) or diesel.get(max([y for y in diesel if y <= year], default=None)) or set()
        th = detail[(detail["year"] == year) & (detail["technology"] == "Thermal")]
        hit = th[th["plant"].map(lambda p: any(key(p) == n or key(p).startswith(n + " ") or n.startswith(key(p))
                                               for n in names))]
        oil[year] = float(hit["mw"].sum())
        matched[year] = sorted(hit["plant"])
    oil = pd.Series(oil)
    out = annual.copy()
    out["Oil"] = oil
    out["Gas"] = out["Thermal"] - out["Oil"]
    return out.drop(columns="Thermal"), matched


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--test", action="store_true")
    args = ap.parse_args()

    this_year = date.today().year
    annual = annual_by_technology(pcc.START.year, this_year)
    print("CNDC capacity by technology (MW, December of each year):\n" + annual.to_string(), flush=True)

    old = pcc.load_monthly(args.out)
    old_plants = pcc.load_sheet(args.out, "CNDC by plant")
    have = set(old.index.year) if not old.empty else set()
    latest = int(annual.index.max())
    todo = [y for y in annual.index if y not in have or y == latest]
    print(f"Archive years {sorted(have)}; fetching plant detail for {todo}", flush=True)

    detail = pd.concat([plants(y) for y in todo], ignore_index=True)
    if not old_plants.empty:
        old_plants = old_plants.reset_index()
        detail = pd.concat([old_plants[~old_plants["year"].isin(todo)], detail], ignore_index=True)
    detail = detail.sort_values(["year", "technology", "mw"], ascending=[True, True, False])
    chk = detail.groupby(["year", "technology"])["mw"].sum().unstack()
    print("Plant detail sums (MW):\n" + chk.round(2).to_string(), flush=True)

    diesel = diesel_plants()
    print(f"Diesel-burning plants per CNDC 'Consumo de Diesel': { {y: sorted(v) for y, v in diesel.items()} }",
          flush=True)
    fuels, matched = split_thermal(annual, detail, diesel)
    print(f"Thermal plants counted as Oil (diesel): {matched}", flush=True)

    rows = fuels.copy()
    rows.index = [pd.Timestamp(year=int(y), month=1, day=1) for y in rows.index]
    rows = pcc.standardise(rows.rename(columns=lambda c: f"{c}_MW"))
    if not old.empty:
        old = pcc.standardise(old)
        rows = pd.concat([old[~old.index.isin(rows.index)], rows]).sort_index()
    print(rows.to_string(), flush=True)

    val = pcc.validation_lines(rows, "Bolivia")
    print("\n".join(val), flush=True)
    oil_names = sorted({p for v in matched.values() for p in v})
    notes = pcc.unit_notes("year") + [
        "Annual series: each row is dated 1 January and holds CNDC's capacity at December of that year (CNDC's "
        "monthly dashboard series only repeats the year's value for each month, so it adds nothing finer).",
        "",
        "COVERAGE",
        f"{rows.index.min():%Y} to {rows.index.max():%Y} ({len(rows)} years), from 2021. Bolivia's National "
        "Interconnected System (SIN) as reported by CNDC; isolated systems outside the SIN are not included.",
        "",
        "SOURCE",
        "CNDC (Comite Nacional de Despacho de Carga) - cndc.bo dashboard API: "
        f"{API}historico/potencia?desde=&hasta= (MW by technology) and {API}historico/potencia/detalle?anio= "
        f"(MW per plant); diesel plants from CNDC's annual 'Consumo de Diesel' workbooks "
        f"({API}estadisticas/documentos?categoria_id={DIESEL_CATEGORY}).",
        "Script: south_america/BOLIVIA_CNDC_CAPACITY.py (scheduled by .github/workflows/bolivia_power_capacity.yml).",
        "",
        "MAPPING",
        "Hydro_MW = Hidroelectrica; Wind_MW = Eolica; Solar_MW = Solar; Bioenergy_MW = Biomasa (sugar-mill bagasse "
        "plants).",
        f"Oil_MW = thermal plants that appear in CNDC's diesel-consumption statistics for that year ({', '.join(oil_names) or 'none'}).",
        "Gas_MW = Termoelectrica minus Oil_MW (natural-gas turbines and combined cycles).",
        "No coal, nuclear or geothermal; Other_MW = 0.",
        "",
        "VALIDATION",
    ] + val
    plant_sheet = detail.set_index("year")
    pcc.write(args.out, rows, notes, {"CNDC by technology": annual, "CNDC by plant": plant_sheet})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
