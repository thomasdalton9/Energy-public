"""
Ecuador generation by source from CENACE (Operador Nacional de
Electricidad) - free, public, no key.

Source: CENACE's "Informacion Operativa" page
(www.cenace.gob.ec/info-operativa/InformacionOperativa.htm). Its
"INFORMACION OPERATIVA DIARIA" tab covers the last complete day:
  - energy by source (MWh): total, hydro, thermal, non-conventional
    renewables, imports, exports
  - thermal split into 'Termica' (fuel oil / diesel) and 'Gas Natural'
  - the big hydro plants (Coca Codo Sinclair, Paute, Sopladora, Mazar ...)
  - a half-hourly generation curve by source (MW)
The page only ever shows that one day - there's no archive - so each run
adds the day to ec_generation_daily.csv and the history builds from the
first run (same as Chile). Run it daily (MASTER_SOUTH_AMERICA.py does).

CENACE doesn't publish reservoir levels here (Mazar / Amaluza); Mazar's
own output is in the plant table.

CENACE's server doesn't send its intermediate TLS certificate, which
Python can't work around the way a browser does, so certificate checks
are skipped for cenace.gob.ec only (public, read-only page).

Outputs:
  ec_generation_daily.csv       one row per day (MWh by source and plant)
  ec_generation_halfhourly.csv  the half-hourly curves (MW)
  ecuador_generation.xlsx       'Daily' (MWh and mean GW), 'Plants', 'Plants by region'
                                 (long format: date, plant, province, region,
                                 country, mwh - see PLANT_LOCATION), 'Half-hourly'

CENACE reports no region/province field itself - PLANT_LOCATION below is
this script's own mapping from each named plant to where it actually is,
looked up by hand (province, then Ecuador's standard 3-zone grouping:
Costa/Sierra/Oriente). "Otras Hidro" is CENACE's own catch-all for every
smaller hydro plant not named individually, so it can't be assigned a
single location - left as region "Nacional (agregado)".

Usage: python3 ECUADOR_CENACE.py
"""

print("STARTING", flush=True)

import base64
import json
import os
import re
import sys
from datetime import date

import numpy as np
import pandas as pd
import requests
import urllib3

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

URL = "https://www.cenace.gob.ec/info-operativa/InformacionOperativa.htm"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36"}
DAILY_CSV = "ec_generation_daily.csv"
HALFHOUR_CSV = "ec_generation_halfhourly.csv"
OUT_FILE = "ecuador_generation.xlsx"
MONTHS = {m: i for i, m in enumerate(["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
                                      "septiembre", "octubre", "noviembre", "diciembre"], 1)}
TOTALS = [("PRODUCCIÓN TOTAL", "total_mwh"), ("EXPORTACIÓN", "exports_mwh"), ("IMPORTACIÓN", "imports_mwh"),
          ("HIDRÁULICA", "hydro_mwh"), ("TÉRMICA", "thermal_mwh"), ("R. NO CONVENCIONAL", "renewable_mwh")]

# CENACE names each large plant but reports no location - mapped here by
# hand (province, then Ecuador's standard Costa/Sierra/Oriente zoning).
# Sopladora/Mazar are the two reservoirs of the same Paute Integral
# complex as Paute (Molino); Agoyán/San Francisco are the two steps of
# the same Pastaza-river cascade. "Otras Hidro" is CENACE's catch-all
# for every smaller hydro plant not named individually - no single
# location applies, so it's tagged as a national aggregate rather than
# guessed into one region.
PLANT_LOCATION = {
    "Coca Codo": ("Napo", "Oriente"),
    "Paute": ("Azuay", "Sierra"),
    "Sopladora": ("Azuay", "Sierra"),
    "Mazar": ("Azuay", "Sierra"),
    "Delsitanisagua": ("Zamora Chinchipe", "Oriente"),
    "San Francisco": ("Tungurahua", "Sierra"),
    "Agoyán": ("Tungurahua", "Sierra"),
    "Minas San Francisco": ("Azuay", "Sierra"),
    "Otras Hidro": (None, "Nacional (agregado)"),
}


def plain(html):
    return re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", html, flags=re.S))


def number(s):
    return float(s.replace(" ", "").replace("\xa0", "").replace(" ", "").replace(",", "."))


def decode(values):
    """Plotly arrays are either plain lists or {'dtype', 'bdata'} (base64)."""
    if isinstance(values, dict) and "bdata" in values:
        return np.frombuffer(base64.b64decode(values["bdata"]), dtype=np.dtype(values["dtype"])).tolist()
    return list(values or [])


def plots(html, start, end):
    """Plotly figures between two positions in the page: [(trace name, type, x, y)]."""
    out = []
    for m in re.finditer(r'Plotly\.newPlot\(\s*["\'][^"\']+["\']\s*,\s*(\[.*?\])\s*,\s*\{', html[start:end], re.S):
        try:
            traces = json.loads(m.group(1))
        except ValueError:
            continue
        out.append([(str(t.get("name") or "").strip(), t.get("type"), decode(t.get("x") or t.get("labels")),
                     decode(t.get("y") or t.get("values"))) for t in traces])
    return out


def parse(html):
    """The last complete day: (day, totals dict, plants dict, half-hourly frame)."""
    starts = [m.start() for m in re.finditer("INFORMACIÓN OPERATIVA DIARIA", html)]
    monthly = html.find("INFORMACIÓN OPERATIVA MENSUAL")
    if not starts or monthly < 0:
        raise RuntimeError("page layout changed - 'INFORMACIÓN OPERATIVA DIARIA' section not found")
    start = starts[-1]
    text = plain(html[start:monthly])
    d = re.search(r"(\d{1,2}) de ([a-záéíóú]+) de (\d{4})", text, re.I)
    day = date(int(d.group(3)), MONTHS[d.group(2).lower()], int(d.group(1)))
    totals = {}
    for label, key in TOTALS:
        m = re.search(re.escape(label) + r"\s+([\d\s., \xa0]+?)(?=\s+[A-ZÁÉÍÓÚ(]|$)", text)
        if m:
            totals[key] = number(m.group(1))
    plants, curves = {}, None
    for fig in plots(html, start, monthly):
        kinds = {t[1] for t in fig}
        if kinds == {"bar"}:
            for name, _, _, y in fig:
                if y:
                    plants[name] = float(y[0])
        elif "scatter" in kinds:
            curves = pd.DataFrame({name: y for name, _, _, y in fig}, index=fig[0][2])
    # thermal split, from the bar chart that holds 'Gas Natural'
    if "Gas Natural" in plants:
        totals["thermal_gas_mwh"] = plants.pop("Gas Natural")
        totals["thermal_oil_mwh"] = plants.pop("Térmica", totals.get("thermal_mwh", 0) - totals["thermal_gas_mwh"])
        plants.pop("Renovable", None)
    if curves is not None:
        curves.index = pd.to_datetime([f"{day} {x}" for x in curves.index])
        curves.index.name = "time"
    return day, totals, plants, curves


def upsert(path, frame):
    if os.path.exists(path):
        old = pd.read_csv(path, index_col=0, parse_dates=True)
        frame = frame.combine_first(old)
    frame = frame.sort_index()
    frame.to_csv(path)
    return frame


def main():
    r = requests.get(URL, headers=HEADERS, timeout=90, verify=False)
    r.raise_for_status()
    day, totals, plants, curves = parse(r.text)
    print(f"CENACE last complete day: {day}  total {totals.get('total_mwh', float('nan')):,.0f} MWh "
          f"(hydro {totals.get('hydro_mwh', 0):,.0f}, thermal {totals.get('thermal_mwh', 0):,.0f})", flush=True)

    row = {**totals, **{f"plant: {k}": v for k, v in plants.items()}}
    daily = upsert(DAILY_CSV, pd.DataFrame([row], index=pd.DatetimeIndex([pd.Timestamp(day)], name="date")))
    if curves is not None:
        upsert(HALFHOUR_CSV, curves)
    half = pd.read_csv(HALFHOUR_CSV, index_col=0, parse_dates=True) if os.path.exists(HALFHOUR_CSV) else pd.DataFrame()

    sources = daily[[c for c in ["total_mwh", "hydro_mwh", "thermal_mwh", "thermal_oil_mwh", "thermal_gas_mwh",
                                 "renewable_mwh", "imports_mwh", "exports_mwh"] if c in daily.columns]]
    gw = (sources / 24 / 1000).round(3)
    gw.columns = [c.replace("_mwh", "_gw") for c in gw.columns]
    plant_cols = [c for c in daily.columns if c.startswith("plant: ")]
    plants_wide = daily[plant_cols].rename(columns=lambda c: c[7:])

    by_region_rows = []
    for plant in plants_wide.columns:
        province, region = PLANT_LOCATION.get(plant, (None, "Unmapped"))
        for dt, mwh in plants_wide[plant].items():
            if pd.notna(mwh):
                by_region_rows.append({"date": dt, "plant": plant, "province": province, "region": region,
                                        "country": "Ecuador", "mwh": mwh})
    by_region = pd.DataFrame(by_region_rows)

    notes = [
        "UNITS",
        "'Daily': MWh generated that day by source, then the same as mean GW (MWh / 24 / 1000). "
        "'Plants': MWh per large plant, one column each. 'Plants by region': the same plant-level MWh reshaped "
        "long, with province/region/country columns added. 'Half-hourly': MW.",
        "",
        "SOURCES",
        "hydro, thermal (split into oil/diesel 'Termica' and natural gas), non-conventional renewables, imports, "
        "exports - as CENACE reports them. Data are CENACE's preliminary SCADA values.",
        "",
        "REGIONS",
        "CENACE reports no location for any plant - province/region in 'Plants by region' is this script's own "
        "mapping (PLANT_LOCATION), looked up by hand: province, then Ecuador's standard Costa/Sierra/Oriente "
        "zoning. 'Otras Hidro' is CENACE's own catch-all for every smaller hydro plant not named individually, so "
        "it can't be assigned a single location - tagged region 'Nacional (agregado)' instead of guessed.",
        "",
        "SOURCE",
        "CENACE Informacion Operativa (cenace.gob.ec/info-operativa/InformacionOperativa.htm), 'Informacion "
        "operativa diaria' tab - the last complete day only, so history starts at the first run and grows daily.",
    ]
    sheets = {"Daily": pd.concat([sources, gw], axis=1), "Plants": plants_wide}
    if not by_region.empty:
        sheets["Plants by region"] = by_region
    if not half.empty:
        sheets["Half-hourly"] = half
    xlsx_notes.write_workbook(OUT_FILE, sheets, notes, {"UNITS", "SOURCES", "REGIONS", "SOURCE"})
    print(f"Saved {OUT_FILE}: {len(daily)} day(s) of history", flush=True)
    print(sources.tail().to_string(), flush=True)


if __name__ == "__main__":
    main()
