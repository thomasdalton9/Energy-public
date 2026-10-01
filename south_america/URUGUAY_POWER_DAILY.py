"""
Uruguay daily power generation by type from ADME (Administracion del
Mercado Electrico) - scheduled entrypoint in the standard layout shared by
the South America raw grid-operator pulls (see power_daily_std.py).

Source: ADME "Generacion por fuente", https://pronos.adme.com.uy/gpf.php
(fecha_ini / fecha_fin, dd/mm/yyyy) -> hourly SCADA .ods
(/cache/gpf_<a>_<b>_horario.ods), read with URUGUAY_ADME.py's own
functions. History from 2021-01-01 (gpf.php goes back to 2019).

Each run fetches only days missing from the workbook (plus the last 3,
which ADME can revise), in 31-day requests.

Usage: python3 URUGUAY_POWER_DAILY.py [--out PATH] [--start YYYY-MM-DD]
The owner's local URUGUAY_ADME.py is unchanged; this only imports it.
"""

print("STARTING", flush=True)

import argparse
import datetime as dt
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import URUGUAY_ADME as U  # noqa: E402
import power_daily_std as std  # noqa: E402

OUT = "output/Data and Chart Outputs/uruguay_power_generation_daily.xlsx"

# ADME 'GPF' sheet column -> standard fuel
MAPPING = {
    "Salto Grande": "Hydro", "Bonete": "Hydro", "Baygorria": "Hydro", "Palmar": "Hydro",
    "Eólica": "Wind",
    "Solar": "Solar",
    "Térmica": "Oil",
    "Biomasa": "Bioenergy",
}
DETAIL = ["Salto Grande", "Bonete", "Baygorria", "Palmar", "Demanda", "Imports", "Exports"]

NOTES = [
    "UNITS",
    "MWh per day. ADME publishes hourly average MW (SCADA); a day's MWh = the mean of its hourly MW x 24. "
    "Days with fewer than 23 hourly readings are left blank and fetched again on the next run.",
    "",
    "CATEGORY MAPPING (ADME 'GPF' column -> sheet 'Daily')",
    "Hydro_MWh = Salto Grande + Bonete (Rincon del Bonete) + Baygorria + Palmar. Salto Grande is the "
    "Argentina/Uruguay binational dam; ADME reports Uruguay's half only, which is what is counted here.",
    "Wind_MWh = 'Eolica' (all wind farms). Solar_MWh = 'Solar' (all PV plants).",
    "Oil_MWh = 'Termica': Central Batlle engines (fuel oil), La Tablada and Punta del Tigre A gas turbines "
    "(gasoil) and the Punta del Tigre B combined cycle (gasoil; it can burn natural gas but Uruguay has had "
    "almost no gas supply). Ember counts the same plants as 'Other Fossil'. Uruguay has no coal or nuclear and "
    "no material gas-fired output, so Gas/Coal/Nuclear columns are left out.",
    "Bioenergy_MWh = 'Biomasa' (pulp mills UPM / Montes del Plata, sawmills, Alur, etc.).",
    "Total_MWh = sum of the fuel columns (domestic generation; imports excluded).",
    "Sheet 'Detail': MWh per day for each hydro plant, demand ('Demanda'), imports (Argentina + Brazil Rivera "
    "and Melo) and exports, same x24 convention.",
    "",
    "SOURCE",
    "ADME 'Generacion por fuente' - https://pronos.adme.com.uy/gpf.php (hourly .ods, SCADA values, "
    "described by ADME as approximate). History in this workbook from 2021-01-01; updated daily by GitHub "
    "Actions (uruguay_power_generation_daily.yml), fetching only missing days plus the last 3.",
]


def fetch(days):
    frames = []
    for a, b in std.ranges(days, U.CHUNK_DAYS):
        print(f"ADME {a}..{b}", flush=True)
        df = U.fetch_range(a, b)
        if not df.empty:
            frames.append(df)
    if not frames:
        return pd.DataFrame()
    hourly = pd.concat(frames)
    return hourly[~hourly.index.duplicated(keep="last")].sort_index()


def to_daily(hourly):
    hourly = hourly.copy()
    hourly["Imports"] = hourly[[c for c in U.IMPORTS if c in hourly]].sum(axis=1, min_count=1)
    by_day = hourly.groupby(hourly.index.normalize())
    counts = by_day.size()
    mean = by_day.mean()
    mean = mean[counts.reindex(mean.index) >= 23]
    fuels = pd.DataFrame(index=mean.index)
    for fuel in dict.fromkeys(MAPPING.values()):
        cols = [c for c, f in MAPPING.items() if f == fuel and c in mean]
        fuels[fuel] = (mean[cols] * 24).sum(axis=1, min_count=1) if cols else float("nan")
    unknown = [c for c in mean.columns if c not in MAPPING and c not in U.IMPORTS + ["Imports", "Exports", "Demanda"]]
    if unknown:
        print(f"  WARNING: ADME columns not mapped (left out): {unknown}", flush=True)
    detail = (mean[[c for c in DETAIL if c in mean]] * 24).round(1)
    detail.index.name = "date"
    return std.standardise(fuels), detail


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start", type=dt.date.fromisoformat, default=std.HISTORY_START)
    args = ap.parse_args()

    existing = std.load_sheet(args.out, "Daily")
    existing_detail = std.load_sheet(args.out, "Detail")
    end = dt.date.today() - dt.timedelta(days=1)
    days = std.missing_days(existing, args.start, end)
    print(f"{len(existing):,} days saved; fetching {len(days):,}", flush=True)
    hourly = fetch(days) if days else pd.DataFrame()
    if hourly.empty and existing.empty:
        print("No data returned.", flush=True)
        sys.exit(1)
    new, new_detail = to_daily(hourly) if not hourly.empty else (pd.DataFrame(), pd.DataFrame())
    daily = std.merge(new, existing)
    detail = std.merge(new_detail, existing_detail)
    std.write(args.out, daily, NOTES, {"Detail": detail})


if __name__ == "__main__":
    main()
