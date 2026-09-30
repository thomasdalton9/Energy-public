"""
Chile hydro reservoir levels (cota, metres above sea level) from CEN
(Coordinador Electrico Nacional)'s public API.

Unlike Brazil/Colombia, this can't produce a %-of-capacity or 5-year
range chart yet:
  - The only reliable CEN endpoint returns just the LATEST reading per
    reservoir (embalse-real/v3/findLast) - the documented date-range
    endpoint (cotas-embalses-reales/v3/findAll) is confirmed broken
    server-side (returns "Internal server error"), per a public
    community project polling the same API - there's no bulk history
    to pull.
  - The value itself is cota (water level in metres above sea level),
    not a %-of-capacity figure - converting one to the other needs
    each reservoir's own level-to-volume rule curve, which isn't
    published anywhere found.

So this script polls the latest cota per reservoir on every pipeline
run and appends it to its own growing history CSV
(chile_hydro_cota_history.csv), replacing today's row if it's already
run today. The dashboard chart is a straight line-per-reservoir plot
in metres, starting from whenever this first ran - not backfilled.
Revisit if CEN's range endpoint gets fixed, or a rule-curve source
turns up to convert cota to a proper %-of-capacity figure.

Requires a CEN_USER_KEY in api_keys.py - register for one (free) at
https://sipub.coordinador.cl/.
"""

print("STARTING", flush=True)

import os
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for api_keys
import api_keys

BASE_DIRECTORY = Path(__file__).resolve().parent
HISTORY_CSV = BASE_DIRECTORY / "chile_hydro_cota_history.csv"

CEN_USER_KEY = api_keys.CEN_USER_KEY
API_URL = "https://sipub.api.coordinador.cl/embalse-real/v3/findLast"

if not CEN_USER_KEY:
    raise RuntimeError(
        "api_keys.CEN_USER_KEY is not set - register for a free key at "
        "https://sipub.coordinador.cl/ and add it to api_keys.py."
    )


def fetch_latest_cotas():
    r = requests.get(
        API_URL,
        params={"user_key": CEN_USER_KEY},
        headers={"accept": "application/json"},
        timeout=60,
    )
    r.raise_for_status()
    records = r.json()

    if not isinstance(records, list) or not records:
        raise RuntimeError(f"Unexpected response from CEN API: {records!r}")

    df = pd.json_normalize(records)

    if "nombre" not in df.columns or "cotaActual" not in df.columns:
        raise KeyError(
            f"Expected 'nombre'/'cotaActual' columns, got {list(df.columns)} - "
            "CEN may have changed its response shape."
        )

    df = df[["nombre", "cotaActual"]].copy()
    df.columns = ["Reservoir", "Cota"]
    df["Reservoir"] = df["Reservoir"].astype(str).str.strip().str.title()
    df["Cota"] = pd.to_numeric(df["Cota"], errors="coerce")
    return df.dropna(subset=["Cota"])


def load_history():
    if HISTORY_CSV.exists():
        df = pd.read_csv(HISTORY_CSV)
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        return df.dropna(subset=["Date"])
    return pd.DataFrame(columns=["Date", "Reservoir", "Cota"])


# ----------------------------------------------------------
# PULL + APPEND
# ----------------------------------------------------------

today = pd.Timestamp(date.today())

print("Fetching latest reservoir cotas from CEN...", flush=True)
latest = fetch_latest_cotas()
latest.insert(0, "Date", today)

history = load_history()
history = history[history["Date"] != today]  # replace today's row if this already ran today
history = pd.concat([history, latest], ignore_index=True)
history = history.sort_values(["Date", "Reservoir"])

history.to_csv(HISTORY_CSV, index=False)

print(f"\n{len(latest)} reservoirs, {today.date()}:", flush=True)
print(
    latest.sort_values("Reservoir").to_string(index=False, float_format="%.2f"),
    flush=True,
)

print(
    f"\nSaved: {HISTORY_CSV} "
    f"({len(history):,} total rows across {history['Date'].nunique()} day(s))",
    flush=True,
)
