"""Shared helpers for the one-off "future" master workbooks (pipelines, plans, outlooks)."""
import io
import re
import time

import pandas as pd
import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Battery storage", "Pumped storage",
         "Other"]


def get(url, **kw):
    last = None
    for attempt in range(3):
        try:
            r = requests.get(url, headers=HEADERS, timeout=(10, 300), **kw)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            last = e
            time.sleep(5 * (attempt + 1))
    raise last


def header_row(raw, must):
    """Index of the first row containing every regex in `must` (case-insensitive)."""
    for i in range(min(len(raw), 60)):
        cells = " | ".join(str(x) for x in raw.iloc[i].tolist())
        if all(re.search(m, cells, re.I) for m in must):
            return i
    raise ValueError(f"no header row with {must}")


def read_table(content, sheet, must):
    raw = pd.read_excel(io.BytesIO(content), sheet_name=sheet, header=None)
    h = header_row(raw, must)
    d = pd.read_excel(io.BytesIO(content), sheet_name=sheet, header=h)
    d.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in d.columns]
    return d.dropna(how="all")


def col(d, *patterns, required=True):
    for p in patterns:
        for c in d.columns:
            if re.search(p, str(c), re.I):
                return c
    if required:
        raise KeyError(f"no column matching {patterns} in {list(d.columns)[:40]}")
    return None


def by_year_fuel(d, year_col, fuel_col, mw_col, start=None, end=None):
    """Long rows -> year x fuel MW table (FUELS order, zero columns dropped)."""
    y = pd.to_numeric(d[year_col], errors="coerce")
    t = d.assign(_y=y, _mw=pd.to_numeric(d[mw_col], errors="coerce")).dropna(subset=["_y"])
    if start:
        t = t[t["_y"] >= start]
    if end:
        t = t[t["_y"] <= end]
    p = t.pivot_table(index="_y", columns=fuel_col, values="_mw", aggfunc="sum").fillna(0)
    p = p.reindex(columns=[f for f in FUELS if f in p.columns] + [c for c in p.columns if c not in FUELS])
    p = p.loc[:, p.sum() > 0]
    p.index = pd.to_datetime(p.index.astype(int).astype(str) + "-01-01")
    p.index.name = "Year"
    return p.round(1)


def fuel_from_text(text):
    """Technology / fuel description -> standard fuel group."""
    t = str(text).lower()
    rules = [(r"pump", "Pumped storage"), (r"batter|bess|storage|flywheel", "Battery storage"),
             (r"wind", "Wind"), (r"solar|photovolt|\bpv\b|fotovolt", "Solar"), (r"hydro|hidr|water|pch|cgh|uhe",
                                                                               "Hydro"),
             (r"nuclear|nucle", "Nuclear"), (r"coal|carv|carb|lignite", "Coal"),
             (r"gas|ccgt|ocgt|combined cycle|ciclo combinado|turbina de gas|gn", "Gas"),
             (r"diesel|oil|petrol|fuel oil|distillate|kerosene|óleo|oleo|gasoil", "Oil"),
             (r"bio|bagasse|bagaço|wood|biomass|landfill|biogas|residuo|resíduo", "Bioenergy")]
    for pat, f in rules:
        if re.search(pat, t):
            return f
    return "Other"
