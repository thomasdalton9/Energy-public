"""
AEMO NEM Registration and Exemption List -> one row per generating unit (DUID) with NEM region, standard fuel and
registered capacity. Shared by AU_NEM_GENERATION.py (unit -> fuel for the SCADA) and AU_POWER_CAPACITY.py.
Descriptor vocabularies come from rest_of_world/aemo_nemweb_power_mix.py; a descriptor not listed there maps to
Other (and is printed) instead of stopping the run.
"""
import io
import os
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "rest_of_world"))
import aemo_nemweb_power_mix as nem  # noqa: E402

CATEGORY = {"coal": "Coal", "gas": "Gas", "oil": "Oil", "hydro": "Hydro", "wind": "Wind", "solar": "Solar",
            "biomass": "Bioenergy", "nuclear": "Nuclear", "battery": "Battery_storage",
            "hydro_pumped": "Pumped_storage"}
URL = nem.REGISTRATION_LIST_URL


def _text(v):
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()


def units():
    r = requests.get(URL, headers=nem.HEADERS, timeout=(10, 120))
    r.raise_for_status()
    d = None
    for sh in nem.REGISTRATION_SHEET_NAMES:
        try:
            d = pd.read_excel(io.BytesIO(r.content), sheet_name=sh, dtype=str, engine="openpyxl")
            break
        except ValueError:
            continue
    if d is None:
        raise RuntimeError(f"none of {nem.REGISTRATION_SHEET_NAMES} in the registration list")
    d = d[d["DUID"].map(_text).ne("") & d["Region"].isin(nem.REGION_TO_STATE_NAME)]
    if "Dispatch Type" in d:   # generating units only (scheduled loads are listed too)
        d = d[~d["Dispatch Type"].map(_text).str.contains("load", case=False)]
    unknown = set()

    def fuel(row):
        tech, src = _text(row.get("Technology Type - Descriptor")), _text(row.get("Fuel Source - Descriptor"))
        c = nem.TECHNOLOGY_OVERRIDE_TO_CATEGORY.get(tech) or nem.FUEL_DESCRIPTOR_TO_CATEGORY.get(src)
        if c is None:
            unknown.add((src, tech))
        return CATEGORY.get(c, "Other")

    capcol = next((c for c in d.columns if "reg cap" in c.lower() and "gen" in c.lower()),
                  next((c for c in d.columns if "reg cap" in c.lower()), None))
    out = pd.DataFrame({"duid": d["DUID"].map(_text), "region": d["Region"].str[:-1], "fuel": d.apply(fuel, axis=1),
                        "mw": pd.to_numeric(d[capcol], errors="coerce") if capcol else float("nan")})
    out = out.drop_duplicates("duid")
    if unknown:
        print(f"  registration list: descriptors mapped to Other: {sorted(unknown)}", flush=True)
    print(f"  registration list: {len(out)} NEM units, {out['mw'].sum() / 1000:.1f} GW ({capcol})", flush=True)
    return out
