"""
AEMO NEM Registration and Exemption List -> one row per generating unit (DUID) with NEM region, standard fuel and
registered capacity. Shared by AU_NEM_GENERATION.py (unit -> fuel for the SCADA) and AU_POWER_CAPACITY.py.
Descriptor vocabularies come from rest_of_world/aemo_nemweb_power_mix.py; a descriptor not listed there maps to
Other (and is printed) instead of stopping the run.
"""
import io
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "rest_of_world"))
import aemo_nemweb_power_mix as nem  # noqa: E402

CATEGORY = {"coal": "Coal", "gas": "Gas", "oil": "Oil", "hydro": "Hydro", "wind": "Wind", "solar": "Solar",
            "biomass": "Bioenergy", "nuclear": "Nuclear", "battery": "Battery_storage",
            "hydro_pumped": "Pumped_storage"}
URL = nem.REGISTRATION_LIST_URL
# regex fallbacks on "fuel | technology" for descriptor variants (first match wins)
FALLBACK = [(r"batter", "Battery_storage"), (r"pump", "Pumped_storage"), (r"solar|photovoltaic", "Solar"),
            (r"wind", "Wind"), (r"hydro|water", "Hydro"),
            (r"landfill|biogas|sewer|waste water|bagasse|wood|biomass", "Bioenergy"),
            (r"natural gas|natrual gas|ethane|coal mine gas|coal seam|methane|\bgas\b", "Gas"),
            (r"coal", "Coal"), (r"diesel|kerosene|fuel oil|distillate", "Oil")]


def _text(v):
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()


def units(include_loads=False):
    """One row per DUID: duid, region, fuel, mw, role ('gen', or 'load' for scheduled loads - battery charging and
    hydro pumps - which are only kept with include_loads=True)."""
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
    is_load = d["Dispatch Type"].map(_text).str.contains("load", case=False) if "Dispatch Type" in d else \
        pd.Series(False, index=d.index)
    if not include_loads:   # generating units only (scheduled loads are listed too)
        d, is_load = d[~is_load], is_load[~is_load]
    unknown = set()

    def fuel(row):
        tech, src = _text(row.get("Technology Type - Descriptor")), _text(row.get("Fuel Source - Descriptor"))
        c = nem.TECHNOLOGY_OVERRIDE_TO_CATEGORY.get(tech) or nem.FUEL_DESCRIPTOR_TO_CATEGORY.get(src)
        if c is not None:
            return CATEGORY[c]
        for pat, f in FALLBACK:   # descriptor variants not in the exact-match vocabulary
            if re.search(pat, f"{src} | {tech}", re.I):
                return f
        unknown.add((src, tech))
        return "Other" if src or tech else "Unknown"

    capcol = next((c for c in d.columns if "reg cap" in c.lower() and "gen" in c.lower()),
                  next((c for c in d.columns if "reg cap" in c.lower()), None))
    out = pd.DataFrame({"duid": d["DUID"].map(_text), "region": d["Region"].str[:-1], "fuel": d.apply(fuel, axis=1),
                        "mw": pd.to_numeric(d[capcol], errors="coerce") if capcol else float("nan"),
                        "role": is_load.map({True: "load", False: "gen"})})
    station = next((c for c in d.columns if "station name" in c.lower()), None)
    if include_loads and station:   # a scheduled load with blank descriptors (battery charging, hydro pumps) takes
        st = d[station].map(_text)  # the fuel of the generating units at the same station
        gen_fuel = out[out["role"].eq("gen") & ~out["fuel"].isin(["Unknown", "Other"])].assign(st=st).groupby("st")["fuel"].first()
        fill = out["role"].eq("load") & out["fuel"].isin(["Unknown", "Other"])
        out.loc[fill, "fuel"] = st[fill].map(gen_fuel).fillna(out.loc[fill, "fuel"])
        print(f"  registration list: {fill.sum()} scheduled loads without descriptors, "
              f"{out.loc[fill, 'fuel'].isin(list(CATEGORY.values())).sum()} matched to their station's fuel", flush=True)
    out = out.drop_duplicates("duid")
    blank = out["fuel"].eq("Unknown")
    if blank.any():   # rows with no fuel or technology at all (loads, placeholders) are not generating capacity
        print(f"  registration list: {blank.sum()} units with blank descriptors ({out.loc[blank, 'mw'].sum():,.0f} MW) "
              f"left out: {sorted(out.loc[blank, 'duid'])[:30]}", flush=True)
        out = out[~blank]
    if "Dispatch Type" in d:
        print(f"  dispatch types kept: {d['Dispatch Type'].map(_text).value_counts().to_dict()}", flush=True)
    if unknown:
        print(f"  registration list: descriptors mapped to Other: {sorted(unknown)}", flush=True)
    print(f"  registration list: {len(out)} NEM units, {out['mw'].sum() / 1000:.1f} GW ({capcol})", flush=True)
    return out


# units retired before the current registration list whose station has no current unit to borrow a fuel from
RETIRED = {"LD0": "Coal", "TORR": "Gas", "OSB-AG": "Gas", "SWAN": "Gas", "MACKAYGT": "Oil", "SNUG": "Oil",
           "LONSDALE": "Oil", "PTSTAN": "Oil", "ANGAST": "Oil", "DRYCGT": "Gas", "MSTUART": "Oil", "QPS": "Gas",
           "PORTWF": "Wind", "PIONEER": "Bioenergy"}   # Portland wind farm (VIC), Pioneer sugar mill (QLD, bagasse)


def latest_mms(table, name, start):
    """An MMS register table (DUDETAILSUMMARY, DUDETAIL: each monthly archive holds the full register) from the
    newest monthly archive AEMO has published - the current month's lags by a month or two."""
    today = pd.Timestamp.today().normalize()
    last = None
    for back in range(0, 5):
        end = (today - pd.offsets.MonthBegin(back + 1)) if back else today
        try:
            return table(name, start, end)
        except Exception as e:  # noqa: BLE001 - archive not published yet: one month back
            last = e
            print(f"  {name} to {end:%Y-%m}: not available ({type(e).__name__}); trying a month earlier", flush=True)
    raise last


def with_history(current, start, end, table):
    """Adds units that left the registration list (retired generators; the separate battery-charging and pump LOAD
    DUIDs that AEMO folded into bidirectional units in 2024) from MMS DUDETAILSUMMARY, which keeps every DUID's
    region, station and dispatch type. A missing unit takes the fuel of a current unit at the same station.
    table(name, start, end) is the caller's nemosis fetch; any failure keeps the current list."""
    try:
        h = latest_mms(table, "DUDETAILSUMMARY", start)
    except Exception as e:  # noqa: BLE001
        print(f"  DUDETAILSUMMARY failed ({type(e).__name__}: {str(e)[:150]}); current registration list only",
              flush=True)
        return current
    h = h.sort_values("START_DATE").drop_duplicates("DUID", keep="last")
    h = h[h["REGIONID"].isin(nem.REGION_TO_STATE_NAME)]
    station = h.set_index("DUID")["STATIONID"]
    st_fuel = current.assign(st=current["duid"].map(station)).dropna(subset=["st"])
    st_fuel = st_fuel[st_fuel["role"].eq("gen")].groupby("st")["fuel"].agg(lambda f: f.mode().iat[0])
    new = h[~h["DUID"].isin(current["duid"])]
    fuel = new["STATIONID"].map(st_fuel)
    for pre, f in RETIRED.items():
        fuel = fuel.where(fuel.notna() | ~new["DUID"].str.upper().str.startswith(pre), f)
    load = new["DISPATCHTYPE"].astype(str).str.upper().str.contains("LOAD")
    add = pd.DataFrame({"duid": new["DUID"], "region": new["REGIONID"].str[:-1], "fuel": fuel, "mw": float("nan"),
                        "role": load.map({True: "load", False: "gen"})}).dropna(subset=["fuel"])
    miss = new[fuel.isna()]
    print(f"  DUDETAILSUMMARY: {len(new)} DUIDs not in the current list, {len(add)} mapped "
          f"({(add['role'] == 'load').sum()} loads: {sorted(add.loc[add['role'] == 'load', 'duid'])[:40]}); "
          f"unmapped: {sorted(zip(miss['DUID'], miss['STATIONID']))[:60]}", flush=True)
    return pd.concat([current, add], ignore_index=True)
