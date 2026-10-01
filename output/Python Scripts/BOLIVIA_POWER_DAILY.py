"""
Bolivia daily power generation by type from CNDC (Comite Nacional de
Despacho de Carga) - scheduled entrypoint in the standard layout shared by
the South America raw grid-operator pulls (see power_daily_std.py).

Sources (CNDC's public WordPress REST API, no key):
  1. Monthly "Generacion Bruta y Demanda Max. Instantanea" workbooks
     (gen_dia_MMYY.xlsx): gross generation per PLANT per day, MWh - the
     official figures, published after each month closes. Listed with
     https://www.cndc.bo/wp-json/cndc/v1/estadisticas/documentos?categoria_id=225&desde=..&hasta=..
     and parsed with BOLIVIA_CNDC.py's parse_gen_dia(). History from 2021-01.
  2. https://www.cndc.bo/wp-json/cndc/v1/rt/generacion?fecha=YYYY-MM-DD -
     CNDC's real-time/post-dispatch series by technology (TERMO, HIDRO,
     SOLAR, EOL, BAGAZO; 96 quarter-hour MW values per day), available from
     Jan-2025 only. Used for the days after the latest monthly workbook (so
     the series runs to yesterday) and for any day a monthly workbook leaves
     blank; replaced by the monthly figures once those are published.

Usage: python3 BOLIVIA_POWER_DAILY.py [--out PATH] [--start YYYY-MM-DD] [--no-rt]
The owner's local BOLIVIA_CNDC.py is unchanged; this only imports it.
"""

print("STARTING", flush=True)

import argparse
import datetime as dt
import os
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import BOLIVIA_CNDC as B  # noqa: E402
import power_daily_std as std  # noqa: E402

OUT = "output/Data and Chart Outputs/bolivia_power_generation_daily.xlsx"
API = "https://www.cndc.bo/wp-json/cndc/v1/"
GEN_CATEGORY = 225  # "Generacion Bruta y Demanda Max. Instantanea"
RT_START = dt.date(2025, 1, 1)
MIN_DAY_MWH = 15000  # a monthly-workbook day below this is treated as incomplete (Bolivia runs ~25-40 GWh/day)

# gen_dia plant column (CNDC's names; a unit suffix such as 'ARANJUEZ MG' is matched on the plant name) -> fuel
PLANT_FUEL = {
    # hydro
    "CORANI": "Hydro", "S. ISABEL": "Hydro", "SAN JOSE I": "Hydro", "SAN JOSE II": "Hydro", "MISICUNI": "Hydro",
    "KANATA": "Hydro", "ZONGO": "Hydro", "TAQUESI": "Hydro", "MIGUILLA": "Hydro", "MIGUILLAS": "Hydro",
    "YURA": "Hydro", "SAN JACINTO": "Hydro", "QUEHATA": "Hydro", "SDB": "Hydro",
    # natural gas (gas turbines, combined cycles, gas engines)
    "WARNES": "Gas", "SUR": "Gas", "ENTRE RIOS": "Gas", "GUARACACHI": "Gas", "SCZ": "Gas", "CARRASCO": "Gas",
    "BULO BULO": "Gas", "V. HERMOSO": "Gas", "ARANJUEZ": "Gas", "EL ALTO": "Gas", "C.EL ALTO": "Gas",
    # diesel (Beni systems joined to the grid)
    "MOXOS": "Oil", "Rurrenabaque": "Oil", "S.A. Yacuma": "Oil", "S.I. Moxos": "Oil", "San Borja": "Oil",
    "Yucumo": "Oil",
    # sugar-mill bagasse
    "AGUAÍ": "Bioenergy", "AGU02": "Bioenergy", "AGUAÍ ENERGÍA": "Bioenergy", "GUABIRA": "Bioenergy",
    "GBE": "Bioenergy", "UNAGRO": "Bioenergy", "EASBA": "Bioenergy",
    # wind
    "QOLLPANA": "Wind", "QOLL I": "Wind", "QOLL II": "Wind", "EDORADO": "Wind", "EEDORADO": "Wind", "EDO": "Wind",
    "ESJULIAN": "Wind", "ESJU": "Wind", "EWARNES": "Wind", "EWA": "Wind",
    # solar
    "ORURO I": "Solar", "ORURO II": "Solar", "UYUNI": "Solar", "UYUNI I": "Solar", "UYUNI II": "Solar",
    "YUNCHARA": "Solar",
}
RT_FUEL = {"HIDRO": "Hydro", "TERMO": "Gas", "EOL": "Wind", "SOLAR": "Solar", "BAGAZO": "Bioenergy", "RENO": "Other"}

NOTES = [
    "UNITS",
    "MWh per day (gross generation).",
    "",
    "CATEGORY MAPPING (CNDC plant -> sheet 'Daily')",
    "Hydro_MWh: Corani, Santa Isabel, San Jose I/II, Misicuni, Kanata, Zongo, Taquesi, Miguillas, Yura, "
    "San Jacinto, Quehata (SDB).",
    "Gas_MWh: Warnes, Del Sur, Entre Rios, Guaracachi, Santa Cruz, Carrasco, Bulo Bulo, Valle Hermoso, "
    "Aranjuez (incl. its small dual-fuel engines), El Alto (Kenko) - natural-gas turbines, combined cycles and "
    "engines.",
    "Oil_MWh: Moxos and the former isolated Beni diesel plants (Rurrenabaque, Santa Ana de Yacuma, San Ignacio "
    "de Moxos, San Borja, Yucumo).",
    "Bioenergy_MWh: sugar-mill bagasse plants Aguai, Guabira, Unagro, EASBA. CNDC files these under 'termo' in "
    "its own totals and Ember shows no Bolivian bioenergy, so Ember's Gas includes them.",
    "Wind_MWh: Qollpana, El Dorado, San Julian, Warnes wind. Solar_MWh: Solar Oruro I/II, Uyuni, Yunchara.",
    "A plant not in this mapping goes to Other_MWh (and is printed as a warning by the pull).",
    "Days from CNDC's real-time series (see 'Source' sheet) map TERMO -> Gas, HIDRO -> Hydro, EOL -> Wind, "
    "SOLAR -> Solar, BAGAZO -> Bioenergy, RENO -> Other; the few diesel MWh are inside TERMO there. MW per "
    "quarter-hour is averaged over the day and x24.",
    "",
    "SOURCE",
    "CNDC monthly statistics 'Generacion Bruta y Demanda Max. Instantanea' (gen_dia_MMYY.xlsx, daily gross "
    "generation per plant) via https://www.cndc.bo/wp-json/cndc/v1/estadisticas/documentos (categoria 225). "
    "Each month appears after it closes, so for the latest weeks the sheet uses CNDC's real-time generation "
    "by technology (https://www.cndc.bo/wp-json/cndc/v1/rt/generacion?fecha=, available from Jan-2025), "
    "replaced by the monthly figures once published. Sheet 'Source' says which one each day uses; 'By plant' "
    "has the monthly workbooks' plant columns.",
    "History from 2021-01-01. Updated daily by GitHub Actions (bolivia_power_generation_daily.yml): "
    "only monthly workbooks with missing days (plus the latest two) and real-time days not yet covered are "
    "downloaded.",
]


def session():
    s = requests.Session()
    s.headers.update(B.HEADERS)
    return s


def gen_dia_months(s, start):
    r = s.get(API + "estadisticas/documentos", params={"categoria_id": GEN_CATEGORY, "desde": start.isoformat(),
                                                       "hasta": dt.date.today().isoformat()}, timeout=B.TIMEOUT)
    r.raise_for_status()
    out = {}
    for g in r.json().get("grupos", []):
        url = next((d["archivo_url"] for d in g.get("docs", []) if d.get("archivo_url", "").endswith(".xlsx")), None)
        if url:
            out[g["periodo"]] = url
    return dict(sorted(out.items()))


def plant_fuel(col):
    name = str(col).strip()
    if name in PLANT_FUEL:
        return PLANT_FUEL[name]
    base = name.rsplit(" ", 1)[0]  # 'ARANJUEZ MG' -> 'ARANJUEZ'
    return PLANT_FUEL.get(base)


def plants_to_fuels(by_plant):
    unknown = [c for c in by_plant.columns if plant_fuel(c) is None]
    if unknown:
        print(f"  WARNING: plants not in PLANT_FUEL (-> Other): {unknown}", flush=True)
    fuels = pd.DataFrame(index=by_plant.index)
    for fuel in std.FUELS:
        cols = [c for c in by_plant.columns if (plant_fuel(c) or "Other") == fuel]
        if cols:
            fuels[fuel] = by_plant[cols].apply(pd.to_numeric, errors="coerce").sum(axis=1, min_count=1)
    return std.standardise(fuels)


def rt_day(s, day):
    r = s.get(API + "rt/generacion", params={"fecha": day.isoformat()}, timeout=B.TIMEOUT)
    r.raise_for_status()
    row = {}
    for series in r.json() or []:
        vals = [v for v in series.get("valores") or [] if v is not None]
        code = series.get("codigo")
        if code in RT_FUEL and len(vals) >= 90:
            row[RT_FUEL[code]] = row.get(RT_FUEL[code], 0.0) + sum(vals) / len(vals) * 24
        elif code not in RT_FUEL and code not in ("PREV", "TOT"):
            print(f"  WARNING: rt/generacion series {code!r} not mapped", flush=True)
    return row


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start", type=dt.date.fromisoformat, default=std.HISTORY_START)
    ap.add_argument("--no-rt", action="store_true", help="monthly workbooks only")
    args = ap.parse_args()
    s = session()
    yesterday = dt.date.today() - dt.timedelta(days=1)

    by_plant = std.load_sheet(args.out, "By plant")
    rt_saved = std.load_sheet(args.out, "Real-time by type")

    # 1. monthly workbooks: any month with days missing from 'By plant', plus the latest two
    months = gen_dia_months(s, args.start)
    have = {d.date() for d in by_plant.index} if not by_plant.empty else set()
    latest = list(months)[-2:]
    fresh = []
    for period, url in months.items():
        y, m = map(int, period.split("-"))
        first = dt.date(y, m, 1)
        days = pd.date_range(first, (pd.Timestamp(first) + pd.offsets.MonthEnd(0)).date()).date
        if period not in latest and all(d in have for d in days):
            continue
        try:
            r = s.get(url, timeout=B.TIMEOUT)
            r.raise_for_status()
            frame = B.parse_gen_dia(r.content)
            frame = frame[(frame.index >= pd.Timestamp(first)) & (frame.index <= pd.Timestamp(days[-1]))]
            fresh.append(frame)
            print(f"  gen_dia {period}: {len(frame)} days, {frame.apply(pd.to_numeric, errors='coerce').sum().sum():,.0f} MWh",
                  flush=True)
        except Exception as e:  # noqa: BLE001 - keep going, the month is retried next run
            print(f"  gen_dia {period}: FAILED ({type(e).__name__}: {e})", flush=True)
        time.sleep(0.3)
    if fresh:
        new = pd.concat(fresh)
        new.index = pd.to_datetime(new.index)
        by_plant = std.merge(new, by_plant)
    by_plant.index.name = "date"
    by_plant = by_plant[by_plant.index >= pd.Timestamp(args.start)]
    monthly_daily = plants_to_fuels(by_plant) if not by_plant.empty else pd.DataFrame()
    good = monthly_daily[monthly_daily["Total_MWh"] >= MIN_DAY_MWH] if not monthly_daily.empty else monthly_daily

    # 2. real-time series for the days the monthly workbooks don't cover (yet)
    if not args.no_rt:
        covered = {d.date() for d in good.index}
        rt_have = {d.date() for d in rt_saved.index} if not rt_saved.empty else set()
        want = [d for d in pd.date_range(max(RT_START, args.start), yesterday).date if d not in covered]
        refresh = sorted(rt_have)[-3:]
        # plus the latest month the monthly workbooks cover, once, to compare the two sources
        check = [d for d in covered if d >= RT_START and (d.year, d.month) == max((c.year, c.month) for c in covered)] \
            if covered else []
        todo = sorted({d for d in want + check if d not in rt_have} | {d for d in refresh if d in want})
        print(f"real-time days to fetch: {len(todo)}", flush=True)
        rows = {}
        for d in todo:
            try:
                row = rt_day(s, d)
                if row:
                    rows[pd.Timestamp(d)] = row
            except (requests.RequestException, ValueError) as e:
                print(f"  rt {d}: FAILED ({type(e).__name__})", flush=True)
            time.sleep(0.2)
        if rows:
            rt_saved = std.merge(std.standardise(pd.DataFrame.from_dict(rows, orient="index")), rt_saved)
        # how the two sources compare where both exist (latest full month of overlap)
        if not rt_saved.empty and not good.empty:
            both = good.index.intersection(rt_saved.index)
            if len(both):
                last = both.max().to_period("M")
                sel = [d for d in both if d.to_period("M") == last]
                ratio = rt_saved.loc[sel, "Total_MWh"].sum() / good.loc[sel, "Total_MWh"].sum()
                print(f"  check {last}: real-time total / monthly-workbook total = {ratio:.3f} over {len(sel)} days",
                      flush=True)

    rt_used = rt_saved[~rt_saved.index.isin(good.index)] if not rt_saved.empty else rt_saved
    daily = std.merge(good, rt_used) if not rt_used.empty else good
    daily = std.standardise(daily.drop(columns="Total_MWh"))
    if daily.empty:
        print("No data.", flush=True)
        sys.exit(1)
    source = pd.DataFrame({"source": ["CNDC monthly gen_dia (by plant)" if d in good.index else "CNDC real-time "
                                      "rt/generacion (provisional)" for d in daily.index]}, index=daily.index)
    source.index.name = "date"
    extra = {"Source": source, "By plant": by_plant.round(2)}
    if not rt_saved.empty:
        extra["Real-time by type"] = rt_saved
    std.write(args.out, daily, NOTES, extra)


if __name__ == "__main__":
    main()
