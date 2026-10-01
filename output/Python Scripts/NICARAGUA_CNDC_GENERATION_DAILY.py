"""
Nicaragua daily power generation by type from CNDC (Centro Nacional de Despacho
de Carga, ENATREL) - raw grid-operator pull in the standard layout.

Source: CNDC "MEMN Diarios -> Postdespacho de Energia"
(https://www.cndc.org.ni/MEMNDiarios/posdespachoEnergia), whose page loads
  https://www.cndc.org.ni/MEMNDiarios/consultarPosdespachoEnergia?fecha=DD/MM/YYYY
-> hourly post-dispatch energy (MWh) for every generating unit (GGD) of every
market agent, plus the national demand and the interconnection lines. CNDC
does not give a fuel per unit in that table, so each unit code is mapped to a
technology here (UNITS below). CNDC's own by-type series
(Inicio/ConsultarTipoGeneracion, graficos/consultarGeneracionPorTipo) only covers
the current day, but it also gives installed capacity by type, which checks the
mapping: summing the post-dispatch units' 'potenciamaxima' by the mapping below
gives wind 195.9 MW (CNDC 195.5), solar 190.8 (190.7), hydro 138.2 (147.4),
geothermal 181.5 (151.5), bagasse 216.2 (174.5), thermal 750.1 (717.1); hourly
profiles confirm the solar (EMPROSA) and wind (ABR) units
(discovery_archive/south_america/CENTRAL_AMERICA_POWER_DISCOVERY4.py / 5.py).
History from 2021-01-01.

Each run fetches only days missing from the workbook plus the last 14 days,
within a time budget, saving as it goes. Units not in UNITS are counted as
Other_MWh and printed as a warning (so a new plant shows up, not silently lost).

Usage: python3 NICARAGUA_CNDC_GENERATION_DAILY.py [--out PATH] [--start YYYY-MM-DD]
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
import central_america_power_common as C  # noqa: E402

OUT = "output/Data and Chart Outputs/nicaragua_power_generation_daily.xlsx"
PAGE = "https://www.cndc.org.ni/MEMNDiarios/posdespachoEnergia"
URL = "https://www.cndc.org.ni/MEMNDiarios/consultarPosdespachoEnergia"

# CNDC unit (GGD) -> (standard fuel, plant / agent description)
UNITS = {
    # wind
    "AMY1": ("Wind", "Amayo I wind farm (Consorcio Eolico Amayo)"),
    "AMY2": ("Wind", "Amayo II wind farm (Consorcio Eolico Amayo II)"),
    "PBP": ("Wind", "Blue Power wind farm (Blue Power & Energy)"),
    "EOL": ("Wind", "Eolo de Nicaragua wind farm"),
    "ABR": ("Wind", "Camilo Ortega / Alba Rivas wind farm (ALBA de Nicaragua)"),
    # hydro
    "PCA1": ("Hydro", "ENEL Planta Centroamerica (hydro)"),
    "PCA2": ("Hydro", "ENEL Planta Centroamerica (hydro)"),
    "PCF1": ("Hydro", "ENEL Planta Carlos Fonseca (hydro)"),
    "PCF2": ("Hydro", "ENEL Planta Carlos Fonseca (hydro)"),
    "PHL1": ("Hydro", "ENEL Planta Hidroelectrica Larreynaga"),
    "PHL2": ("Hydro", "ENEL Planta Hidroelectrica Larreynaga"),
    "HPA1": ("Hydro", "Hidropantasma"),
    "HPA2": ("Hydro", "Hidropantasma"),
    "IHC": ("Hydro", "Interamerican Hydroelectric"),
    "PHD": ("Hydro", "Inversiones Hidroelectricas (IHSA)"),
    # geothermal -> Other
    "PEN3": ("Other", "Polaris Energy Nicaragua, San Jacinto-Tizate (geothermal)"),
    "PEN4": ("Other", "Polaris Energy Nicaragua, San Jacinto-Tizate (geothermal)"),
    "PEN5": ("Other", "Polaris Energy Nicaragua, San Jacinto-Tizate (geothermal)"),
    "PMT1": ("Other", "Momotombo (geothermal)"),
    "PMT2": ("Other", "Momotombo (geothermal)"),
    "PMT3": ("Other", "Momotombo (geothermal)"),
    # solar
    "PSL": ("Solar", "Solaris"),
    "PEJ": ("Solar", "Interamerican Solar (Intersolar)"),
    "PSI": ("Solar", "Nordic Solar"),
    "PSP": ("Solar", "Sun Power"),
    "PES1": ("Solar", "EMPROSA (ENACAL) solar park"),
    "PES2": ("Solar", "EMPROSA (ENACAL) solar park"),
    # bioenergy: sugar-mill bagasse cogeneration (seasonal, Nov-Jun harvest)
    "MTL": ("Bioenergy", "Cogeneracion Green Power (Ingenio Montelimar)"),
    "MTR": ("Bioenergy", "Monte Rosa (sugar mill)"),
    "NSL": ("Bioenergy", "Nicaragua Sugar Estates (Ingenio San Antonio)"),
    "EGR": ("Bioenergy", "EGERSA, Rivas (sugar-mill cogeneration)"),
    "GSR": ("Bioenergy", "Generadora San Rafael (GESARSA)"),
    # oil: bunker / diesel engines and steam units
    "CEN": ("Oil", "CENSA (bunker engines)"),
    "EEC": ("Oil", "Empresa Energetica Corinto (bunker barge)"),
    "EEC20": ("Oil", "Empresa Energetica Corinto (bunker)"),
    "TPC": ("Oil", "Tipitapa Power Company (bunker)"),
    "PNI1": ("Oil", "GEOSA Planta Nicaragua, Puerto Sandino (bunker steam)"),
    "PNI2": ("Oil", "GEOSA Planta Nicaragua, Puerto Sandino (bunker steam)"),
    "PLB1": ("Oil", "ENEL Planta Las Brisas (diesel)"),
    "PLB2": ("Oil", "ENEL Planta Las Brisas (diesel)"),
    "PMG3": ("Oil", "ENEL Planta Managua (bunker)"),
    "PMG4": ("Oil", "ENEL Planta Managua (bunker)"),
    "PMG5": ("Oil", "ENEL Planta Managua (bunker)"),
    "PMN": ("Oil", "ALBA Generacion (bunker)"),
    "PHC1": ("Oil", "ALBA Generacion, Planta Hugo Chavez (diesel)"),
    "PHC2": ("Oil", "ALBA Generacion, Planta Hugo Chavez (diesel)"),
    "HEM": ("Oil", "HEMCO Mineros (own thermal)"),
    **{f"PCG{i}": ("Oil", "ALBA de Nicaragua, Planta Che Guevara (bunker engines)") for i in range(1, 10)},
}
# agents whose rows are not generation
NOT_GENERATION = {"DEMANDA", "INT. NORTE", "INT. SUR"}
# whole agents with one technology (used for units not listed above)
AGENT_FUEL = {"NFE": "Gas"}  # New Fortress Energy (Nicaragua Development Partners), Puerto Sandino LNG

NOTES = [
    "UNITS",
    "MWh per day: CNDC post-dispatch energy per unit (posdespachoTotal, the sum of 24 hourly MWh), "
    "summed by technology.",
    "",
    "COVERAGE",
    "@COVERAGE",
    "",
    "SOURCE",
    "CNDC / ENATREL Nicaragua - MEMN Diarios, Postdespacho de Energia: " + PAGE,
    "Data: " + URL + "?fecha=DD/MM/YYYY (JSON, one day per call).",
    "Updated daily by GitHub Actions (nicaragua_power_generation.yml): only missing days plus the last 14.",
    "",
    "MAPPING",
    "CNDC reports energy per generating unit; each unit is mapped to a technology (full list below and in "
    "the script). Geothermal (Polaris San Jacinto-Tizate, Momotombo) is put in Other_MWh.",
    "Oil_MWh = bunker/diesel plants (Che Guevara, Hugo Chavez, ALBA Generacion, CENSA, Corinto, Tipitapa, "
    "Planta Nicaragua at Puerto Sandino, Las Brisas, Planta Managua, HEMCO).",
    "Gas_MWh: CNDC lists New Fortress Energy's Puerto Sandino LNG plant (agent NFE) as a market agent, but no "
    "NFE unit has any post-dispatch energy in 2021 to date, so Gas_MWh is 0; NFE units, if they appear, go to Gas.",
    "Bioenergy_MWh = sugar-mill bagasse cogeneration (Montelimar/Green Power, Monte Rosa, San Antonio/NSEL, "
    "EGERSA, GESARSA).",
    "Interconnection lines (agents INT. NORTE / INT. SUR) and DEMANDA are not generation and are left out; "
    "they are in sheet 'Detail' (Demand_MWh, Interconnection_MWh: + = export, - = import).",
    "Nicaragua has no coal or nuclear generation: those columns are 0.",
    "Total_MWh = sum of the fuel columns.",
    "Sheet 'Detail': MWh per day for every unit (column = AGENT|UNIT), as published.",
    "",
    "Unit -> technology:",
] + [f"{u} -> {f}_MWh ({desc})" for u, (f, desc) in UNITS.items()]


def fetch_day(session, day):
    r = session.get(URL, params={"fecha": day.strftime("%d/%m/%Y")}, timeout=(15, 90))
    r.raise_for_status()
    rows = r.json().get("posdespachoTotal") or []
    out = {}
    for x in rows:
        v = pd.to_numeric(str(x.get("total", "")).replace(",", "."), errors="coerce")
        out[f'{x.get("agente", "").strip()}|{x.get("ggd", "").strip()}'] = v
    return out or None


def to_frames(per_day):
    detail = pd.DataFrame(per_day).T
    detail.index = pd.to_datetime(detail.index)
    detail.index.name = "date"
    fuels = pd.DataFrame(index=detail.index)
    unknown = []
    inter = []
    for col in detail.columns:
        agent, unit = col.split("|", 1)
        if agent in NOT_GENERATION:
            if agent.startswith("INT"):
                inter.append(col)
            continue
        if unit in UNITS:
            fuel = UNITS[unit][0]
        elif agent in AGENT_FUEL:
            fuel = AGENT_FUEL[agent]
        else:
            unknown.append(col)
            fuel = "Other"
        fuels[fuel] = fuels.get(fuel, 0) + detail[col].fillna(0)
    if unknown:
        print(f"  WARNING: CNDC units not in UNITS, counted as Other: {unknown}", flush=True)
    detail = detail.copy()
    detail.insert(0, "Interconnection_MWh", detail[inter].sum(axis=1, min_count=1) if inter else float("nan"))
    if "DEMANDA|DEMANDA" in detail:
        detail.insert(0, "Demand_MWh", detail.pop("DEMANDA|DEMANDA"))
    return C.standardise(fuels), detail.round(2)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start", type=dt.date.fromisoformat, default=C.HISTORY_START)
    ap.add_argument("--refresh-days", type=int, default=14)
    ap.add_argument("--budget-min", type=float, default=50, help="stop fetching after this many minutes")
    ap.add_argument("--save-every", type=int, default=150, help="write the workbook every N days fetched")
    args = ap.parse_args()
    t0 = time.time()
    C.print_mapping({u: f for u, (f, _) in UNITS.items()}, "CNDC unit -> standard fuel")
    print(f"  agents left out (not generation): {sorted(NOT_GENERATION)}; agent-level: {AGENT_FUEL}", flush=True)

    daily = C.load_sheet(args.out, "Daily")
    detail = C.load_sheet(args.out, "Detail")
    end = dt.date.today() - dt.timedelta(days=1)
    days = sorted(C.days_to_fetch(daily, args.start, end, args.refresh_days), reverse=True)  # newest first
    print(f"{len(daily):,} days saved; fetching {len(days):,}", flush=True)

    s = requests.Session()
    s.headers.update({"User-Agent": C.USER_AGENT, "Referer": PAGE, "X-Requested-With": "XMLHttpRequest"})
    try:
        s.get(PAGE, timeout=(15, 60))  # session cookies, as the page does
    except requests.RequestException as e:
        print(f"  page: {e}", flush=True)
    per_day, failed = {}, []

    def flush():
        nonlocal daily, detail, per_day
        if per_day:
            new, new_detail = to_frames(per_day)
            daily, detail = C.merge(new, daily), C.merge(new_detail, detail)
            per_day = {}
        if not daily.empty:
            notes = [x for line in NOTES for x in (C.coverage_lines(daily) if line == "@COVERAGE" else [line])]
            C.write(args.out, daily, notes, detail)

    for k, day in enumerate(days, 1):
        if time.time() - t0 > args.budget_min * 60:
            print(f"time budget reached; {len(days) - k + 1} days left for later runs", flush=True)
            break
        try:
            got = fetch_day(s, day)
        except (requests.RequestException, ValueError) as e:
            print(f"  {day}: {type(e).__name__}: {e}", flush=True)
            got = None
            time.sleep(3)
        gen = sum(v for kk, v in (got or {}).items() if kk.split("|")[0] not in NOT_GENERATION and pd.notna(v))
        if not got or gen <= 0:
            failed.append(day)
        else:
            per_day[pd.Timestamp(day)] = got
        if k % 50 == 0:
            print(f"  {k}/{len(days)} ({day})", flush=True)
        if k % args.save_every == 0:
            flush()
    flush()
    if failed:
        print(f"{len(failed)} days without post-dispatch data: {[str(d) for d in failed[:20]]}", flush=True)
    if daily.empty:
        print("No data returned.", flush=True)
        sys.exit(1)
    C.print_monthly(daily)


if __name__ == "__main__":
    main()
