"""
Argentina daily power generation by type from CAMMESA - scheduled
entrypoint in the standard layout shared by the South America raw
grid-operator pulls (see power_daily_std.py).

Source: CAMMESA's "Parte control post-operativo" (nemo
PARTE_POST_OPERATIVO), one POyymmdd.zip per day on the same public
document API argentina_generation_mix.py uses
(https://api.cammesa.com/pub-svc/public/findDocumentosByNemoRango and
.../findAttachmentByNemoId). Each ZIP holds an Access database
(POyymmdd.mdb, read with mdbtools' mdb-export) with, for every
generating unit and hour, the ACTUAL energy produced (VALORES_GENERADORES.
ENERGIA, MWh), each unit's technology (GENERADORES.TIPO/SUBTIPO, plus an
INTERCAMBIO flag for import nodes) and the share of each fuel it burned
(COMBUSTIBLE_PORCENTAJE_DET) - so thermal output is split into gas, oil
(gasoil / fuel oil) and coal. Available from 2021 on (checked Jan-2021,
Jul-2022, Jun-2024, Sep-2026; AUB_POWER_DISCOVERY.py).

(argentina_generation_mix.py, the owner's local script, reads the
Programacion Diaria - the day-ahead dispatch programme - which only has a
CSV from Oct-2024 and no fuel split; it is unchanged.)

The ZIPs are ~2-10 MB and there is one per day, so the history is built
up over several runs: each run fetches missing days newest-first until
--budget-min minutes have passed, saving as it goes; the last 3 saved days
are re-fetched (CAMMESA re-issues some).

Usage: python3 ARGENTINA_POWER_DAILY.py [--out PATH] [--start YYYY-MM-DD] [--budget-min N]
Needs mdbtools (apt-get install mdbtools).
"""

print("STARTING", flush=True)

import argparse
import csv
import datetime as dt
import io
import os
import subprocess
import sys
import tempfile
import time
import zipfile
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import argentina_generation_mix as A  # noqa: E402  (shared CAMMESA endpoints / unit-type codes)
import power_daily_std as std  # noqa: E402

OUT = "output/Data and Chart Outputs/argentina_power_generation_daily.xlsx"
NEMO = "PARTE_POST_OPERATIVO"
RAW_SHEET = "By type and fuel (raw)"
CODES_SHEET = "Fuel codes"

# unit SUBTIPO -> fuel for non-thermal units (thermal ones are split by fuel code below)
SUBTYPE_FUEL = {"HI": "Hydro", "HR": "Hydro", "HB": "Hydro", "MH": "Hydro", "NU": "Nuclear", "EO": "Wind",
                "FV": "Solar", "BG": "Bioenergy", "BM": "Bioenergy"}
THERMAL = {"TV", "TG", "CC", "DI", "AG"}
# CAMMESA fuel codes (COMBUSTIBLE_PORCENTAJE_DET.COMB; descriptions/units from the mdb's CVP and
# COMBUSTIBLES_QUEMADOS_DET_TOTAL tables: gas is in dam3, gasoil in m3, fuel oil / coal / biomass in t)
FUEL_CODE = {
    # natural gas: own (propio), provided by CAMMESA (provisto), Plan Gas, agreement, imported LNG
    "GN": "Gas", "GA": "Gas", "GQ": "Gas", "GR": "Gas", "GX": "Gas", "GZ": "Gas", "GH": "Gas", "PI": "Gas",
    "GM": "Gas", "GU": "Gas", "LG": "Gas", "GD": "Gas",
    # gasoil (GO/GP/GY/GC 'GO provisto', GF/GG 'gas oil forzado') and fuel oil (FO/FP/FX), OX
    "GO": "Oil", "GP": "Oil", "GY": "Oil", "GC": "Oil", "GF": "Oil", "GG": "Oil", "FO": "Oil", "FP": "Oil",
    "FX": "Oil", "OX": "Oil", "OI": "Oil",
    "CM": "Coal",  # carbon mineral (San Nicolas, Rio Turbio)
    "BM": "Bioenergy", "BG": "Bioenergy", "BC": "Bioenergy", "BD": "Bioenergy",
    "U2": "Nuclear", "UA": "Nuclear", "UE": "Nuclear",  # uranium (nuclear units are counted by subtype NU anyway)
}
NO_FUEL = "--"  # thermal unit-hours with no fuel row: mostly the steam halves of combined cycles -> Gas

NOTES = [
    "UNITS",
    "MWh per day: the sum of CAMMESA's 24 hourly unit-level energy values (actual, post-operation).",
    "",
    "CATEGORY MAPPING (CAMMESA unit type / fuel -> sheet 'Daily')",
    "Hydro_MWh: unit subtypes HI (hydro), HR (renewable small hydro), HB (pumped storage, generation only - no "
    "negative/pumping values in the data), MH (mini hydro). Includes Yacyreta's whole output (YACYHI plus the "
    "Paraguayan share delivered to Argentina, YACYHIPY) and Argentina's half of Salto Grande (SGDEHIAR); "
    "Uruguay's half of Salto Grande is an import node and is excluded.",
    "Nuclear_MWh: NU (Atucha I/II, Embalse). Wind_MWh: EO. Solar_MWh: FV.",
    "Thermal units (TV steam, TG gas turbine, CC combined cycle, DI diesel/engines) are split hour by hour by "
    "the fuel shares CAMMESA reports per unit (COMBUSTIBLE_PORCENTAJE_DET):",
    "  Gas_MWh: GN, GA, GQ, GR, GX, GZ, GH, PI, GM, GU, LG (natural gas - own, CAMMESA-provided, Plan Gas, "
    "agreement, LNG; GD also gas, in dam3); plus thermal hours with no fuel reported (code '--', mostly the steam turbines of "
    "combined cycles, which burn no fuel of their own) - an assumption, see the '--' columns in the raw sheet.",
    "  Oil_MWh: GO, GP, GY, GC, OX (gasoil), GF, GG (gasoil 'forzado'), OI (gasoil industria), FO, FP, FX (fuel oil).",
    "  Coal_MWh: CM (carbon mineral).",
    "Bioenergy_MWh: BG (biogas) and BM (biomass) units, plus BM/BG/BC/BD fuel shares burned in thermal units.",
    "Other_MWh: any unit subtype or fuel code not listed above (none so far; new ones are printed by the pull).",
    "Excluded: units flagged INTERCAMBIO='S' - import nodes (Brazil via Garabi, Uruguay's Salto Grande half "
    "and thermal plants, Paraguay, Bolivia, Chile). Their MWh are kept in the raw sheet as 'IMPORT|...'.",
    f"Sheet '{RAW_SHEET}': daily MWh per 'SUBTIPO|FUEL CODE' as read, so the mapping can change without "
    "re-downloading. Sheet 'Fuel codes': each code's sub-fuel and unit as CAMMESA reports them and where it is "
    "mapped; a code not in the list above is placed by its unit (dam3 -> Gas, m3 -> Oil), else Other.",
    "",
    "SOURCE",
    "CAMMESA 'Parte control post-operativo' (PARTE_POST_OPERATIVO), one POyymmdd.zip per day with an Access "
    "database: https://api.cammesa.com/pub-svc/public/findDocumentosByNemoRango (nemo=PARTE_POST_OPERATIVO) "
    "and .../findAttachmentByNemoId. Actual post-operation values (the owner's argentina_generation_mix.py "
    "uses the day-ahead Programacion Diaria instead). Published ~1-2 days after the day.",
    "History from 2021-01-01. One ~2-10 MB download per day, so the archive is filled newest-first over "
    "several runs (each run has a time budget); afterwards each daily run (argentina_power_generation_daily."
    "yml) fetches only missing days plus the last 3.",
]


# ------------------------------------------------------------------ fetch
def month_docs(s, first):
    """{attachment id 'POyymmdd.zip': (doc, attachment)} for one month (one API call)."""
    nxt = (pd.Timestamp(first) + pd.offsets.MonthBegin(1)).date()
    r = s.get(A.LOOKUP_URL, params={"fechadesde": first.strftime(A.TIME_FMT), "fechahasta": nxt.strftime(A.TIME_FMT),
                                    "nemo": NEMO}, timeout=60)
    r.raise_for_status()
    docs = r.json()
    out = {}
    for doc in docs if isinstance(docs, list) else []:
        for att in doc.get("adjuntos", []):
            aid = att.get("id", "")
            if aid.startswith("PO") and aid.endswith(".zip"):
                out[aid] = (doc, att)  # later docs (re-issues) win
    return out


def mdb_rows(path, table):
    out = subprocess.run(["mdb-export", path, table], capture_output=True, text=True, timeout=300)
    if out.returncode != 0:
        raise RuntimeError(f"mdb-export {table}: {out.stderr[:200]}")
    return list(csv.DictReader(io.StringIO(out.stdout)))


def parse_day(zip_bytes):
    """({'SUBTIPO|FUEL' or 'IMPORT|SUBTIPO': MWh}, {fuel code: (sub-fuel, unit)}) for one day's
    post-operative database."""
    zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
    name = next(n for n in zf.namelist() if n.lower().endswith(".mdb"))
    with tempfile.NamedTemporaryFile(suffix=".mdb", delete=False) as f:
        f.write(zf.read(name))
        path = f.name
    try:
        gens = {g["GRUPO"]: g for g in mdb_rows(path, "GENERADORES")}
        energy = defaultdict(float)
        for v in mdb_rows(path, "VALORES_GENERADORES"):
            try:
                energy[(v["GRUPO"], v["HORA"])] += float(v["ENERGIA"] or 0)
            except ValueError:
                continue
        shares = defaultdict(dict)
        for c in mdb_rows(path, "COMBUSTIBLE_PORCENTAJE_DET"):
            try:
                pct = float(c["PORCENTAJE"] or 0)
            except ValueError:
                continue
            if pct > 0:
                key = (c["GRUPO"], c["HORA"])
                shares[key][c["COMB"]] = shares[key].get(c["COMB"], 0) + pct
        codes = {}
        try:
            for c in mdb_rows(path, "COMBUSTIBLES_QUEMADOS_DET_TOTAL"):
                codes[c.get("COMB")] = (c.get("SUBCOMB") or "", c.get("UNIDAD") or "")
        except (RuntimeError, subprocess.SubprocessError):
            pass
    finally:
        os.unlink(path)
    out = defaultdict(float)
    hours = set()
    for (unit, hour), e in energy.items():
        hours.add(hour)
        g = gens.get(unit, {})
        sub = g.get("SUBTIPO") or "?"
        if g.get("INTERCAMBIO") == "S":
            out[f"IMPORT|{sub}"] += e
        elif sub in THERMAL or sub in ("BG", "BM"):
            fuel = shares.get((unit, hour))
            if fuel:
                tot = sum(fuel.values())
                for code, pct in fuel.items():
                    out[f"{sub}|{code}"] += e * pct / tot
            else:
                out[f"{sub}|{NO_FUEL}"] += e
        else:
            out[f"{sub}|"] += e
    if len(hours) < 23:
        raise RuntimeError(f"only {len(hours)} hours in VALORES_GENERADORES")
    return dict(out), codes


# ------------------------------------------------------------------ shape
UNIT_FUEL = {"DAM3": "Gas", "M3": "Oil"}


def raw_to_fuels(raw, codes=None):
    """Raw 'SUBTIPO|CODE' columns -> standard fuel columns. A fuel code not in FUEL_CODE is placed by the
    unit CAMMESA measures it in (dam3 -> Gas, m3 -> Oil; 'Fuel codes' sheet), else Other."""
    units = {} if codes is None or codes.empty else codes["unit"].astype(str).str.upper().to_dict()
    fuels = defaultdict(lambda: pd.Series(0.0, index=raw.index))
    unknown = set()
    for col in raw.columns:
        sub, code = col.split("|", 1)
        if sub == "IMPORT":
            continue
        if sub in SUBTYPE_FUEL and (not code or sub in ("BG", "BM")):
            fuel = SUBTYPE_FUEL[sub]
        elif code == NO_FUEL:
            fuel = "Gas"
        elif code in FUEL_CODE:
            fuel = FUEL_CODE[code]
        elif units.get(code) in UNIT_FUEL:
            fuel = UNIT_FUEL[units[code]]
            print(f"  note: fuel code {code} not in FUEL_CODE, placed by its unit ({units[code]}) -> {fuel}", flush=True)
        else:
            fuel = "Other"
            unknown.add(col)
        fuels[fuel] = fuels[fuel] + raw[col].fillna(0)
    if unknown:
        print(f"  WARNING: unmapped types/fuel codes (-> Other): {sorted(unknown)}", flush=True)
    out = pd.DataFrame(dict(fuels), index=raw.index)
    out.loc[raw.isna().all(axis=1)] = float("nan")
    return std.standardise(out)


def save(path, raw, codes):
    raw = raw.sort_index()
    raw.index.name = "date"
    raw = raw[sorted(raw.columns)]
    daily = raw_to_fuels(raw, codes)
    std.write(path, daily, NOTES, {RAW_SHEET: raw.round(1), CODES_SHEET: codes.sort_index()})


def load_codes(path):
    try:
        codes = pd.read_excel(path, sheet_name=CODES_SHEET, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        codes = pd.DataFrame(columns=["sub_fuel", "unit", "mapped_to"])
    codes.index = codes.index.astype(str)
    codes.index.name = "code"
    return codes


def add_codes(codes, found):
    for code, (sub, unit) in found.items():
        if not code:
            continue
        codes.loc[code, "sub_fuel"] = sub
        codes.loc[code, "unit"] = unit
        codes.loc[code, "mapped_to"] = FUEL_CODE.get(code) or UNIT_FUEL.get(str(unit).upper(), "Other")
    return codes


def fetch_one(s, day, hit, tries=3):
    """Download and parse one day; a failed or truncated download (bad zip / zlib error) is retried."""
    doc, att = hit
    for attempt in range(tries):
        try:
            r = s.get(A.ATTACHMENT_URL, params={"attachmentId": att["id"], "docId": doc["id"],
                                                "nemo": doc.get("nemo") or NEMO}, timeout=180)
            r.raise_for_status()
            out, codes = parse_day(r.content)
            return out, codes, len(r.content)
        except Exception:  # noqa: BLE001 - network, HTTP, zip/zlib and mdb errors alike
            if attempt == tries - 1:
                raise
            time.sleep(5 * (attempt + 1))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start", type=dt.date.fromisoformat, default=std.HISTORY_START)
    ap.add_argument("--budget-min", type=float, default=40, help="stop fetching after this many minutes")
    ap.add_argument("--save-every", type=int, default=24, help="write the workbook every N fetched days")
    ap.add_argument("--workers", type=int, default=3, help="parallel downloads")
    args = ap.parse_args()
    t0 = time.time()

    raw = std.load_sheet(args.out, RAW_SHEET)
    daily_saved = std.load_sheet(args.out, "Daily")
    end = dt.date.today() - dt.timedelta(days=1)
    todo = sorted(std.missing_days(daily_saved, args.start, end), reverse=True)  # newest first
    print(f"{len(raw):,} days saved; {len(todo):,} to fetch (budget {args.budget_min:.0f} min)", flush=True)

    codes = load_codes(args.out)
    s = A.make_session()
    listings, rows, done = {}, {}, 0
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        for i in range(0, len(todo), args.save_every):
            if time.time() - t0 > args.budget_min * 60:
                print(f"time budget reached; {len(todo) - done} days left for later runs", flush=True)
                break
            batch = todo[i:i + args.save_every]
            jobs = {}
            for day in batch:
                first = day.replace(day=1)
                try:
                    if first not in listings:
                        listings[first] = month_docs(s, first)
                except (requests.RequestException, ValueError) as e:
                    print(f"  listing {first:%Y-%m}: FAILED ({type(e).__name__}: {e})", flush=True)
                    listings[first] = {}
                hit = listings[first].get(day.strftime("PO%y%m%d.zip"))
                if hit is None:
                    print(f"  {day}: not published (yet)", flush=True)
                    continue
                jobs[pool.submit(fetch_one, s, day, hit)] = day
            for fut in as_completed(jobs):
                day = jobs[fut]
                try:
                    out, found, size = fut.result()
                    rows[pd.Timestamp(day)] = out
                    codes = add_codes(codes, found)
                    tot = sum(v for k, v in out.items() if not k.startswith("IMPORT"))
                    print(f"  {day}: {tot:,.0f} MWh ({size / 1e6:.1f} MB, {time.time() - t0:.0f}s)", flush=True)
                except Exception as e:  # noqa: BLE001 - one bad day must not stop the backfill
                    print(f"  {day}: FAILED ({type(e).__name__}: {str(e)[:200]}) - retried next run", flush=True)
            done += len(batch)
            if rows:
                raw = std.merge(pd.DataFrame.from_dict(rows, orient="index"), raw)
                rows = {}
                save(args.out, raw, codes)
    if raw.empty:
        print("No data returned.", flush=True)
        sys.exit(1)
    save(args.out, raw, codes)


if __name__ == "__main__":
    main()
