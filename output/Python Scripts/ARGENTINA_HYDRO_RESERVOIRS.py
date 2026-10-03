"""
Argentina hydro reservoirs: daily lake levels (m above sea level) of the
main hydro reservoirs and the river flows that feed them, in one workbook
(sheet "Daily"), for the AGSI-style water-year charts (add_charts.py).

Sources (all public, no key; found with
discovery_archive/south_america/ARGENTINA_HYDRO_DISCOVERY.py):
  1. CAMMESA "Cotas Diarias" - https://cammesaweb.cammesa.com/download/cotas-diarias/
     one xlsx (sheet COTAS) with the daily level of each big hydro plant's
     reservoir from 2023-01-01 to the end of the last closed month:
     ALICHI Alicura, CHOCHI El Chocon, PAGUHI Piedra del Aguila, PBANHI
     Cerros Colorados (Planicie Banderita - the level is Los Barreales lake),
     PPLEHI Pichi Picun Leufu, FUTAHI Futaleufu, YACYHI Yacyreta, SGDEHIAR
     Salto Grande. Re-read every run (it is one ~0.6 MB file); days it no
     longer carries are kept from the workbook.
  2. CAMMESA "Caudales Diarios" - https://cammesaweb.cammesa.com/download/caudales-diarios/
     daily mean flows (m3/s) of the Limay, Neuquen, Collon Cura and
     Futaleufu rivers, same span.
  3. CAMMESA weekly programme (nemo PROGRAMACION_SEMANAL on
     api.cammesa.com/pub-svc, psemWWYY.zip -> Access .MDB, table COTAS):
     'CotaIni', the level CAMMESA expects at the start of the programmed
     week (FECHA.FInicio, a Monday), set 3-4 days earlier. One value per
     week from 2005 on - used only for weeks with no daily value
     (before 2023, and the weeks after the daily file ends). Checked
     against the daily file: usually within 0.1 m (El Chocon, Cerros
     Colorados, Futaleufu, Piedra del Aguila), up to ~1 m at Alicura.
     Read with mdbtools.
  4. AIC (Autoridad Interjurisdiccional de las Cuencas de los rios Limay,
     Neuquen y Negro) - https://www.aic.gob.ar/sitio/embalses: each lake's
     page shows today's 'Nivel Actual'. AIC keeps no public history, so a
     small DAILY job (--aic-only, .github/workflows/argentina_aic_snapshot.yml)
     appends that day's reading to south_america/argentina_aic_snapshots.csv;
     the full pull (1st and 15th of the month) merges that file into the
     sheet "AIC snapshots" and uses it for days with no CAMMESA daily value.
  5. INA (Instituto Nacional del Agua) a5 database - https://alerta.ina.gob.ar/a5:
     daily mean flow of the Parana entering Yacyreta (series 26684, from
     2006) and of the Uruguay at Salto Grande (26674), and the daily mean
     level at Salto Grande Arriba (26319, the lake side of the dam) - long,
     current series for the two run-of-river binational plants.

Stored volumes (hm3) / % of useful volume are not published by any of
these sources (AIC and CAMMESA give levels only), so the workbook has
levels and flows and no volume-weighted Comahue %.

Incremental: the workbook is the archive. Weekly programmes already read
are listed in "Weekly programme" and not fetched again (history is filled
newest-first within --budget-min minutes per run); INA is re-read from 30
days before its last stored day; AIC adds today's reading.

Usage: python3 ARGENTINA_HYDRO_RESERVOIRS.py --out PATH [--start YYYY-MM-DD] [--budget-min N]
       python3 ARGENTINA_HYDRO_RESERVOIRS.py --aic-only      (daily: just store today's AIC reading)
"""

print("STARTING", flush=True)

import argparse
import datetime as dt
import io
import os
import re
import subprocess
import sys
import tempfile
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pandas as pd
import requests

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

OUT = "output/Data and Chart Outputs/argentina_hydro_reservoirs.xlsx"
ap = argparse.ArgumentParser(description="Argentina hydro reservoir levels and river flows")
ap.add_argument("--out", default=OUT)
ap.add_argument("--start", default="2005-01-01", help="first week of the weekly-programme backfill")
ap.add_argument("--budget-min", type=float, default=15, help="minutes to spend on the weekly-programme backfill")
ap.add_argument("--aic-only", action="store_true", help="only append today's AIC reading to AIC_CSV and stop")
ARGS = ap.parse_args()
AIC_CSV = str(Path(__file__).resolve().parent / "argentina_aic_snapshots.csv")   # daily AIC readings (no history at AIC)
T0 = time.time()

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"}
CAMMESA_UA = {"User-Agent": "gas-demand-scripts/1.0"}
COTAS_PAGE = "https://cammesaweb.cammesa.com/download/cotas-diarias/"
CAUDALES_PAGE = "https://cammesaweb.cammesa.com/download/caudales-diarios/"
LOOKUP_URL = "https://api.cammesa.com/pub-svc/public/findDocumentosByNemoRango"
ATTACHMENT_URL = "https://api.cammesa.com/pub-svc/public/findAttachmentByNemoId"
PSEM_NEMO = "PROGRAMACION_SEMANAL"
TIME_FMT = "%Y-%m-%dT%H:%M:%S.000Z"
AIC_LIST = "https://www.aic.gob.ar/sitio/embalses"
INA_OBS = "https://alerta.ina.gob.ar/a5/obs/puntual/series/{sid}/observaciones"
TODAY = (pd.Timestamp.now(tz="America/Argentina/Buenos_Aires")).tz_localize(None).normalize()

# reservoir -> (column stem, CAMMESA daily-file code, weekly-programme code, AIC lake name, height band m:
#   AIC's minimum extraordinary level less ~2 m to its maximum level plus ~1 m; Futaleufu from CAMMESA's data)
RESERVOIRS = [
    ("Chocon", "CHOCHI", "CHOCHI", "El Chocón", (365, 383)),
    ("PiedraAguila", "PAGUHI", "PDAGHI", "Piedra del Águila", (560, 593)),
    ("Alicura", "ALICHI", "ALICHI", "Alicurá", (690, 706)),
    ("CerrosColorados", "PBANHI", "PBANHI", "Los Barreales", (408, 423)),
    ("PichiPicun", "PPLEHI", "PICUHI", "Pichi Picún Leufú", (476, 480.5)),
    ("Futaleufu", "FUTAHI", "FUTAHI", None, (455, 500)),
    ("Yacyreta", "YACYHI", None, None, (75, 90)),
    ("SaltoGrande", "SGDEHIAR", None, None, (25, 42)),
]
COMAHUE = ["Chocon", "PiedraAguila", "Alicura", "CerrosColorados", "PichiPicun"]
# CAMMESA river name (normalised) -> column
RIVERS = {"limay": "Limay_m3s", "neuquen": "Neuquen_m3s", "c. cura": "CollonCura_m3s", "collon cura": "CollonCura_m3s",
          "futaleufu": "Futaleufu_m3s", "parana": "Parana_m3s_CAMMESA", "uruguay": "Uruguay_m3s_CAMMESA"}
# INA a5 series id -> column
INA_SERIES = {26684: "Parana_Yacyreta_m3s", 26674: "Uruguay_SaltoGrande_m3s", 26319: "SaltoGrandeLevel_m_INA"}

S = requests.Session()
S.headers.update(UA)


def norm(text):
    t = str(text).strip().lower()
    for a, b in zip("áéíóúñ", "aeioun"):
        t = t.replace(a, b)
    return t


def elapsed_min():
    return (time.time() - T0) / 60


# ------------------------------------------------------------------ archive

def load_sheet(name, index_col=0):
    try:
        df = pd.read_excel(ARGS.out, sheet_name=name, index_col=index_col)
    except (FileNotFoundError, ValueError, KeyError):
        return pd.DataFrame()
    if index_col is not None:
        df.index = pd.to_datetime(df.index, errors="coerce")
        df = df[df.index.notna()].sort_index()
    return df


# ------------------------------------------------------------------ 1+2. CAMMESA daily files

def download_wpdm(page):
    """CAMMESA's WordPress Download Manager page -> the file behind its 'Descargar' button."""
    r = S.get(page, timeout=60)
    r.raise_for_status()
    links = re.findall(r'(https?://[^"\'\s<>]*\?wpdmdl=\d+[^"\'\s<>]*)', r.text)
    if not links:
        raise RuntimeError(f"no download link on {page}")
    f = S.get(links[0].replace("&amp;", "&"), timeout=180)
    f.raise_for_status()
    if f.content[:2] != b"PK":
        raise RuntimeError(f"{page}: download is not an xlsx ({f.headers.get('Content-Type')})")
    return f.content


def detail_table(content, key_header):
    """The 'DETALLE ... DIARIOS' block of a CAMMESA cotas/caudales sheet: rows below the
    header row that starts with AÑO, columns picked by header name (FECHA, key, value)."""
    raw = pd.read_excel(io.BytesIO(content), header=None)
    first = raw.iloc[:, 0].astype(str).map(norm)
    hdr = raw.index[first.isin(["ano", "año"])]
    if not len(hdr):
        raise RuntimeError("no AÑO header row in the CAMMESA file")
    hdr = hdr[0]
    names = [norm(x) for x in raw.iloc[hdr].tolist()]
    body = raw.iloc[hdr + 1:]

    def col(*want):
        for w in want:
            if w in names:
                return names.index(w)
        raise RuntimeError(f"no column {want} in {names}")

    out = pd.DataFrame({"date": body.iloc[:, col("fecha")], "key": body.iloc[:, col(*key_header)],
                        "value": pd.to_numeric(body.iloc[:, col("cota", "caudal", "valor")], errors="coerce")})
    # FECHA is a real Excel date; anything else (a string) is read day-first, as CAMMESA writes it
    dates = out["date"].map(lambda v: pd.Timestamp(v) if isinstance(v, (dt.datetime, pd.Timestamp))
                            else pd.to_datetime(str(v), dayfirst=True, errors="coerce"))
    out["date"] = pd.DatetimeIndex(dates).normalize()
    out["key"] = out["key"].astype(str).str.strip()
    return out.dropna(subset=["date", "value"])


PLANT_RANGE = {code: rng for _, code, _, _, rng in RESERVOIRS}
PLANT_RANGE["RGDEHB"] = (860, 890)        # Rio Grande pumped storage (Cordoba), in the file but not charted
OVERLAP = (476, 480.5)                    # Pichi Picun Leufu's whole range; Futaleufu sometimes sits in it too
REPAIRED_DAYS = []


def unscramble(t):
    """CAMMESA's file sometimes puts a day's levels against the wrong plant codes (from 2026-05-03:
    some values repeated on other rows, some plants missing). Each lake lives in its own height band,
    so on such a day each value is given to the one plant whose band holds it; Futaleufu vs Pichi Picun
    (overlapping bands) is settled by closeness to each plant's last good value, or left blank."""
    rows, last = [], {}
    for date, g in t.sort_values("date").groupby("date"):
        vals = dict(zip(g["key"], g["value"]))
        known = {k: v for k, v in vals.items() if k in PLANT_RANGE}
        clean_day = (all(PLANT_RANGE[k][0] <= v <= PLANT_RANGE[k][1] for k, v in known.items())
                     and not pd.Series(list(known.values())).round(3).duplicated().any())
        if clean_day:
            assigned = known
        else:
            REPAIRED_DAYS.append(date)
            distinct = sorted(set(round(v, 3) for v in known.values()))
            assigned = {}
            for code, (lo, hi) in PLANT_RANGE.items():
                if code in ("FUTAHI", "PPLEHI"):
                    continue
                cands = [v for v in distinct if lo <= v <= hi]
                if len(cands) == 1:
                    assigned[code] = cands[0]
            fut = {"FUTAHI": [], "PPLEHI": []}
            for v in [v for v in distinct if PLANT_RANGE["FUTAHI"][0] <= v <= PLANT_RANGE["FUTAHI"][1]]:
                if not (OVERLAP[0] <= v <= OVERLAP[1]):
                    fut["FUTAHI"].append(v)
                    continue
                near_f = "FUTAHI" in last and abs(v - last["FUTAHI"]) < 1.0
                near_p = "PPLEHI" in last and abs(v - last["PPLEHI"]) < 1.0
                if near_f != near_p:
                    fut["FUTAHI" if near_f else "PPLEHI"].append(v)
            for code, cands in fut.items():
                if len(cands) == 1:
                    assigned[code] = cands[0]
        last.update(assigned)
        rows += [(date, k, v) for k, v in assigned.items()]
    return pd.DataFrame(rows, columns=["date", "key", "value"])


def cammesa_levels():
    t = detail_table(download_wpdm(COTAS_PAGE), ("central",))
    dup = t.duplicated(["date", "key"], keep=False)
    if dup.any():
        print(f"  cotas: {dup.sum()} duplicate date/plant rows - last one kept", flush=True)
    t = unscramble(t.drop_duplicates(["date", "key"], keep="last"))
    if REPAIRED_DAYS:
        print(f"  cotas: {len(REPAIRED_DAYS)} day(s) with levels against the wrong plant codes, reassigned by height band "
              f"({min(REPAIRED_DAYS):%Y-%m-%d}..{max(REPAIRED_DAYS):%Y-%m-%d})", flush=True)
    piv = t.pivot(index="date", columns="key", values="value")
    print(f"  CAMMESA cotas diarias: {len(piv):,} days {piv.index.min():%Y-%m-%d}..{piv.index.max():%Y-%m-%d}, "
          f"plants {list(piv.columns)}", flush=True)
    return piv


def cammesa_flows():
    t = detail_table(download_wpdm(CAUDALES_PAGE), ("rio", "río"))
    t["col"] = t["key"].map(lambda k: RIVERS.get(norm(k)))
    unknown = sorted(set(t.loc[t["col"].isna(), "key"]))
    if unknown:
        print(f"  caudales: rivers not mapped (kept out): {unknown}", flush=True)
    t = t.dropna(subset=["col"]).drop_duplicates(["date", "col"], keep="last")
    piv = t.pivot(index="date", columns="col", values="value")
    print(f"  CAMMESA caudales diarios: {len(piv):,} days {piv.index.min():%Y-%m-%d}..{piv.index.max():%Y-%m-%d}, "
          f"rivers {list(piv.columns)}", flush=True)
    return piv


# ------------------------------------------------------------------ 3. weekly programme

def psem_docs(start, end):
    """Weekly-programme documents published between start and end (a quarter per request, 8 in parallel:
    each lookup takes ~10 s)."""
    spans, cur = [], start
    while cur <= end:
        nxt = min(cur + dt.timedelta(days=91), end + dt.timedelta(days=1))
        spans.append((cur, nxt))
        cur = nxt

    def lookup(span):
        a, b = span
        params = {"fechadesde": a.strftime(TIME_FMT), "fechahasta": b.strftime(TIME_FMT), "nemo": PSEM_NEMO}
        for attempt in range(3):
            try:
                r = requests.get(LOOKUP_URL, params=params, headers=CAMMESA_UA, timeout=90)
                r.raise_for_status()
                return r.json()
            except (requests.RequestException, ValueError) as e:
                if attempt == 2:
                    print(f"  psem list {a}..{b}: FAILED {type(e).__name__}: {e}", flush=True)
                time.sleep(3)
        return []

    with ThreadPoolExecutor(max_workers=8) as pool:
        docs = [d for part in pool.map(lookup, spans) for d in part]
    out = {}
    for d in docs:
        for a in d.get("adjuntos", []):
            name = str(a.get("id", ""))
            if re.fullmatch(r"psem\d{4}\.zip", name, re.I):
                out[name.lower()] = (d, a)   # re-issues of a week keep the latest document
    return out


def mdb_table(path, table):
    out = subprocess.run(["mdb-export", path, table], capture_output=True, text=True, timeout=120)
    return pd.read_csv(io.StringIO(out.stdout)) if out.stdout.strip() else pd.DataFrame()


def read_psem(doc, att):
    r = requests.get(ATTACHMENT_URL, params={"attachmentId": att["id"], "docId": doc["id"], "nemo": doc.get("nemo", PSEM_NEMO)},
                     headers=CAMMESA_UA, timeout=180)
    r.raise_for_status()
    zf = zipfile.ZipFile(io.BytesIO(r.content))
    mdb = next((n for n in zf.namelist() if n.lower().endswith(".mdb")), None)
    if mdb is None:
        return None
    with tempfile.NamedTemporaryFile(suffix=".mdb", delete=False) as f:
        f.write(zf.read(mdb))
        path = f.name
    try:
        fecha = mdb_table(path, "FECHA")
        cotas = mdb_table(path, "COTAS")
    finally:
        os.remove(path)
    if fecha.empty or cotas.empty:
        return None
    week_start = pd.to_datetime(fecha.iloc[0]["FInicio"], format="%m/%d/%y %H:%M:%S", errors="coerce")
    if pd.isna(week_start):
        return None
    row = {"week_start": week_start.normalize(), "published": doc.get("fecha")}
    for _, c in cotas.iterrows():
        row[str(c["CentHidr"]).strip()] = pd.to_numeric(c["CotaIni"], errors="coerce")
    return row


def update_weekly(archive):
    """archive: DataFrame indexed by file name. Fetch weeks not yet read: newest first, within the budget."""
    start = pd.Timestamp(ARGS.start).date()
    have = set(archive.index.str.lower()) if not archive.empty else set()
    # the recent end every run (new programmes), the full span only while history is incomplete
    recent_from = (TODAY - pd.Timedelta(days=60)).date()
    docs = psem_docs(recent_from, TODAY.date() + dt.timedelta(days=7))
    oldest = pd.to_datetime(archive["published"], dayfirst=True, errors="coerce").min() if have else pd.NaT
    # whole span while the backfill is unfinished, and on the 1st of each month to retry weeks that failed
    if pd.isna(oldest) or oldest > pd.Timestamp(start) + pd.Timedelta(days=365) or TODAY.day == 1:
        docs = {**psem_docs(start, recent_from), **docs}
    todo = sorted((k for k in docs if k not in have),
                  key=lambda k: pd.to_datetime(docs[k][0].get("fecha"), dayfirst=True, errors="coerce"), reverse=True)
    print(f"  weekly programme: {len(docs)} listed, {len(have)} already read, {len(todo)} to fetch", flush=True)
    rows = []

    def fetch(k):
        doc, att = docs[k]
        if elapsed_min() > ARGS.budget_min:
            return None
        try:
            row = read_psem(doc, att)
        except Exception as e:  # noqa: BLE001
            print(f"  {k}: FAILED {type(e).__name__}: {e}", flush=True)
            return None
        return {"file": k, **(row or {"published": doc.get("fecha")})}   # an unreadable week is remembered too

    with ThreadPoolExecutor(max_workers=5) as pool:   # ~10 s per programme; more in parallel draws connect timeouts
        for res in pool.map(fetch, todo):
            if res is not None:
                rows.append(res)
                if len(rows) % 100 == 0:
                    print(f"  ... {len(rows)} weekly programmes read ({elapsed_min():.1f} min)", flush=True)
    if len(rows) < len(todo):
        print(f"  {len(todo) - len(rows)} weekly programmes left for the next run (time budget / failures)", flush=True)
    if rows:
        new = pd.DataFrame(rows).set_index("file")
        archive = new if archive.empty else pd.concat([archive[~archive.index.isin(new.index)], new])
    if not archive.empty:
        archive["week_start"] = pd.to_datetime(archive["week_start"], errors="coerce")
        archive = archive.sort_values("week_start")
    return archive


# ------------------------------------------------------------------ 4. AIC

def aic_snapshot():
    r = S.get(AIC_LIST, timeout=60)
    r.raise_for_status()
    out = {}
    for href in sorted(set(re.findall(r'href=["\']([^"\']*embalses-detalle\?a=\d+[^"\'#]*)', r.text))):
        d = S.get(requests.compat.urljoin(r.url, href.replace("&amp;", "&")), timeout=60)
        text = re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", d.text, flags=re.S | re.I))
        name = re.search(r"Caudales Programados Embalses (.+?) Ubicaci", text)
        level = re.search(r"Nivel Actual ([\d.,]+) msnm", text)
        if name and level:
            out[name.group(1).strip()] = float(level.group(1).replace(",", "."))
    print(f"  AIC Nivel Actual ({TODAY:%Y-%m-%d}): {out}", flush=True)
    return out


def load_aic_csv():
    try:
        df = pd.read_csv(AIC_CSV, index_col=0, parse_dates=True)
    except FileNotFoundError:
        return pd.DataFrame()
    return df[df.index.notna()].sort_index()


def store_aic(snap):
    """Append (or replace) today's AIC reading in AIC_CSV and return the whole file."""
    arc = load_aic_csv()
    if snap:
        row = pd.DataFrame([snap], index=pd.DatetimeIndex([TODAY], name="date"))
        arc = row if arc.empty else row.combine_first(arc)
        arc.index.name = "date"
        arc.sort_index().to_csv(AIC_CSV, date_format="%Y-%m-%d")
    return arc


# ------------------------------------------------------------------ 5. INA

def ina_series(sid, since):
    url = INA_OBS.format(sid=sid)
    r = S.get(url, params={"timestart": since.strftime("%Y-%m-%d"), "timeend": (TODAY + pd.Timedelta(days=1)).strftime("%Y-%m-%d"),
                           "format": "json"}, timeout=300)
    r.raise_for_status()
    obs = r.json()
    s = pd.Series({pd.Timestamp(o["timestart"]).tz_convert("America/Argentina/Buenos_Aires").tz_localize(None).normalize():
                   pd.to_numeric(o["valor"], errors="coerce") for o in obs}, dtype=float)
    return s.dropna().sort_index()


# ------------------------------------------------------------------ main

def main():
    levels_arc = load_sheet("CAMMESA daily levels")
    flows_arc = load_sheet("CAMMESA daily flows")
    aic_arc = load_sheet("AIC snapshots")
    ina_arc = load_sheet("INA")
    try:   # indexed by programme file name (psemWWYY.zip), not by date
        weekly_arc = pd.read_excel(ARGS.out, sheet_name="Weekly programme", index_col=0)
        weekly_arc.index = weekly_arc.index.astype(str)
    except (FileNotFoundError, ValueError, KeyError):
        weekly_arc = pd.DataFrame()

    print("CAMMESA daily files...", flush=True)
    try:
        levels = cammesa_levels()
        levels_arc = levels if levels_arc.empty else levels.combine_first(levels_arc)
    except Exception as e:  # noqa: BLE001
        print(f"  cotas FAILED ({type(e).__name__}: {e}); keeping the stored days", flush=True)
    try:
        flows = cammesa_flows()
        flows_arc = flows if flows_arc.empty else flows.combine_first(flows_arc)
    except Exception as e:  # noqa: BLE001
        print(f"  caudales FAILED ({type(e).__name__}: {e}); keeping the stored days", flush=True)

    print("INA a5...", flush=True)
    ina_new = {}
    for sid, col in INA_SERIES.items():
        since = (ina_arc[col].dropna().index.max() - pd.Timedelta(days=30)) if col in ina_arc and ina_arc[col].notna().any() \
            else pd.Timestamp("1995-01-01")
        try:
            ina_new[col] = ina_series(sid, since)
            print(f"  {col} (series {sid}): {len(ina_new[col]):,} days from {since:%Y-%m-%d}", flush=True)
        except Exception as e:  # noqa: BLE001
            print(f"  {col} (series {sid}) FAILED ({type(e).__name__}: {e})", flush=True)
    if ina_new:
        fresh = pd.DataFrame(ina_new)
        ina_arc = fresh if ina_arc.empty else fresh.combine_first(ina_arc)

    print("AIC...", flush=True)
    try:
        snap = aic_snapshot()
    except Exception as e:  # noqa: BLE001
        snap = {}
        print(f"  AIC FAILED ({type(e).__name__}: {e})", flush=True)
    daily_aic = store_aic(snap)   # the daily job's readings since the last full pull, plus today's
    if not daily_aic.empty:
        aic_arc = daily_aic if aic_arc.empty else daily_aic.combine_first(aic_arc)

    print("CAMMESA weekly programmes...", flush=True)
    try:
        weekly_arc = update_weekly(weekly_arc)
    except Exception as e:  # noqa: BLE001
        print(f"  weekly programme FAILED ({type(e).__name__}: {e}); keeping the stored weeks", flush=True)

    daily = build_daily(levels_arc, flows_arc, weekly_arc, aic_arc, ina_arc)
    write(daily, levels_arc, flows_arc, weekly_arc, aic_arc, ina_arc)


def clean(series, rng, label):
    lo, hi = rng
    bad = series[(series < lo) | (series > hi)]
    if len(bad):
        print(f"  {label}: {len(bad)} value(s) outside {lo}-{hi} m dropped, e.g. {bad.head(3).round(2).to_dict()}", flush=True)
    return series[(series >= lo) & (series <= hi)]


def build_daily(levels_arc, flows_arc, weekly_arc, aic_arc, ina_arc):
    cols, source = {}, {}
    weekly = pd.DataFrame()
    if not weekly_arc.empty and "week_start" in weekly_arc:
        weekly = weekly_arc.dropna(subset=["week_start"]).copy()
        # a few programmes carry a wrong FInicio (wrong year or month, e.g. psem0126.zip published
        # 23/12/2025 says 29/12/2026): then use the Monday after the publication date
        start = pd.to_datetime(weekly["week_start"], errors="coerce")
        pub = pd.to_datetime(weekly["published"], dayfirst=True, errors="coerce")
        lag = (start - pub).dt.days
        next_monday = pub + pd.to_timedelta((7 - pub.dt.weekday) % 7 + 7 * (pub.dt.weekday == 0), unit="D")
        weekly["week_start"] = start.where(lag.between(0, 7) | pub.isna(), next_monday)
        bad = int((~lag.between(0, 7) & pub.notna()).sum())
        if bad:
            print(f"  weekly programme: {bad} programme(s) with an implausible FInicio placed on the Monday after "
                  "publication", flush=True)
        weekly = weekly.dropna(subset=["week_start"]).set_index("week_start").sort_index()
        weekly = weekly[~weekly.index.duplicated(keep="last")]
        weekly = weekly[weekly.index <= TODAY]          # programmes for a week still ahead are forecasts
    for stem, code, wcode, aic_name, rng in RESERVOIRS:
        parts = []
        if code in levels_arc:
            parts.append(("CAMMESA daily", clean(levels_arc[code].dropna(), rng, f"{stem} CAMMESA daily")))
        if stem == "SaltoGrande" and "SaltoGrandeLevel_m_INA" in ina_arc:
            parts.append(("INA daily", clean(ina_arc["SaltoGrandeLevel_m_INA"].dropna(), rng, f"{stem} INA")))
        if aic_name and aic_name in aic_arc:
            parts.append(("AIC", clean(aic_arc[aic_name].dropna(), rng, f"{stem} AIC")))
        if wcode and wcode in weekly:
            parts.append(("CAMMESA weekly programme", clean(pd.to_numeric(weekly[wcode], errors="coerce").dropna(), rng,
                                                          f"{stem} weekly")))
        s, src = pd.Series(dtype=float), pd.Series(dtype=object)
        for name, p in parts:                      # first source wins on any day
            p = p[~p.index.isin(s.index)]
            s = pd.concat([s, p])
            src = pd.concat([src, pd.Series(name, index=p.index)])
        cols[f"{stem}Level_m"] = s.sort_index()
        source[stem] = src.sort_index()
    for c in ["Limay_m3s", "Neuquen_m3s", "CollonCura_m3s", "Futaleufu_m3s"]:
        if c in flows_arc:
            cols[c] = flows_arc[c].dropna()
    for ina_col, cam_col in [("Parana_Yacyreta_m3s", "Parana_m3s_CAMMESA"), ("Uruguay_SaltoGrande_m3s", "Uruguay_m3s_CAMMESA")]:
        s = ina_arc[ina_col].dropna() if ina_col in ina_arc else pd.Series(dtype=float)
        if cam_col in flows_arc:
            s = s.combine_first(flows_arc[cam_col].dropna())
        cols[ina_col] = s
    daily = pd.DataFrame(cols).sort_index()
    daily = daily[daily.index >= pd.Timestamp(ARGS.start)].dropna(how="all")
    daily["Comahue_level_source"] = pd.concat([source[s] for s in COMAHUE], axis=1, sort=True).reindex(daily.index).apply(
        lambda r: " + ".join(sorted(set(r.dropna()))) or None, axis=1)
    daily.index.name = "date"
    return daily.round(3)


def write(daily, levels_arc, flows_arc, weekly_arc, aic_arc, ina_arc):
    def dated(df):
        df = df.copy()
        df.index = pd.DatetimeIndex(df.index).date
        df.index.name = "date"
        return df

    lev = [c for c in daily.columns if c.endswith("Level_m")]
    last = {c: (daily[c].last_valid_index(), daily[c].dropna().iloc[-1]) for c in daily.columns
            if c.endswith(("_m", "_m3s")) and daily[c].notna().any()}
    weekly_span = ""
    if not weekly_arc.empty and weekly_arc["week_start"].notna().any():
        ws = pd.to_datetime(weekly_arc["week_start"]).dropna()
        weekly_span = f"{ws.min():%Y-%m-%d} to {ws.max():%Y-%m-%d} ({len(ws)} weeks read)"
    notes = [
        "UNITS",
        "*Level_m: reservoir (lake) level, metres above sea level (msnm), one value per day.",
        "*_m3s: river flow, daily mean, m3/s.",
        "Comahue_level_source: where that day's Comahue lake levels come from (see METHOD).",
        "",
        "COVERAGE",
        "Comahue (Limay / Neuquen basins): ChoconLevel_m = El Chocon (Ezequiel Ramos Mexia lake); PiedraAguilaLevel_m;",
        "AlicuraLevel_m; CerrosColoradosLevel_m = Cerros Colorados complex (CAMMESA plant PBANHI, Planicie Banderita),",
        "the level of Los Barreales lake, which regulates the complex (Mari Menuco, below it, is not in CAMMESA's file);",
        "PichiPicunLevel_m = Pichi Picun Leufu. Patagonia: FutaleufuLevel_m (Amutui Quimey lake).",
        "Run-of-river binational plants: YacyretaLevel_m (CAMMESA, from 2023), SaltoGrandeLevel_m (INA 'Salto Grande Arriba',",
        "the lake side of the dam, from 1995; CAMMESA's SGDEHIAR value where INA has no day).",
        "Flows: Limay_m3s, Neuquen_m3s, CollonCura_m3s, Futaleufu_m3s (CAMMESA, from 2023); Parana_Yacyreta_m3s (INA, Parana",
        "entering Yacyreta, from 2006) and Uruguay_SaltoGrande_m3s (INA, Uruguay at Salto Grande, from 1995).",
        f"Daily sheet: {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}. Latest values:",
        *[f"  {c}: {v:,.2f} on {d:%Y-%m-%d}" for c, (d, v) in last.items()],
        "Stored volume (hm3) and % of useful volume are NOT published by CAMMESA, AIC or INA, so there is no volume-",
        "weighted Comahue storage %; the charts show levels (and flows for the run-of-river plants).",
        "",
        "SOURCE",
        f"CAMMESA 'Cotas Diarias' {COTAS_PAGE} (daily levels, 2023-01-01 to the end of the last closed month).",
        f"CAMMESA 'Caudales Diarios' {CAUDALES_PAGE} (daily river flows, same span).",
        f"CAMMESA weekly programme (Programacion Semanal), {LOOKUP_URL}?nemo={PSEM_NEMO} - psemWWYY.zip, Access table COTAS;",
        f"  read so far: {weekly_span}.",
        f"AIC (Autoridad Interjurisdiccional de las Cuencas de los rios Limay, Neuquen y Negro) {AIC_LIST} - 'Nivel Actual' per lake.",
        "INA (Instituto Nacional del Agua) a5 database https://alerta.ina.gob.ar/a5 - series 26684 (Yacyreta afluente, daily",
        "  mean flow), 26674 (Salto Grande Arriba, daily mean flow), 26319 (Salto Grande Arriba, daily mean level).",
        "Raw tabs: 'CAMMESA daily levels' (by CAMMESA plant code), 'CAMMESA daily flows', 'Weekly programme' (CotaIni by",
        "weekly-programme plant code; PDAGHI = Piedra del Aguila, PICUHI = Pichi Picun Leufu), 'AIC snapshots', 'INA'.",
        "",
        "METHOD",
        "Each day's level is taken from the first source that has it, in this order: CAMMESA daily file; (Salto Grande: INA);",
        "AIC reading stored on the day of the run; CAMMESA weekly programme. The weekly programme gives one value per week:",
        "CotaIni, the level CAMMESA expects at the start of the programmed week (a Monday), set 3-4 days before - usually within",
        "0.1 m of the daily file (up to ~1 m at Alicura, whose level moves fastest). So before 2023 the level columns have one",
        "value per week (placed on the week's Monday) and are blank in between; the water-year charts join those weekly points",
        "with straight lines (gaps of up to 15 days only) so the 5-year band is comparable across years. Weeks after the daily",
        "file ends use the weekly programme and AIC readings until CAMMESA publishes the month.",
        "A few programmes give an implausible week start (FInicio more than a week after, or before, its publication); the value is",
        "placed on the Monday after publication (programmes are published on the Thursday/Friday before the week).",
        "Values outside each lake's height band (about 2 m below AIC's minimum extraordinary level to 1 m above its maximum",
        "level) are dropped. CAMMESA's daily file sometimes lists a day's levels against the wrong plant codes (seen from",
        "2026-05-03: some values repeated on other plants' rows, some plants missing). On such days each value is given to the",
        "one lake whose height band holds it; Futaleufu and Pichi Picun Leufu share a band, so a value there goes to the lake",
        "whose previous good value is within 1 m, otherwise neither (left blank)."
        + (f" Days repaired this way in this run's file: {len(REPAIRED_DAYS)} ({min(REPAIRED_DAYS):%Y-%m-%d} to "
           f"{max(REPAIRED_DAYS):%Y-%m-%d})." if REPAIRED_DAYS else ""),
        "Updated daily by .github/workflows/argentina_hydro_reservoirs.yml (incremental: the workbook is the archive;",
        "weekly programmes already read are not fetched again, INA is re-read from 30 days before its last stored day).",
    ]
    sheets = {"Daily": dated(daily)}
    for name, df in [("CAMMESA daily levels", levels_arc), ("CAMMESA daily flows", flows_arc), ("AIC snapshots", aic_arc),
                     ("INA", ina_arc)]:
        if not df.empty:
            sheets[name] = dated(df.sort_index())
    if not weekly_arc.empty:
        w = weekly_arc.copy()
        w["week_start"] = pd.to_datetime(w["week_start"], errors="coerce").dt.date
        w.index.name = "file"
        sheets["Weekly programme"] = w
    Path(ARGS.out).parent.mkdir(parents=True, exist_ok=True)
    xlsx_notes.write_workbook(ARGS.out, sheets, notes, {"UNITS", "COVERAGE", "SOURCE", "METHOD"})
    print(f"\nWrote {ARGS.out}: {len(daily):,} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d} "
          f"({elapsed_min():.1f} min)", flush=True)
    with pd.option_context("display.width", 250, "display.max_columns", 30):
        print(daily[lev + ["Comahue_level_source"]].tail(12).to_string(), flush=True)
        print(daily.drop(columns=lev + ["Comahue_level_source"]).tail(5).to_string(), flush=True)
        print("non-null days per column:", daily.notna().sum().to_dict(), flush=True)


if __name__ == "__main__":
    if ARGS.aic_only:
        arc = store_aic(aic_snapshot())
        print(f"{AIC_CSV}: {len(arc)} days of AIC readings", flush=True)
    else:
        main()
