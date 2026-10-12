"""
Europe FUTURE power capacity (pipeline, retirements, targets): projects and national pipeline totals that add or remove generating
capacity from 2025 on, in the shared row schema of the master future-capacity table.

Output   output/Data and Chart Outputs/europe_future_power_capacity.xlsx
         Projects (the schema; one row per project per phase; MW as published in capacity_value, MW in capacity_mw, retirements
         NEGATIVE), History (changes of date / capacity / status between runs, rows dropped), Summary (MW by country x fuel x year x
         status, rows flagged sum_flag = Y only), Sources, Tried (sources checked and not usable), Releases (fingerprint of every
         source file, whole files are re-read only when their Last-Modified / name / hash changes), native chart sheets
         (net additions by fuel by year: Europe and the six countries with most pipeline, firm = under construction + FID vs
         including permitted + planned, retirements below zero) and PNG copies in output/PNG Charts/.

Columns  region | country | asset_type | project | operator | fuel_or_type | capacity_value | capacity_unit | capacity_mw | capacity_bcfd |
         status | expected_online | online_date_basis | first_seen | last_seen | source_publisher | source_document | source_url |
         as_of_date | evidence | notes   (the shared schema)  + row_id | dedup_key | overlaps_with | sum_flag | source_key.
         overlaps_with = other sources that cover the same country and fuel (aggregates and plants must not be summed together
         blindly); sum_flag = N for a row that duplicates a row of a more official source (kept, with the pointer in notes) and for
         (nothing else).

Sources  (all fetched in GitHub Actions; keyless)
  Germany  MaStR Gesamtdatenexport 'In Planung' units with GeplantesInbetriebnahmedatum (plants >= 50 MW individually, the rest as
           power_aggregate per fuel x planned quarter; same range-request reader as GERMANY_MASTR_CAPACITY_MONTHLY.py);
           Bundesnetzagentur Kraftwerksliste 'Zu- und Rueckbau' (conventional units under construction / in test operation with the
           expected year, and expected closures 2026-2029: KVBG lignite reduction path, notified and further planned closures).
  UK       DESNZ REPD (Under Construction, Awaiting Construction, Application Submitted), CfD allocation rounds AR6 / AR7 / AR7a
           (delivery year and target commissioning date, merged into the REPD rows by name and size), NESO TEC register (connection
           dates, status QUEUE).
  France   ODRE 'suivi-projet-raccordement-enr' (renewable projects in development, RTE / Enedis / ELD queue, 31 Dec 2024).
  Europe   official-document facts (nuclear new build, closures, targets) that are only written when the document
           fetched in this run contains the number and the date (regex check, the matched text is the evidence).
  What was tried and is not usable is listed on 'Tried' (and in discovery_archive/future_capacity/EUROPE_PIPELINE_PROBE1-8.py).

Rules     * statuses: UNDER_CONSTRUCTION | FID | PERMITTED | PLANNED | QUEUE | SUSPENDED | CANCELLED | COMMISSIONED_RECENT;
            retirements: statutory / notified closures = FID, other announced closures = PLANNED, capacity_mw NEGATIVE.
          * nothing guessed: a date a source does not give is NOT_STATED; MaStR planned dates already in the past are kept on plant
            rows (note) and put in an undated aggregate for the small units.
          * charts count UNDER_CONSTRUCTION + FID ('firm') and + PERMITTED + PLANNED ('incl. planned'); QUEUE (connection-queue totals) is never
            added to project rows.
Usage    python3 europe/future/EUROPE_POWER_PIPELINE.py [--force] [--skip NAME,NAME] [--out PATH]
"""
import argparse
import collections
import concurrent.futures as cf
import codecs
import difflib
import hashlib
import importlib.util
import io
import os
import re
import sys
import time
from html import unescape

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, ROOT)
import xlsx_charts  # noqa: E402
import xlsx_notes  # noqa: E402

OUT = os.path.join("output", "Data and Chart Outputs", "europe_future_power_capacity.xlsx")
PNG_DIR = os.path.join("output", "PNG Charts")
UA = {"User-Agent": "Mozilla/5.0 (research; energy data)"}
TODAY = pd.Timestamp.today().normalize()
TODAY_S = TODAY.strftime("%Y-%m-%d")

SCHEMA = ["region", "country", "asset_type", "project", "operator", "fuel_or_type", "capacity_value", "capacity_unit", "capacity_mw",
          "capacity_bcfd", "status", "expected_online", "online_date_basis", "first_seen", "last_seen", "source_publisher",
          "source_document", "source_url", "as_of_date", "evidence", "notes"]
EXTRA = ["row_id", "dedup_key", "overlaps_with", "sum_flag", "source_key"]
COLS = SCHEMA + EXTRA
STATUSES = ["UNDER_CONSTRUCTION", "FID", "PERMITTED", "PLANNED", "QUEUE", "SUSPENDED", "CANCELLED", "COMMISSIONED_RECENT"]
FUELS = ["Hydro", "PumpedStorage", "Gas", "Coal", "Nuclear", "Oil", "Wind_Onshore", "Wind_Offshore", "Solar", "Bioenergy", "Storage", "Other"]
CHART_FUELS = ["Nuclear", "Gas", "Coal", "Oil", "Wind_Onshore", "Wind_Offshore", "Solar", "Hydro", "Bioenergy", "PumpedStorage", "Storage", "Other"]
FIRM = ["UNDER_CONSTRUCTION", "FID"]
INCL = FIRM + ["PERMITTED", "PLANNED"]
BIG = 50.0     # MW: plants at or above are individual rows, below are aggregated


def log(*a):
    print(*a, flush=True)


# ------------------------------------------------------------------------------------------------ helpers
def http(method, url, tries=4, **kw):
    last = None
    for k in range(tries):
        try:
            r = requests.request(method, url, headers={**UA, **kw.pop("headers", {})}, timeout=(20, 240), **kw)
            if r.status_code < 400:
                return r
            last = f"HTTP {r.status_code}"
            if r.status_code in (403, 404):
                break
        except Exception as e:  # noqa: BLE001
            last = repr(e)[:120]
        time.sleep(3 * (k + 1))
    raise RuntimeError(f"{last} for {url[:120]}")


def head_fp(url):
    """Last-Modified | ETag | length of a file, '' when the server gives none."""
    try:
        r = http("HEAD", url, tries=2, allow_redirects=True)
        h = r.headers
        parts = [h.get("Last-Modified", ""), h.get("ETag", ""), h.get("Content-Length", "")]
        return "|".join(parts) if any(parts) else ""
    except Exception:  # noqa: BLE001
        return ""


def sha(b):
    return "sha256:" + hashlib.sha256(b).hexdigest()[:24]


def num(v):
    try:
        x = float(str(v).replace(",", "").strip())
        return x if x == x else None
    except (TypeError, ValueError):
        return None


STOP = {"wind", "farm", "windfarm", "offshore", "onshore", "energy", "project", "limited", "ltd", "phase", "extension", "the", "solar", "battery",
        "bess", "storage", "power", "station", "array", "park", "plc", "and", "of", "gmbh", "ag", "kg", "co", "energie", "kraftwerk"}


def norm(s):
    s = unescape(str(s)).lower()
    for a, b in (("ä", "a"), ("ö", "o"), ("ü", "u"), ("ß", "ss"), ("é", "e"), ("è", "e"), ("ô", "o")):
        s = s.replace(a, b)
    return " ".join(t for t in re.findall(r"[a-z0-9]+", s) if t not in STOP)


def similar(a, b):
    if not a or not b:
        return False
    if a == b:
        return True
    if min(len(a), len(b)) >= 7 and (a in b or b in a):
        return True
    if abs(len(a) - len(b)) > 0.3 * max(len(a), len(b)):
        return False
    return difflib.SequenceMatcher(None, a, b).ratio() >= 0.9


def month_of(v):
    """'YYYY-MM' from a date-like value, else ''."""
    t = pd.to_datetime(v, errors="coerce", dayfirst=True)
    return "" if pd.isna(t) else t.strftime("%Y-%m")


def make_row(**kw):
    r = {c: "" for c in COLS}
    r["region"] = "Europe"
    r["capacity_bcfd"] = ""
    r["sum_flag"] = "Y"
    r.update(kw)
    if r["capacity_mw"] != "" and r["capacity_value"] == "":
        r["capacity_value"], r["capacity_unit"] = r["capacity_mw"], "MW"
    if not r["online_date_basis"]:
        r["online_date_basis"] = "NOT_STATED" if not r["expected_online"] else ("SOURCED_MONTH" if len(str(r["expected_online"])) == 7 else "SOURCED_YEAR")
    r["dedup_key"] = f"{r['country']}|{r['fuel_or_type']}|{norm(r['project'])}"
    return r


class Unchanged(Exception):
    def __init__(self, fp):
        self.fp = fp


class Ctx:
    def __init__(self, prev_projects, prev_rel, force):
        self.prev = prev_projects
        self.prev_rel = prev_rel          # key -> dict
        self.rel = {}                     # key -> dict (this run)
        self.tried = []
        self.force = force

    def gate(self, key, fp):
        """Raise Unchanged when the source fingerprint equals the last run's and its rows are stored."""
        if fp and not self.force and self.prev_rel.get(key, {}).get("fingerprint") == fp and (self.prev["source_key"] == key).any():
            raise Unchanged(fp)
        return fp

    def prev_rows(self, key):
        d = self.prev[self.prev["source_key"] == key]
        return [{c: ("" if pd.isna(v) else v) for c, v in r.items()} for r in d.reindex(columns=COLS).to_dict("records")]

    def fail(self, source, what, why):
        self.tried.append({"Source": source, "What": what, "Result": why, "Checked": TODAY_S})


def run_source(ctx, key, label, fn):
    t0 = time.time()
    try:
        rows, fp, note = fn(ctx, key)
        ctx.rel[key] = {"Source": label, "fingerprint": fp, "fetched": TODAY_S, "rows": len(rows), "note": note}
        log(f"[{key}] {len(rows)} rows ({time.time() - t0:.0f}s) {note}")
        return rows, True
    except Unchanged as u:
        rows = ctx.prev_rows(key)
        for r in rows:
            r["last_seen"] = TODAY_S
        ctx.rel[key] = {**ctx.prev_rel.get(key, {}), "Source": label, "fingerprint": u.fp, "rows": len(rows), "note": "unchanged since " + str(ctx.prev_rel.get(key, {}).get("fetched", "?")) + ", stored rows kept"}
        log(f"[{key}] unchanged, {len(rows)} stored rows kept")
        return rows, True
    except Exception as e:  # noqa: BLE001
        rows = ctx.prev_rows(key)
        ctx.fail(label, "fetch / parse", f"{type(e).__name__}: {str(e)[:200]}" + (f" (kept {len(rows)} stored rows)" if rows else ""))
        ctx.rel[key] = {**ctx.prev_rel.get(key, {}), "Source": label, "rows": len(rows), "note": f"FAILED this run ({type(e).__name__}); stored rows kept"}
        log(f"[{key}] FAILED {type(e).__name__}: {str(e)[:200]}")
        return rows, False


FUEL_BY_MASTR = {"Solar": "Solar", "Wind onshore": "Wind_Onshore", "Wind offshore": "Wind_Offshore", "Hydro": "Hydro", "Pumped storage": "PumpedStorage",
                 "Bioenergy": "Bioenergy", "Nuclear": "Nuclear", "Coal (hard)": "Coal", "Lignite": "Coal", "Oil (>= 10 MW)": "Oil",
                 "Oil (< 10 MW)": "Oil", "Gas (natural)": "Gas", "Gas (other gases)": "Gas", "Battery": "Storage"}

# ================================================================================================ Germany: MaStR In Planung
G = None
MASTR_RE = re.compile(r"^Einheiten(Solar|Wind|StromSpeicher|Verbrennung|Biomasse|Kernkraft|Wasser|GeothermieGrubengasDruckentspannung)(_\d+)?\.xml$")
MASTR_FIELDS = ("EinheitBetriebsstatus", "GeplantesInbetriebnahmedatum", "Nettonennleistung", "Land", "Energietraeger", "Technologie",
                "WindAnLandOderAufSee", "ArtDerWasserkraftanlage", "NameKraftwerk", "NameWindpark", "NameStromerzeugungseinheit",
                "EinheitMastrNummer", "Ort", "Registrierungsdatum")
MASTR_FIELD_RE = re.compile(r"<(%s)>([^<]*)(?=<)" % "|".join(MASTR_FIELDS))


def load_germany():
    global G
    if G is None:
        spec = importlib.util.spec_from_file_location("germany_mastr", os.path.join(ROOT, "europe", "capacity", "GERMANY_MASTR_CAPACITY_MONTHLY.py"))
        G = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(G)
    return G


def mastr_worker(args):
    url, name, kind, maps = args
    g = load_germany()
    tag = g.RECORD_TAG[kind]
    close = "</%s>" % tag
    for attempt in range(4):
        out = []
        n = 0
        try:
            z = g.open_zip(url)
            dec = codecs.getincrementaldecoder("utf-16")()
            carry = ""
            with z.open(name) as f:
                while True:
                    b = f.read(1 << 24)
                    if not b:
                        break
                    parts = (carry + dec.decode(b)).split(close)
                    carry = parts.pop()
                    for rec in parts:
                        n += 1
                        if "<EinheitBetriebsstatus>31<" not in rec:
                            continue
                        d = dict(MASTR_FIELD_RE.findall(rec))
                        if d.get("Land", "84") != "84":
                            continue
                        key = g.classify(kind, d, *maps)
                        nm = d.get("NameKraftwerk") or d.get("NameWindpark") or d.get("NameStromerzeugungseinheit") or ""
                        out.append((key, float(d.get("Nettonennleistung") or 0) / 1000.0, d.get("GeplantesInbetriebnahmedatum", ""), nm,
                                    d.get("NameStromerzeugungseinheit", ""), d.get("EinheitMastrNummer", ""), d.get("Ort", "")))
            return name, out, n
        except Exception as e:  # noqa: BLE001
            log(f"   {name}: attempt {attempt + 1} failed ({repr(e)[:100]})")
            time.sleep(10)
    raise RuntimeError("member failed: " + name)


def collect_mastr(ctx, key):
    g = load_germany()
    url = g.export_url()
    fname = url.rsplit("/", 1)[1]
    ctx.gate(key, fname)
    z = g.open_zip(url)
    vals = g.read_catalogue(z)
    maps = g.catalogue_maps(vals)
    tasks = []
    for i in z.infolist():
        m = MASTR_RE.match(i.filename)
        if m:
            tasks.append((i.compress_size, (url, i.filename, m.group(1), maps)))
    tasks.sort(key=lambda t: -t[0])
    log(f"MaStR {fname}: {len(tasks)} members")
    units = []
    total = 0
    with cf.ProcessPoolExecutor(max_workers=14) as pool:
        futs = [pool.submit(mastr_worker, t[1]) for t in tasks]
        for k, fu in enumerate(cf.as_completed(futs), 1):
            name, out, n = fu.result()
            units.extend(out)
            total += n
    log(f"MaStR: {total:,} units read, {len(units):,} 'In Planung'")
    export_day = re.search(r"_(\d{8})_", fname)
    asof = pd.Timestamp(export_day.group(1)).strftime("%Y-%m-%d") if export_day else TODAY_S
    q_now = TODAY.to_period("Q")
    rows = []
    agg = collections.defaultdict(lambda: [0.0, 0, ""])
    stale_small = collections.defaultdict(lambda: [0.0, 0, [], 0.0])
    for key_, mw, planned, nm, unit_nm, mastr, ort in units:
        fuel = FUEL_BY_MASTR.get(key_, "Other")
        if mw <= 0:
            continue
        pt = pd.to_datetime(planned, errors="coerce")
        if mw >= BIG:
            proj = (nm or unit_nm or ort or "unnamed") + (f" / {unit_nm}" if unit_nm and nm and unit_nm != nm else "") + (f" ({mastr})" if mastr else "")
            past = (not pd.isna(pt)) and pt < TODAY
            rows.append(make_row(country="Germany", asset_type="power_plant", project=proj, operator="", fuel_or_type=fuel, capacity_mw=round(mw, 2),
                                 status="PLANNED", expected_online=("" if pd.isna(pt) else pt.strftime("%Y-%m")),
                                 source_publisher="Bundesnetzagentur (MaStR)", source_document=f"Marktstammdatenregister Gesamtdatenexport {fname}, unit status 'In Planung'",
                                 source_url="https://www.marktstammdatenregister.de/MaStR/Datendownload", as_of_date=asof, evidence="SOURCED",
                                 notes=(f"MaStR unit {mastr}; location {ort}; GeplantesInbetriebnahmedatum {planned or 'not given'}"
                                        + ("; planned date already passed (stale registration?)" if past else "")), source_key=key, ))
            continue
        if pd.isna(pt) or pt < q_now.start_time:
            s = stale_small[fuel]
            s[0] += mw; s[1] += 1
            if not pd.isna(pt):
                s[2].append(pt)
            continue
        qk = (fuel, f"{pt.year}-Q{pt.quarter}")
        a = agg[qk]; a[0] += mw; a[1] += 1
    for (fuel, q), (mw, n, _) in sorted(agg.items()):
        y, qq = q.split("-Q")
        rows.append(make_row(country="Germany", asset_type="power_aggregate", project=f"MaStR 'In Planung' units < {BIG:.0f} MW, planned {q}", fuel_or_type=fuel,
                             capacity_mw=round(mw, 1), status="PLANNED", expected_online=f"{y}-{int(qq) * 3:02d}", online_date_basis="SOURCED_QUARTER",
                             source_publisher="Bundesnetzagentur (MaStR)", source_document=f"Marktstammdatenregister Gesamtdatenexport {fname}, unit status 'In Planung', summed by planned quarter",
                             source_url="https://www.marktstammdatenregister.de/MaStR/Datendownload", as_of_date=asof, evidence="DERIVED",
                             notes=f"sum of {n} registered units below {BIG:.0f} MW; expected_online = last month of the quarter of GeplantesInbetriebnahmedatum; excludes the >= {BIG:.0f} MW plants listed individually",
                             source_key=key))
    for fuel, (mw, n, dts, _) in sorted(stale_small.items()):
        rng = f"{min(dts):%Y-%m-%d} to {max(dts):%Y-%m-%d}" if dts else "none given"
        rows.append(make_row(country="Germany", asset_type="power_aggregate", project=f"MaStR 'In Planung' units < {BIG:.0f} MW, no current planned date", fuel_or_type=fuel,
                             capacity_mw=round(mw, 1), status="PLANNED", expected_online="", online_date_basis="NOT_STATED",
                             source_publisher="Bundesnetzagentur (MaStR)", source_document=f"Marktstammdatenregister Gesamtdatenexport {fname}, unit status 'In Planung'",
                             source_url="https://www.marktstammdatenregister.de/MaStR/Datendownload", as_of_date=asof, evidence="DERIVED",
                             notes=f"sum of {n} units with no planned date or one already passed (registered dates {rng}); treated as undated", source_key=key))
    return rows, fname, f"{len(units):,} In-Planung units from {total:,}"


# ================================================================================================ Germany: BNetzA Zu- und Rueckbau
BNETZA_START = ("https://www.bundesnetzagentur.de/DE/Fachthemen/ElektrizitaetundGas/Versorgungssicherheit/Erzeugungskapazitaeten/Kraftwerksliste/start.html")
DE_FUEL = {"abfall": "Other", "erdgas": "Gas", "pumpspeicher": "PumpedStorage", "batteriespeicher": "Storage", "braunkohle": "Coal", "steinkohle": "Coal",
           "mineraloelprodukte": "Oil", "mineralölprodukte": "Oil", "waerme": "Other", "wärme": "Other", "kernenergie": "Nuclear", "biomasse": "Bioenergy"}


def de_fuel(s):
    t = str(s).strip().lower()
    for k, v in DE_FUEL.items():
        if t.startswith(k):
            return v
    return "Other"


def first_name(x):
    return next(sn for sn in x.sheet_names if "Allgemein" in sn)


def collect_bnetza(ctx, key):
    html = http("GET", BNETZA_START).text
    links = re.findall(r'href="([^"]*ZuUndRueckbau\.xlsx[^"]*)"', html)
    if not links:
        raise RuntimeError("no ZuUndRueckbau.xlsx link on the Kraftwerksliste page")
    url = "https://www.bundesnetzagentur.de" + unescape(links[0]) if links[0].startswith("/") else unescape(links[0])
    fp = head_fp(url) or re.sub(r".*v=", "v=", url)
    ctx.gate(key, fp + "|" + url[-12:])
    b = http("GET", url).content
    x = pd.ExcelFile(io.BytesIO(b))
    sheet = [s for s in x.sheet_names if "einzeln" in s.lower()][0]
    d = x.parse(sheet, header=None)
    txt = " ".join(str(v) for v in d.iloc[:4, 1].tolist())
    m = re.search(r"Datenstand\s+(\d\d)\.(\d\d)\.(\d{4})", txt)
    asof = f"{m.group(3)}-{m.group(2)}-{m.group(1)}" if m else TODAY_S
    doc = "Kraftwerksliste Bundesnetzagentur zum erwarteten Zu- und Rückbau 2026 bis 2029 (Datenstand %s), sheet '%s'" % (asof, sheet)
    rows = []
    # block 1: additions
    i0 = next(i for i in range(len(d)) if "Erwarteter Zubau" in str(d.iat[i, 1]))
    hdr = i0 + 1
    cols = {re.sub(r"\s+", " ", str(v)): j for j, v in enumerate(d.iloc[hdr].tolist()) if pd.notna(v)}

    def cidx(pat):
        for k, j in cols.items():
            if re.search(pat, k, re.I):
                return j
        raise KeyError(pat)
    c_mastr, c_op, c_name, c_ort, c_fuel, c_mw, c_year = (cidx(p) for p in (r"MaStR", r"Anlagenbetreiber", r"Anlagenname", r"^Ort", r"Energietr.ger", r"Nettonennleistung", r"Voraussichtliche"))
    r_ = hdr + 1
    while r_ < len(d) and pd.notna(d.iat[r_, c_op]) or (r_ < len(d) and pd.notna(d.iat[r_, c_name])):
        mw = num(d.iat[r_, c_mw])
        if mw:
            yr = num(d.iat[r_, c_year])
            mastr = str(d.iat[r_, c_mastr]).strip() if pd.notna(d.iat[r_, c_mastr]) else ""
            fuel = de_fuel(d.iat[r_, c_fuel])
            rows.append(make_row(country="Germany", asset_type="power_plant", project=f"{d.iat[r_, c_name]}" + (f" ({mastr})" if mastr else ""), operator=str(d.iat[r_, c_op]),
                                 fuel_or_type=fuel, capacity_mw=round(mw, 2), status="UNDER_CONSTRUCTION", expected_online=str(int(yr)) if yr else "",
                                 source_publisher="Bundesnetzagentur", source_document=doc, source_url=url, as_of_date=asof, evidence="SOURCED",
                                 notes=f"'im Bau oder in Probebetrieb', net rated power >= 10 MW per site (§ 35 EnWG notifications); location {d.iat[r_, c_ort]}; MaStR {mastr or 'no number yet'}",
                                 source_key=key))
        r_ += 1
        if r_ < len(d) and str(d.iat[r_, 1]).startswith("*"):
            break
    # block 2: closures
    j0 = next(i for i in range(len(d)) if "Erwartete aus dem Mark" in str(d.iat[i, 1]))
    hdr2 = j0 + 1
    cols2 = {re.sub(r"\s+", " ", str(v)): j for j, v in enumerate(d.iloc[hdr2].tolist()) if pd.notna(v)}

    def cidx2(pat):
        for k, j in cols2.items():
            if re.search(pat, k, re.I):
                return j
        raise KeyError(pat)
    k_mastr, k_op, k_name, k_ort, k_cat, k_fuel, k_mw = (cidx2(p) for p in (r"MaStR", r"Anlagenbetreiber", r"Anlagenname", r"^Ort", r"Kategorie", r"Energietr.ger", r"Leistung"))
    extra_cols = [j for k, j in cols2.items() if j > k_mw]
    cats = collections.Counter()
    r_ = hdr2 + 1
    while r_ < len(d) and pd.notna(d.iat[r_, k_mw]) and pd.notna(d.iat[r_, k_name]):
        mw = num(d.iat[r_, k_mw])
        cat = re.sub(r"\s+", " ", str(d.iat[r_, k_cat]))
        cats[cat] += 1
        yr = re.search(r"(20\d\d)", cat) or next((re.search(r"(20\d\d)", str(d.iat[r_, j])) for j in extra_cols if re.search(r"(20\d\d)", str(d.iat[r_, j]))), None)
        statutory = cat.lower().startswith(("reduktionspfad", "anzeigen"))
        mastr = str(d.iat[r_, k_mastr]).strip() if pd.notna(d.iat[r_, k_mastr]) else ""
        if mw:
            rows.append(make_row(country="Germany", asset_type="power_plant", project=f"{d.iat[r_, k_name]}" + (f" ({mastr})" if mastr else ""), operator=str(d.iat[r_, k_op]),
                                 fuel_or_type=de_fuel(d.iat[r_, k_fuel]), capacity_mw=-round(mw, 2), status="FID" if statutory else "PLANNED",
                                 expected_online=yr.group(1) if yr else "", source_publisher="Bundesnetzagentur", source_document=doc, source_url=url, as_of_date=asof,
                                 evidence="SOURCED", notes=f"closure / exit from the market; BNetzA category: {cat}" + ("" if yr else " (no year in the list)"), source_key=key))
        r_ += 1
    # national yearly totals of the KVBG lignite reduction path (the unit list carries no year)
    first = next((x.parse(sn, header=None) for sn in x.sheet_names if "Allgemein" in sn), None)
    kv = 0
    if first is not None:
        yr_row = next(i for i in range(len(first)) if sum(num(first.iat[i, j]) in (2026, 2027, 2028, 2029) for j in range(first.shape[1])) >= 3)
        yr_col = {int(num(first.iat[yr_row, j])): j for j in range(first.shape[1]) if num(first.iat[yr_row, j]) in (2026, 2027, 2028, 2029)}
        lr = next((i for i in range(len(first)) if str(first.iat[i, 1]).strip().startswith("davon gesetzlicher")), None)
        if lr is not None:
            for yy, j in sorted(yr_col.items()):
                v = num(first.iat[lr, j])
                if v:
                    kv += 1
                    rows.append(make_row(country="Germany", asset_type="power_aggregate", project=f"KVBG lignite reduction path: closures {yy}", fuel_or_type="Coal", capacity_mw=-round(v, 1),
                                         status="FID", expected_online=str(yy), source_publisher="Bundesnetzagentur", source_document=doc.replace(sheet, first_name(x)), source_url=url,
                                         as_of_date=asof, evidence="SOURCED", notes="statutory closure of the large lignite units under KVBG, national total for the year as published; the unit list gives no year per unit, so the unit rows (same MW) are kept undated with sum_flag N",
                                         source_key=key))
            for r in rows:
                if r["asset_type"] == "power_plant" and r["notes"].find("Reduktionspfad") >= 0:
                    r["sum_flag"] = "N"
                    r["notes"] += " | counted in the dated KVBG national-total rows, not here"
    note = f"{kv} KVBG year totals; closure categories: " + "; ".join(f"{k} x{v}" for k, v in cats.items())
    return rows, fp + "|" + url[-12:], note


# ================================================================================================ UK: REPD + CfD
REPD_PAGE = "https://www.gov.uk/government/publications/renewable-energy-planning-database-monthly-extract"
REPD_FUEL = {"Wind Onshore": "Wind_Onshore", "Wind Offshore": "Wind_Offshore", "Solar Photovoltaics": "Solar", "Battery": "Storage",
             "Pumped Storage Hydroelectricity": "PumpedStorage", "Large Hydro": "Hydro", "Small Hydro": "Hydro", "Biomass (dedicated)": "Bioenergy",
             "Biomass (co-firing)": "Bioenergy", "Anaerobic Digestion": "Bioenergy", "Landfill Gas": "Bioenergy", "EfW Incineration": "Bioenergy",
             "Sewage Sludge Digestion": "Bioenergy", "Advanced Conversion Technologies": "Bioenergy", "Compressed Air Energy Storage": "Storage",
             "Liquid Air Energy Storage": "Storage"}
REPD_STATUS = {"Under Construction": "UNDER_CONSTRUCTION", "Awaiting Construction": "PERMITTED", "Application Submitted": "PLANNED"}


def cfd_fuel(t):
    t = str(t).lower()
    if "offshore" in t or "floating" in t:
        return "Wind_Offshore"
    if "wind" in t:
        return "Wind_Onshore"
    if "solar" in t:
        return "Solar"
    if re.search(r"biomass|efw|waste|landfill|anaerobic|advanced conversion|chp", t):
        return "Bioenergy"
    if "hydro" in t:
        return "Hydro"
    return "Other"


def read_cfd_sheet(content, label):
    out = []
    x = pd.ExcelFile(io.BytesIO(content))
    for s in x.sheet_names:
        d = x.parse(s, header=None)
        h = next((i for i in range(min(len(d), 12)) if any(re.match(r"^project name", str(v).strip(), re.I) for v in d.iloc[i].tolist())), None)
        if h is None:
            continue
        cols = {re.sub(r"\s+", " ", str(v)).strip(): j for j, v in enumerate(d.iloc[h].tolist()) if pd.notna(v)}

        def ci(pat, opt=False):
            for k, j in cols.items():
                if re.search(pat, k, re.I):
                    return j
            if opt:
                return None
            raise KeyError(pat)
        c_name, c_unit, c_app, c_tech, c_size, c_del, c_tgt, c_pot = (ci(r"^project name"), ci(r"name of cfd unit", True), ci(r"applicant", True), ci(r"technology"), ci(r"size"),
                                                                      ci(r"delivery year"), ci(r"target commissioning d"), ci(r"^pot", True))
        for i in range(h + 1, len(d)):
            nm = d.iat[i, c_name]
            sz = num(d.iat[i, c_size])
            if pd.isna(nm) or not sz:
                continue
            out.append({"project": str(nm).strip(), "unit": str(d.iat[i, c_unit]).strip() if c_unit is not None and pd.notna(d.iat[i, c_unit]) else "",
                        "applicant": str(d.iat[i, c_app]).strip() if c_app is not None and pd.notna(d.iat[i, c_app]) else "", "tech": str(d.iat[i, c_tech]),
                        "mw": sz, "delivery": str(d.iat[i, c_del]).strip(), "target": month_of(d.iat[i, c_tgt]),
                        "pot": str(d.iat[i, c_pot]) if c_pot is not None and pd.notna(d.iat[i, c_pot]) else "", "round": label})
    return out


def collect_uk_repd(ctx, key):
    html = http("GET", REPD_PAGE).text
    links = re.findall(r'href="(https://assets\.publishing\.service\.gov\.uk/[^"]+\.csv)"', html)
    if not links:
        raise RuntimeError("no REPD csv link on the page")
    url = links[0]
    cfd_files = []
    for ar in ("6", "7"):
        c = http("GET", f"https://www.gov.uk/api/content/government/publications/contracts-for-difference-cfd-allocation-round-{ar}-results").json()
        for a in c.get("details", {}).get("attachments", []):
            u = a.get("url", "")
            if re.search(r"\.xlsx?$", u):
                cfd_files.append((f"AR{ar}" + ("a" if "AR7a" in u or "7a" in str(a.get("title", "")) else ""), u))
    fp = "|".join([url] + [u for _, u in cfd_files])
    ctx.gate(key, fp)
    d = pd.read_csv(io.BytesIO(http("GET", url).content), encoding="cp1252", low_memory=False)
    d.columns = [str(c).strip() for c in d.columns]
    release = os.path.basename(url)
    asof_m = re.search(r"Q([1-4])_(\d{4})", release)
    asof = (pd.Timestamp(int(asof_m.group(2)), 3 * int(asof_m.group(1)), 1) + pd.offsets.MonthEnd(0)).strftime("%Y-%m-%d") if asof_m else TODAY_S
    d["mw"] = pd.to_numeric(d["Installed Capacity (MWelec)"], errors="coerce")
    d["fuel"] = d["Technology Type"].map(REPD_FUEL).fillna("Other")
    d["nname"] = d["Site Name"].map(norm)
    page = "https://www.gov.uk/government/publications/renewable-energy-planning-database-monthly-extract"
    rows = []
    pipe = d[d["Development Status (short)"].isin(REPD_STATUS) & (d["mw"] > 0)]
    big = pipe[pipe["mw"] >= BIG]
    small = pipe[pipe["mw"] < BIG]
    idx_by_fuel = collections.defaultdict(list)
    for i, r in d.iterrows():
        if r["Development Status (short)"] in ("Operational", "Under Construction", "Awaiting Construction") and r["mw"] > 0:
            idx_by_fuel[r["fuel"]].append(i)
    rowmap = {}
    for i, r in big.iterrows():
        st = REPD_STATUS[r["Development Status (short)"]]
        uc = r.get("Under Construction")
        cfd = r.get("CfD Allocation Round")
        rw = make_row(country="United Kingdom", asset_type="power_plant", project=f"{r['Site Name']} (REPD {r['Ref ID']})", operator=str(r.get("Operator (or Applicant)", "")),
                      fuel_or_type=r["fuel"], capacity_mw=round(float(r["mw"]), 2), status=st, expected_online="", online_date_basis="NOT_STATED",
                      source_publisher="DESNZ", source_document=f"Renewable Energy Planning Database {release}", source_url=url, as_of_date=asof, evidence="SOURCED",
                      notes=(f"REPD status {r['Development Status']}; REPD gives no expected operational date" + (f"; construction start {uc}" if pd.notna(uc) else "")
                             + (f"; CfD round {cfd}" if pd.notna(cfd) else "") + f"; {r.get('Country', '')}"), source_key=key)
        rowmap[i] = rw
        rows.append(rw)
    # small projects: aggregate by fuel x status
    for (fuel, st), g in small.groupby(["fuel", "Development Status (short)"]):
        rows.append(make_row(country="United Kingdom", asset_type="power_aggregate", project=f"REPD {st} projects < {BIG:.0f} MW", fuel_or_type=fuel,
                             capacity_mw=round(float(g["mw"].sum()), 1), status=REPD_STATUS[st], source_publisher="DESNZ", source_document=f"Renewable Energy Planning Database {release}",
                             source_url=url, as_of_date=asof, evidence="DERIVED", notes=f"sum of {len(g)} REPD projects below {BIG:.0f} MW with status {st}; no expected dates in REPD; excludes the >= {BIG:.0f} MW projects listed individually",
                             source_key=key))
    # CfD: delivery year / target commissioning for REPD projects, otherwise own rows
    cfd_rows = []
    for label, u in cfd_files:
        try:
            cfd_rows += read_cfd_sheet(http("GET", u).content, label)
        except Exception as e:  # noqa: BLE001
            ctx.fail(f"CfD {label} results workbook", u[-60:], f"{type(e).__name__}: {str(e)[:120]}")
    matched = 0
    small_cfd = 0
    cand = collections.defaultdict(list)
    for fuel_, ids_ in idx_by_fuel.items():
        cand[fuel_] = [(i, d.at[i, "nname"], float(d.at[i, "mw"]), d.at[i, "Development Status (short)"]) for i in ids_]
    for c in cfd_rows:
        fuel = cfd_fuel(c["tech"])
        nm = norm(c["project"])
        best = None
        for i, nn, mw_, st_ in cand.get(fuel, []):
            if abs(mw_ - c["mw"]) <= 0.15 * max(c["mw"], 1) and similar(nm, nn):
                best = (i, st_)
                break
        if best is not None:
            i, st_ = best
            if st_ == "Operational":
                continue                                    # already built
            if i in rowmap:
                rw = rowmap[i]
                if not rw["expected_online"]:
                    rw["expected_online"] = c["target"]
                    rw["online_date_basis"] = "SOURCED_MONTH" if c["target"] else "NOT_STATED"
                    rw["notes"] += f"; CfD {c['round']} target commissioning {c['target'] or 'n/a'}, delivery year {c['delivery']}, unit {c['unit'] or c['project']} {c['mw']:g} MW"
                    rw["source_document"] += f" + CfD {c['round']} successful applicants"
                matched += 1
                continue
        if c["mw"] < BIG:
            small_cfd += 1                                  # small CfD units: covered by the REPD aggregates, no dates carried
            continue
        past = bool(c["target"]) and c["target"] < TODAY.strftime("%Y-%m")
        rows.append(make_row(country="United Kingdom", asset_type="power_plant",
                             project=f"{c['project']}" + (f" - {c['unit']}" if c["unit"] and c["unit"] != c["project"] else "") + f" (CfD {c['round']})",
                             operator=c["applicant"], fuel_or_type=fuel, capacity_mw=round(c["mw"], 2), status="PLANNED", expected_online=c["target"],
                             source_publisher="DESNZ", source_document=f"Contracts for Difference Allocation Round {c['round'][2:]} results: successful applicants",
                             source_url=[u for l, u in cfd_files if l == c["round"]][0], as_of_date=asof, evidence="SOURCED",
                             notes=f"CfD awarded, delivery year {c['delivery']}, pot {c['pot']}; target commissioning date as given" + ("; target date already passed, not matched to an REPD record" if past else ""),
                             source_key=key))
    return rows, fp, f"{len(rows)} rows; {matched} REPD projects dated from CfD results; {len(cfd_rows)} CfD units read, {small_cfd} small ones not carried"


# ================================================================================================ UK: NESO TEC register
def tec_fuel(pt):
    t = str(pt).lower()
    for pat, f in ((r"wind offshore", "Wind_Offshore"), (r"wind onshore", "Wind_Onshore"), (r"nuclear", "Nuclear"), (r"ccgt|ocgt|gas|chp", "Gas"), (r"coal", "Coal"),
                   (r"biomass|waste", "Bioenergy"), (r"pump", "PumpedStorage"), (r"hydrogen|fuel cell", "Other"), (r"hydro", "Hydro"), (r"pv|solar|photo", "Solar"),
                   (r"tidal|wave|geothermal|hydrogen|fuel cell", "Other"), (r"storage", "Storage")):
        if re.search(pat, t):
            return f
    return None


def collect_tec(ctx, key):
    j = http("GET", "https://api.neso.energy/api/3/action/package_show?id=transmission-entry-capacity-tec-register").json()
    res = [r for r in j["result"]["resources"] if str(r.get("format", "")).upper() == "CSV"]
    if not res:
        raise RuntimeError("no CSV resource in the TEC register dataset")
    url = res[0]["url"]
    ctx.gate(key, url)
    t = pd.read_csv(io.BytesIO(http("GET", url).content), low_memory=False)
    asof_m = re.search(r"(\d{1,2})-([a-z]+)-(\d{4})", url, re.I)
    asof = pd.to_datetime(" ".join(asof_m.groups()), errors="coerce") if asof_m else pd.NaT
    asof = TODAY_S if pd.isna(asof) else asof.strftime("%Y-%m-%d")
    t["fuel"] = t["Plant Type"].map(tec_fuel)
    t["inc"] = pd.to_numeric(t["MW Increase / Decrease"], errors="coerce")
    t["eff"] = pd.to_datetime(t["MW Effective From"], dayfirst=True, errors="coerce")
    t = t[t["fuel"].notna() & (t["Project Status"] != "Built") & t["inc"].notna() & (t["inc"] != 0)]
    rows = []
    doc = f"NESO Transmission Entry Capacity (TEC) register {os.path.basename(url)}"
    page = "https://www.neso.energy/data-portal/transmission-entry-capacity-tec-register"
    t["st"] = t["Project Status"].map(lambda s: "UNDER_CONSTRUCTION" if "Construction" in str(s) else "QUEUE")
    big = t[t["inc"].abs() >= BIG]
    for _, r in big.iterrows():
        hybrid = ";" in str(r["Plant Type"])
        rows.append(make_row(country="United Kingdom", asset_type="power_plant", project=f"{r['Project Name']} ({r['Project Number']})", operator=str(r["Customer Name"]),
                             fuel_or_type=r["fuel"], capacity_mw=round(float(r["inc"]), 2), status="PLANNED" if r["inc"] < 0 and r["st"] == "QUEUE" else r["st"],
                             expected_online=("" if pd.isna(r["eff"]) else r["eff"].strftime("%Y-%m")), source_publisher="NESO", source_document=doc, source_url=url,
                             as_of_date=asof, evidence="SOURCED",
                             notes=(f"TEC contracted connection, project status '{r['Project Status']}', site {r['Connection Site']}, plant type {r['Plant Type']}"
                                    + ("; hybrid site: MW is the whole TEC increment, shown under its generating technology" if hybrid else "")
                                    + ("; DECREASE of contracted TEC (closure / reduction)" if r["inc"] < 0 else "")
                                    + "; QUEUE = signed connection agreement, not a construction decision"), source_key=key))
    small = t[t["inc"].abs() < BIG]
    small = small.assign(yr=small["eff"].dt.year.astype("Int64").astype(str).replace("<NA>", ""))
    for (fuel, st, yr), g in small.groupby(["fuel", "st", "yr"]):
        rows.append(make_row(country="United Kingdom", asset_type="power_aggregate", project=f"NESO TEC register entries < {BIG:.0f} MW, {st.lower().replace('_', ' ')}, connection {yr or 'date not given'}",
                             fuel_or_type=fuel, capacity_mw=round(float(g["inc"].sum()), 1), status=st, expected_online=yr, online_date_basis="SOURCED_YEAR" if yr else "NOT_STATED",
                             source_publisher="NESO", source_document=doc, source_url=url, as_of_date=asof, evidence="DERIVED",
                             notes=f"sum of {len(g)} TEC entries below {BIG:.0f} MW (net of decreases) by effective-from year", source_key=key))
    return rows, url, f"{len(big)} entries >= {BIG:.0f} MW; {len(small)} below"


# ================================================================================================ France: ODRE queue
ODRE_FUEL = {"eolien en mer": "Wind_Offshore", "eolien terrestre": "Wind_Onshore", "solaire": "Solar", "hydraulique": "Hydro", "bioénergies": "Bioenergy"}


def collect_odre(ctx, key):
    base = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/suivi-projet-raccordement-enr"
    meta = http("GET", base).json()
    mod = str(meta.get("metas", {}).get("default", {}).get("data_processed", "") or meta.get("metas", {}).get("default", {}).get("modified", ""))
    ctx.gate(key, mod + "|" + str(meta.get("metas", {}).get("default", {}).get("records_count", "")))
    j = http("GET", base + "/records?limit=100&select=filiere%2Creseau%2Csum(projets_en_developpement_au_31_12_2024_mw)%20as%20mw&group_by=filiere%2Creseau").json()
    rows = []
    for r in j["results"]:
        fuel = ODRE_FUEL.get(str(r["filiere"]).strip().lower())
        if not fuel or not r.get("mw"):
            continue
        rows.append(make_row(country="France", asset_type="power_aggregate", project=f"Renewable projects in development ({r['reseau']})", operator="RTE / Enedis / ELD (grid queue)",
                             fuel_or_type=fuel, capacity_mw=round(float(r["mw"]), 1), status="QUEUE", expected_online="", online_date_basis="NOT_STATED",
                             source_publisher="ODRE (RTE, Enedis, ELD, SER, Agence ORE)", source_document="Projets en développement en électricité renouvelable (suivi-projet-raccordement-enr), at 31 Dec 2024",
                             source_url="https://odre.opendatasoft.com/explore/dataset/suivi-projet-raccordement-enr/", as_of_date="2024-12-31", evidence="SOURCED",
                             notes="national total of projects with a grid connection request / agreement; no dates, no project names; overlaps any project-level French source", source_key=key))
    return rows, mod + "|" + str(meta.get("metas", {}).get("default", {}).get("records_count", "")), f"{len(rows)} aggregates"


SOURCES_FUNCS = [("DE_MASTR", "Germany MaStR 'In Planung' units", collect_mastr),
                 ("DE_BNETZA", "Germany BNetzA Zu- und Rückbau list", collect_bnetza),
                 ("UK_REPD_CFD", "UK REPD + CfD allocation rounds", collect_uk_repd),
                 ("UK_TEC", "UK NESO TEC register", collect_tec),
                 ("FR_ODRE", "France ODRE renewable projects in development", collect_odre)]


# ================================================================================================ official-document facts (nuclear, closures)
FACT_TEXT = {}


def fact_text(url):
    """Visible text of a page / pdf, cached for the run."""
    if url in FACT_TEXT:
        return FACT_TEXT[url]
    r = http("GET", url)
    if r.content[:4] == b"%PDF":
        import pdfplumber
        with pdfplumber.open(io.BytesIO(r.content)) as p:
            t = "\n".join((pg.extract_text() or "") for pg in p.pages[:150])
    else:
        t = unescape(re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", r.text, flags=re.S))
    t = re.sub(r"[\s ]+", " ", t)
    FACT_TEXT[url] = t
    return t


EDF_RES = "https://www.edf.fr/en/2025-annual-results"
EDF_HPC = "https://www.edfenergy.com/energy/nuclear-new-build-projects/hinkley-point-c"
CMA_SZC = "https://www.gov.uk/cma-cases/referral-of-the-proposed-sizewell-c-final-investment-decision-fid-scheme-by-the-department-for-energy-security-and-net-zero-desnz-and-uk-infrastructure-bank-ukib"
GOV_SZC = "https://www.gov.uk/government/news/sizewell-c-gets-green-light-with-final-investment-decision"
SZC_FAQ = "https://www.sizewellc.com/about-sizewell-c/faqs/"
SE_M4 = "https://www.seas.sk/tlacove-spravy/mochovce-4-dostali-zelenu-od-jadroveho-dozoru-slovenske-elektrarne-startuju-energeticke-spustanie/"
PAA_PEJ = "https://www.gov.pl/web/paa-en/resolution-on-the-construction-of-first-nuclear-power-plant-approved-by-the-government"
SNN_C34 = "https://nuclearelectrica.ro/snn/en/investment-projects/reactors-3-and-4/"
AXPO_PR = "https://www.axpo.com/group/en/news-and-stories/media-releases.detail.html/news-and-stories/media-releases/2024/axpo-will-operate-the-beznau-nuclear-power-plant-until-2033--inv.html"
AXPO_KKB = "https://www.axpo.com/ch/en/energy/generation-and-distribution/nuclear-power/beznau.html"

# each fact: fields of the row + checks (url, regex, what it proves); the row is written only if every check matches in the page fetched now
FACTS = [
    dict(country="United Kingdom", project="Hinkley Point C - Unit 1", operator="EDF", fuel="Nuclear", status="UNDER_CONSTRUCTION", mw=1630, ev="DERIVED",
         expected="2030", pub="EDF", doc="EDF Group 2025 annual results (20 Feb 2026), Hinkley Point C: 'Schedule for the start of production by Unit 1 adjusted to 2030'; EDF Energy HPC page: twin unit 3,260MW",
         url=EDF_RES, notes="twin-unit UK EPR 3,260 MW split equally (DERIVED); date is the start of production of Unit 1",
         checks=[(EDF_RES, r"Schedule for the start of production by Unit 1 adjusted to 2030", "date"), (EDF_HPC, r"3,260\s?MW", "capacity")]),
    dict(country="United Kingdom", project="Hinkley Point C - Unit 2", operator="EDF", fuel="Nuclear", status="UNDER_CONSTRUCTION", mw=1630, ev="DERIVED",
         expected="", pub="EDF", doc="EDF Energy Hinkley Point C page: twin unit UK EPR 3,260MW", url=EDF_HPC,
         notes="twin-unit UK EPR 3,260 MW split equally (DERIVED); no start date for Unit 2 in the opened documents", checks=[(EDF_HPC, r"3,260\s?MW", "capacity")]),
    dict(country="United Kingdom", project="Sizewell C - Unit 1", operator="EDF / UK government / investors", fuel="Nuclear", status="FID", mw=1600, ev="DERIVED",
         expected="", pub="UK Government / Sizewell C", doc="CMA referral of the Sizewell C FID scheme ('two-unit 3.2 Gigawatt nuclear power station'); gov.uk 'Sizewell C gets green light with final investment decision'; Sizewell C FAQ (generating in the early 2030s)",
         url=CMA_SZC, notes="3.2 GW two-unit station split equally (DERIVED); Sizewell C FAQ: starts generating 'in the early 2030s' (no year given)",
         checks=[(CMA_SZC, r"two-unit 3\.2 Gigawatt nuclear power station", "capacity"), (GOV_SZC, r"final investment decision", "status"), (SZC_FAQ, r"starts generating in the early 2030s", "date")]),
    dict(country="United Kingdom", project="Sizewell C - Unit 2", operator="EDF / UK government / investors", fuel="Nuclear", status="FID", mw=1600, ev="DERIVED",
         expected="", pub="UK Government / Sizewell C", doc="CMA referral of the Sizewell C FID scheme ('two-unit 3.2 Gigawatt nuclear power station')", url=CMA_SZC,
         notes="3.2 GW two-unit station split equally (DERIVED); no Unit 2 date in the opened documents",
         checks=[(CMA_SZC, r"two-unit 3\.2 Gigawatt nuclear power station", "capacity"), (GOV_SZC, r"final investment decision", "status")]),
    dict(country="Slovakia", project="Mochovce 4 (EMO B4)", operator="Slovenské elektrárne", fuel="Nuclear", status="UNDER_CONSTRUCTION", mw=471, ev="SOURCED",
         expected="", pub="Slovenské elektrárne", doc="SE press release 21.09.2026: Mochovce 4 entered energy start-up (power ascension); VVER-440/V213 'inštalovaným elektrickým výkonom 471 MW'", url=SE_M4,
         notes="in commissioning (energetické spúšťanie since 21 Sep 2026), commercial operation date not stated in the release",
         checks=[(SE_M4, r"inštalovaným elektrickým výkonom 471 MW", "capacity"), (SE_M4, r"Mochovce 4 vstúpili v pondelok 21\. septembra 2026", "status")]),
    dict(country="Poland", project="First nuclear power plant (Lubiatowo-Kopalino, AP1000)", operator="Polskie Elektrownie Jądrowe", fuel="Nuclear", status="PLANNED", mw=3750, ev="SOURCED",
         expected="", pub="National Atomic Energy Agency (PAA) / Polish Government", doc="Resolution on the construction of the first nuclear power plant (3 Nov 2022): 'up to 3,750 MWe', AP1000", url=PAA_PEJ,
         notes="the 2022 resolution gave completion 2033 (superseded; a later first-unit date was not read from an opened document, so none is given); 'up to' 3,750 MWe",
         checks=[(PAA_PEJ, r"up to 3,750 MWe", "capacity")]),
    dict(country="Romania", project="Cernavoda unit 3", operator="Nuclearelectrica", fuel="Nuclear", status="PLANNED", mw=700, ev="SOURCED", expected="", pub="Nuclearelectrica",
         doc="Nuclearelectrica, Reactors 3 and 4: 'Capacity 700 MW/Unit; 1400 MW/2 units'", url=SNN_C34, notes="CANDU 6; page gives stages, no reliable commissioning year in the text read",
         checks=[(SNN_C34, r"Capacity 700 MW/Unit; 1400 MW/2 units", "capacity")]),
    dict(country="Romania", project="Cernavoda unit 4", operator="Nuclearelectrica", fuel="Nuclear", status="PLANNED", mw=700, ev="SOURCED", expected="", pub="Nuclearelectrica",
         doc="Nuclearelectrica, Reactors 3 and 4: 'Capacity 700 MW/Unit; 1400 MW/2 units'", url=SNN_C34, notes="CANDU 6; page gives stages, no reliable commissioning year in the text read",
         checks=[(SNN_C34, r"Capacity 700 MW/Unit; 1400 MW/2 units", "capacity")]),
    dict(country="Switzerland", project="Beznau 2 (end of operation)", operator="Axpo", fuel="Nuclear", status="FID", mw=-365, ev="SOURCED", expected="2032", pub="Axpo",
         doc="Axpo media release 5 Dec 2024: 'Unit 2 ... will remain on the grid until 2032 and Unit 1 until 2033'; Beznau plant page: 365 MW each", url=AXPO_PR,
         notes="decision of the operator (4 Dec 2024) after review; expected_online = year of last operation, closure follows",
         checks=[(AXPO_PR, r"Unit 2 of the nuclear power plant will remain on the grid until 2032", "date"), (AXPO_KKB, r"365 megawatts \(MW\) each", "capacity")]),
    dict(country="Switzerland", project="Beznau 1 (end of operation)", operator="Axpo", fuel="Nuclear", status="FID", mw=-365, ev="SOURCED", expected="2033", pub="Axpo",
         doc="Axpo media release 5 Dec 2024: 'Unit 2 ... until 2032 and Unit 1 until 2033'; Beznau plant page: 365 MW each", url=AXPO_PR,
         notes="decision of the operator (4 Dec 2024) after review; expected_online = year of last operation, closure follows",
         checks=[(AXPO_PR, r"Unit 1 until 2033", "date"), (AXPO_KKB, r"365 megawatts \(MW\) each", "capacity")]),
]


def collect_facts(ctx, key):
    rows = []
    h = hashlib.sha256()
    ok = bad = 0
    for f in FACTS:
        ev, miss = [], []
        for url, rx, what in f["checks"]:
            try:
                m = re.search(rx, fact_text(url))
            except Exception as e:  # noqa: BLE001
                miss.append(f"{what}: {type(e).__name__}")
                continue
            if m:
                ev.append(f"{what}: '{m.group(0)[:110]}'")
                h.update(m.group(0).encode())
            else:
                miss.append(f"{what}: text not found")
        if miss or f["mw"] is None:
            if miss:
                bad += 1
                ctx.fail("Official-document fact", f["project"], "not written - " + "; ".join(miss) + f" ({f['url'][:90]})")
            if f["mw"] is None and not miss:
                ctx.fail("Official-document fact", f["project"], "dates verified but no MW in an opened document - no row written. " + f["notes"])
            continue
        ok += 1
        rows.append(make_row(country=f["country"], asset_type=f.get("asset", "power_plant"), project=f["project"], operator=f["operator"], fuel_or_type=f["fuel"], capacity_mw=f["mw"],
                             status=f["status"], expected_online=f["expected"], source_publisher=f["pub"], source_document=f["doc"], source_url=f["url"], as_of_date=TODAY_S,
                             evidence=f["ev"], notes=f["notes"] + " | verified in this run - " + "; ".join(ev),
                             source_key=key))
    return rows, h.hexdigest()[:16], f"{ok} facts verified, {bad} not verifiable"


# ================================================================================================ assemble
def clean_prev(path):
    try:
        sheets = pd.read_excel(path, sheet_name=None)
    except Exception:  # noqa: BLE001
        return {}
    return sheets


def finalize(rows, ctx, failed):
    df = pd.DataFrame(rows, columns=COLS)
    df = df.fillna("")
    prev = ctx.prev
    pfirst = dict(zip(prev["row_id"], prev["first_seen"])) if len(prev) else {}
    df["row_id"] = [hashlib.md5(f"{r.source_key}|{r.country}|{r.fuel_or_type}|{r.project}|{r.asset_type}".encode()).hexdigest()[:12] for r in df.itertuples()]
    df = df.drop_duplicates("row_id", keep="first")
    df["first_seen"] = [pfirst.get(i) or fs or TODAY_S for i, fs in zip(df["row_id"], df["first_seen"])]
    df["last_seen"] = [ls if (sk in failed and ls) else TODAY_S for ls, sk in zip(df["last_seen"], df["source_key"])]
    # duplicates across sources (flag, never delete)
    notes = df["notes"].astype(str).tolist()
    flag = df["sum_flag"].tolist()
    ids = df["row_id"].tolist()

    def mark(i, other, why):
        flag[i] = "N"
        notes[i] = (notes[i] + f" | DUPLICATE of {df.at[df.index[other], 'source_key']} row {ids[other]} ({why}); not counted in Summary").strip(" |")
    see = {}
    for i, (p, sk) in enumerate(zip(df["project"], df["source_key"])):
        m = re.search(r"\((SEE\w+)\)", str(p))
        if m and sk == "DE_BNETZA":
            see[m.group(1)] = i
    for i, (p, sk) in enumerate(zip(df["project"], df["source_key"])):
        m = re.search(r"\((SEE\w+)\)", str(p))
        if m and sk == "DE_MASTR" and m.group(1) in see:
            mark(i, see[m.group(1)], "same MaStR unit number in the Bundesnetzagentur list, the more official source")
    repd = collections.defaultdict(list)
    for i, r in enumerate(df.itertuples()):
        if r.source_key == "UK_REPD_CFD" and r.asset_type == "power_plant" and r.capacity_mw != "" and r.capacity_mw > 0:
            repd[r.fuel_or_type].append((i, norm(re.sub(r"\((?:REPD|CfD)[^)]*\)", "", r.project))))
    for i, r in enumerate(df.itertuples()):
        if r.source_key == "UK_TEC" and r.asset_type == "power_plant" and r.capacity_mw != "" and r.capacity_mw > 0:
            nm = norm(re.sub(r"\(PRO-\d+\)", "", r.project))
            for j, rn in repd.get(r.fuel_or_type, []):
                if similar(nm, rn) and abs(float(df.iloc[j]["capacity_mw"]) - float(r.capacity_mw)) <= 0.25 * max(float(r.capacity_mw), 1):
                    mark(i, j, "same project name, fuel and size; REPD/CfD row kept")
                    break
    df["notes"], df["sum_flag"] = notes, flag
    # overlap hints: other sources with pipeline rows for the same country and fuel
    cover = collections.defaultdict(set)
    for r in df.itertuples():
        if True:
            cover[(r.country, r.fuel_or_type)].add(r.source_key)
    df["overlaps_with"] = [";".join(sorted(cover[(c, f)] - {s}))
                           for c, f, s in zip(df["country"], df["fuel_or_type"], df["source_key"])]
    df["capacity_value"] = [mw if cv == "" else cv for mw, cv in zip(df["capacity_mw"], df["capacity_value"])]
    df["capacity_unit"] = ["MW" if (u == "" and cv != "") else u for u, cv in zip(df["capacity_unit"], df["capacity_value"])]
    df["as_of_date"] = df["as_of_date"].astype(str)
    order = {s: i for i, s in enumerate(STATUSES)}
    df = df.assign(_o=df["status"].map(order), _m=pd.to_numeric(df["capacity_mw"], errors="coerce")).sort_values(
        ["country", "_o", "fuel_or_type", "expected_online", "_m"], ascending=[True, True, True, True, False]).drop(columns=["_o", "_m"])
    return df.reset_index(drop=True)


def history_rows(df, ctx, ok_keys):
    out = []
    prev = ctx.prev
    if not len(prev):
        return out
    p = prev.set_index("row_id")
    n = df.set_index("row_id")
    for rid in n.index.intersection(p.index):
        for f in ("expected_online", "capacity_mw", "status"):
            o, v = p.at[rid, f], n.at[rid, f]
            so = "" if pd.isna(o) else str(o)
            sv = "" if pd.isna(v) else str(v)
            try:
                if so != "" and sv != "" and abs(float(so) - float(sv)) < 0.05:
                    continue
            except ValueError:
                pass
            if so != sv:
                out.append({"date_seen": TODAY_S, "row_id": rid, "country": n.at[rid, "country"], "project": n.at[rid, "project"], "fuel": n.at[rid, "fuel_or_type"],
                            "field": f, "old": so, "new": sv, "source_key": n.at[rid, "source_key"]})
    for rid in p.index.difference(n.index):
        if p.at[rid, "source_key"] in ok_keys:
            out.append({"date_seen": TODAY_S, "row_id": rid, "country": p.at[rid, "country"], "project": p.at[rid, "project"], "fuel": p.at[rid, "fuel_or_type"],
                        "field": "row", "old": "listed", "new": "no longer in the source", "source_key": p.at[rid, "source_key"]})
    return out


def year_of(v):
    s = str(v)
    return int(s[:4]) if re.match(r"^\d{4}", s) else None


def make_summary(df):
    d = df[(df["sum_flag"] == "Y") & (df["capacity_mw"] != "")].copy()
    d["capacity_mw"] = d["capacity_mw"].astype(float)
    d["year"] = [("Undated" if year_of(e) is None else (str(year_of(e)) if year_of(e) >= 2026 else "2025 or earlier (date passed)")) for e in d["expected_online"]]
    d["kind"] = d["asset_type"].map(lambda a: "plants_MW" if a == "power_plant" else "aggregates_MW")
    g = d.pivot_table(index=["country", "fuel_or_type", "year", "status"], columns="kind", values="capacity_mw", aggfunc="sum", fill_value=0.0)
    for c in ("plants_MW", "aggregates_MW"):
        if c not in g.columns:
            g[c] = 0.0
    g["total_MW"] = g["plants_MW"] + g["aggregates_MW"]
    g["rows"] = d.groupby(["country", "fuel_or_type", "year", "status"]).size()
    return g[["plants_MW", "aggregates_MW", "total_MW", "rows"]].round(1).reset_index()


YEARS = list(range(2026, 2031))
LABN = {"firm": "firm", "incl": "incl_planned"}
LABEL = {"Wind_Onshore": "Wind onshore", "Wind_Offshore": "Wind offshore", "PumpedStorage": "Pumped storage", "Storage": "Batteries"}


def chart_tables(df, scope, statuses):
    d = df[(df["sum_flag"] == "Y") & df["status"].isin(statuses) & (df["capacity_mw"] != "")].copy()
    d["mw"] = d["capacity_mw"].astype(float)
    d["y"] = [year_of(e) for e in d["expected_online"]]
    d = d[d["y"].isin(YEARS)]
    if scope != "Europe":
        d = d[d["country"] == scope]
    idx = pd.to_datetime([f"{y}-01-01" for y in YEARS])
    add = pd.DataFrame(0.0, index=idx, columns=CHART_FUELS)
    ret = pd.DataFrame(0.0, index=idx, columns=CHART_FUELS)
    for r in d.itertuples():
        t = pd.Timestamp(f"{int(r.y)}-01-01")
        f = r.fuel_or_type if r.fuel_or_type in CHART_FUELS else "Other"
        (add if r.mw > 0 else ret).at[t, f] += r.mw / 1000.0
    add = add.loc[:, add.abs().sum() > 0]
    ret = ret.loc[:, ret.abs().sum() > 0]
    out = pd.concat([add.rename(columns=LABEL), ret.rename(columns=lambda c: LABEL.get(c, c) + " (retirements)")], axis=1)
    out.index.name = "Year"
    return out.round(3)


def series_colour(name):
    base = name.replace(" (retirements)", "")
    key = {v: k for k, v in LABEL.items()}.get(base, base)
    c = xlsx_charts.fuel_colour(key)
    if key in ("Wind_Offshore", "PumpedStorage"):
        c = xlsx_charts._lighten(c, 0.35)
    if name.endswith("(retirements)"):
        c = xlsx_charts._lighten(c, 0.5)
    return c


def add_chart_sheets(path, df):
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    firm_all = df[(df["sum_flag"] == "Y") & df["status"].isin(INCL) & (df["capacity_mw"] != "") & df["expected_online"].map(lambda e: year_of(e) in YEARS)]
    top = firm_all.assign(mw=firm_all["capacity_mw"].astype(float)).query("mw > 0").groupby("country")["mw"].sum().sort_values(ascending=False).head(6).index.tolist()
    scopes = ["Europe"] + top
    wb = load_workbook(path)
    made = {}
    for sc in scopes:
        name = f"Chart - {sc}"[:31]
        if name in wb.sheetnames:
            del wb[name]
        ws = wb.create_sheet(name)
        made[sc] = {}
        row0 = 1
        for lab, sts, title in (("firm", FIRM, "under construction + FID"), ("incl", INCL, "incl. permitted + planned")):
            t = chart_tables(df, sc, sts)
            made[sc][lab] = t
            ws.cell(row=row0, column=1, value=f"{sc}: net capacity additions by fuel and year, {title} (GW; retirements below zero; QUEUE rows not included)")
            xlsx_charts.write_table(ws, t, "%Y", start_row=row0 + 1, start_col=1)
            if t.shape[1]:
                ch = BarChart()
                ch.type, ch.grouping, ch.overlap, ch.gapWidth = "col", "stacked", 100, 40
                n = row0 + 1 + len(t)
                ch.add_data(Reference(ws, min_col=2, max_col=1 + t.shape[1], min_row=row0 + 1, max_row=n), titles_from_data=True)
                ch.set_categories(Reference(ws, min_col=1, min_row=row0 + 2, max_row=n))
                for s, c in zip(ch.series, t.columns):
                    s.graphicalProperties = GraphicalProperties(solidFill=series_colour(c), ln=LineProperties(noFill=True))
                ch.title = f"{sc}: capacity additions and retirements by year, {title} (GW)"
                ch.x_axis.number_format = "yyyy"
                ch.y_axis.title = "GW"
                ch.x_axis.tickLblPos = "low"
                ch.x_axis.delete = False
                ch.y_axis.delete = False
                ch.legend.position = "b"
                ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
                ch.plot_area.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
                ch.width, ch.height = 24, 11
                xlsx_charts.tidy_layout(ch)
                ws.add_chart(ch, f"{get_column_letter(t.shape[1] + 4)}{row0 + 1}")
            row0 += 22
    xlsx_charts.save_atomic(wb, path)
    return scopes, made


def png_charts(scopes, made):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    os.makedirs(PNG_DIR, exist_ok=True)
    out = []

    def draw(ax, t, title):
        yrs = [d.year for d in t.index]
        pos = [0.0] * len(yrs)
        neg = [0.0] * len(yrs)
        for c in t.columns:
            vals = t[c].tolist()
            col = "#" + series_colour(c)
            if c.endswith("(retirements)"):
                ax.bar(yrs, vals, bottom=neg, color=col, width=0.7, label=c)
                neg = [a + b for a, b in zip(neg, vals)]
            else:
                ax.bar(yrs, vals, bottom=pos, color=col, width=0.7, label=c)
                pos = [a + b for a, b in zip(pos, vals)]
        ax.axhline(0, color="#444", lw=0.6)
        ax.set_title(title, fontsize=12, fontweight="bold", loc="left")
        ax.set_ylabel("GW", fontsize=9)
        ax.set_xticks(yrs)
        for s in ("top", "right"):
            ax.spines[s].set_visible(False)
    for lab, nice in (("firm", "under construction + FID"), ("incl", "incl. permitted + planned")):
        t = made["Europe"][lab]
        if not t.shape[1]:
            continue
        fig, ax = plt.subplots(figsize=(10, 5.2))
        draw(ax, t, f"Europe: capacity additions and retirements by year, {nice} (GW)")
        ax.legend(ncol=4, fontsize=8, frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.08))
        fig.text(0.01, 0.005, "Sources: MaStR, Bundesnetzagentur, DESNZ REPD + CfD results, NESO TEC, official operator / government documents; sum_flag = Y rows only. Planned dates as published.", fontsize=7, color="#555")
        fig.tight_layout(rect=(0, 0.02, 1, 1))
        p = os.path.join(PNG_DIR, f"europe_future_power_capacity__Europe_{LABN[lab]}.png")
        fig.savefig(p, dpi=130)
        plt.close(fig)
        out.append(p)
        others = [s for s in scopes if s != "Europe"]
        if others:
            fig, axes = plt.subplots(2, 3, figsize=(15, 8), sharey=False)
            for ax, sc in zip(axes.ravel(), others):
                tt = made[sc][lab]
                if tt.shape[1]:
                    draw(ax, tt, sc)
            for ax in axes.ravel()[len(others):]:
                ax.axis("off")
            h, l = {}, {}
            for ax in axes.ravel():
                for hh, ll in zip(*ax.get_legend_handles_labels()):
                    h[ll] = hh
            fig.legend(list(h.values()), list(h.keys()), ncol=6, fontsize=8, frameon=False, loc="lower center")
            fig.suptitle(f"Top six countries by pipeline: additions and retirements by year, {nice} (GW)", fontsize=12, fontweight="bold", x=0.01, ha="left")
            fig.tight_layout(rect=(0, 0.05, 1, 0.95))
            p = os.path.join(PNG_DIR, f"europe_future_power_capacity__Top6_{LABN[lab]}.png")
            fig.savefig(p, dpi=120)
            plt.close(fig)
            out.append(p)
    return out


STATIC_TRIED = [
    ("Bundesnetzagentur offshore wind tender results and Kohleausstieg auction pages", "guessed page URLs under /Ausschreibungen/", "HTTP 404; the Ausschreibungen overview page is reachable but holds no structured tender results (WindSeeG awards are not read; offshore units come from MaStR 'In Planung' and CfD results)"),
    ("BSH Flächenentwicklungsplan 2025 (German offshore sites, commissioning years)", "bsh.de Meeresfachplanung page", "page reachable (links to FEP 2025 / 2023 PDFs), PDFs not parsed - the table layout was not read in this build"),
    ("France RTE / Enedis connection queue by project ('file d'attente')", "data.rte-france.com, ODRE catalogue search", "RTE open-data portal needs a login; ODRE's only queue dataset is the regional aggregate 'suivi-projet-raccordement-enr' (used, aggregates only); no project list"),
    ("France PPE decree text, Legifrance", "legifrance.gouv.fr", "HTTP 403 (Cloudflare); the PPE offshore targets are read from ecologie.gouv.fr instead"),
    ("Spain REE access and connection grants (nudos de acceso, capacidad concedida)", "ree.es/es/clientes/generador/acceso-conexion", "'conoce-el-estado-de-tu-solicitud' 404; 'acceso-conexion' reachable but lists no data file; MITECO energy landing has no machine-readable queue"),
    ("Spain nuclear closure schedule per unit in MW", "MITECO 7th PGRR (27 Dec 2023) and BOE-A-2026-17756 (14 Aug 2026)", "dates verified (Almaraz I 11/2027 -> extended with II to 8 Jun 2030 by the 2026 order; Asco I 10/2030, Cofrentes 11/2030, Asco II 9/2032, Vandellos II 2/2035, Trillo 5/2035) but no per-unit electrical MW in an opened official document, so no negative-MW rows; no row is written"),
    ("Italy Terna connection requests", "terna.it connection-requests page", "reachable, only 921 characters of text and no figures; Terna's Econnextion queue is not downloadable as a file"),
    ("Netherlands RVO offshore wind tender results, Denmark Energistyrelsen tenders, Poland offshore (gov.pl)", "english.rvo.nl, rvo.nl, ens.dk, gov.pl", "pages carry no GW / MW figures in the text read (RVO, gov.pl) or answer 404 (ens.dk); not used"),
    ("TenneT, IAEA PRIS, Legifrance, World Nuclear Association country profile", "tennet.eu, iaea.org, world-nuclear.org", "HTTP 403 / 404 from Actions"),
    ("NECP national targets and ENTSO-E ERAA / TYNDP scenario capacities", "owner's decision 12 Oct 2026", "dropped on purpose: raw project / unit data only, no target rows, no targets code"),
    ("Paks II, Dukovany II, Belgian Doel 4 / Tihange 3 extension, Hungarian / Czech / Polish dates", "paks2.hu, mpo.gov.cz, Engie newsroom", "pages read but carry no MW / year pair in their text (Paks II EN page: news only; MPO: no MW; Engie closing release: no MW and no 2035); not written"),
    ("EDF UK AGR closure dates (Heysham 1, Hartlepool, Heysham 2, Torness)", "edfenergy.com/energy/nuclear-power-stations", "HTTP 404"),
    ("Poland PSE coal plans, Italy coal exit, Spain coal", "pse.pl", "only capacity-market notices; no dated closure list read"),
    ("Poland PEJ site (ppej.pl)", "ppej.pl", "connection reset"),
    ("CfD Allocation Round 5 successful-projects workbook and AR4 results (pdf)", "gov.uk content API attachment", "AR5 .xls URL answered 404; AR4 is a PDF and mostly under construction in REPD; only AR6, AR7 and AR7a workbooks are read"),
    ("NESO TEC register 'Built' rows", "api.neso.energy", "left out on purpose (already connected)"),
]

SOURCES_TABLE = [
    ("Bundesnetzagentur", "Marktstammdatenregister Gesamtdatenexport (unit status In Planung)", "https://www.marktstammdatenregister.de/MaStR/Datendownload", "Germany planned units (>= 50 MW individually, rest aggregated by planned quarter)"),
    ("Bundesnetzagentur", "Kraftwerksliste zum erwarteten Zu- und Rückbau 2026 bis 2029", BNETZA_START, "Germany conventional units under construction / test operation; expected closures (KVBG, notified, planned)"),
    ("DESNZ", "Renewable Energy Planning Database (quarterly)", REPD_PAGE, "UK Under Construction / Awaiting Construction / Application Submitted"),
    ("DESNZ", "Contracts for Difference Allocation Round 6, 7, 7a results (successful applicants)", "https://www.gov.uk/government/publications/contracts-for-difference-cfd-allocation-round-7-results", "UK delivery year and target commissioning date, merged into REPD rows"),
    ("NESO", "Transmission Entry Capacity (TEC) register", "https://www.neso.energy/data-portal/transmission-entry-capacity-tec-register", "UK contracted connection dates (QUEUE)"),
    ("ODRE (RTE, Enedis, ELD, SER, Agence ORE)", "Projets en développement en électricité renouvelable", "https://odre.opendatasoft.com/explore/dataset/suivi-projet-raccordement-enr/", "France renewable queue aggregates, 31 Dec 2024"),
    ("EDF, UK Government, Sizewell C, Slovenské elektrárne, PAA, Nuclearelectrica, Axpo", "nuclear new build and closure documents (see each row)", "see source_url per row", "official-document facts, written only when the number and date are found in the page fetched in the run"),
]


def notes_lines():
    return ["UNITS", "capacity_mw in MW (retirements NEGATIVE); capacity_value / capacity_unit as published; expected_online YYYY-MM or YYYY; capacity_bcfd blank (power rows).", "",
            "WHAT THIS IS", "Future capacity pipeline for Europe's power generation: projects and national totals that add (positive) or remove (negative) capacity from 2025 on, in the shared row schema of the master future-capacity table, plus 5 extra columns: row_id, dedup_key, overlaps_with, sum_flag, source_key.", "",
            "STATUS", "UNDER_CONSTRUCTION | FID | PERMITTED | PLANNED | QUEUE. Closures: statutory or notified (KVBG reduction path, notified shutdowns, operator decisions) = FID, other announced closures = PLANNED. 'firm' on the charts = UNDER_CONSTRUCTION + FID, 'incl. planned' adds PERMITTED + PLANNED. QUEUE (grid-connection queue totals, TEC register) is never added to project rows. No targets (NECP, ERAA) are included: raw data only.", "",
            "DOUBLE COUNTING", "A plant is listed once per project and phase. The same project in two sources: the less official row keeps sum_flag = N and a pointer in notes (MaStR vs Bundesnetzagentur by MaStR unit number; NESO TEC vs REPD / CfD by name, fuel and size). Aggregates (power_aggregate) and plants of the same country and fuel come from different sources: overlaps_with names the other sources; the Summary adds rows with sum_flag = Y only. UK REPD small projects (< 50 MW), TEC small entries and MaStR small units are separate aggregates that exclude the >= 50 MW plants listed individually, but REPD and TEC aggregates overlap each other for the same fuel (TEC aggregates are QUEUE, REPD aggregates are PERMITTED / PLANNED / UNDER_CONSTRUCTION, so they sit in different status rows).", "",
            "DATES", "Dates are as the source gives them (month, quarter-end month or year, online_date_basis says which). Nothing is guessed: REPD gives no expected operation date (CfD target commissioning dates are used where an AR6 / AR7 / AR7a project matches), MaStR planned dates already in the past stay on plant rows with a note and small units with a passed date are grouped as undated. Summary buckets dates before 2026 as '2025 or earlier (date passed)'.", "",
            "UPDATES", "Each source is fingerprinted on 'Releases' (Last-Modified / file name / hash); a source that has not changed keeps its stored rows. 'History' logs changes in expected_online, capacity_mw or status between runs and rows that disappear from a source. first_seen / last_seen are run dates.", "",
            "CAVEATS", "MaStR 'In Planung' contains speculative and stale registrations; REPD 'Awaiting Construction' is planning permission, not a construction decision; TEC entries are contracted connections, not projects with funding. Spanish nuclear closures: dates verified but no per-unit MW in an official document, see Tried."]


def build_workbook(df, hist, rel, tried, out):
    summ = make_summary(df)
    sheets = {}
    pr = df[COLS].copy()
    pr["capacity_mw"] = pd.to_numeric(pr["capacity_mw"], errors="coerce")
    pr["capacity_value"] = pd.to_numeric(pr["capacity_value"], errors="coerce")
    sheets["Projects"] = pr.set_index("region")
    sheets["History"] = pd.DataFrame(hist, columns=["date_seen", "row_id", "country", "project", "fuel", "field", "old", "new", "source_key"]).set_index("date_seen")
    sheets["Summary"] = summ.set_index("country")
    sheets["Sources"] = pd.DataFrame(SOURCES_TABLE, columns=["publisher", "document", "url", "used for"]).set_index("publisher")
    sheets["Tried"] = pd.DataFrame(tried, columns=["Source", "What", "Result", "Checked"]).set_index("Source")
    sheets["Releases"] = pd.DataFrame([{"key": k, **v} for k, v in rel.items()]).set_index("key")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    xlsx_notes.write_workbook(out, sheets, notes_lines(), {"UNITS", "WHAT THIS IS", "STATUS", "DOUBLE COUNTING", "DATES", "UPDATES", "CAVEATS"})
    return summ


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--skip", default="")
    a = ap.parse_args()
    skip = {s.strip() for s in a.skip.split(",") if s.strip()}
    prev_sheets = clean_prev(a.out)
    prev = prev_sheets.get("Projects", pd.DataFrame(columns=COLS))
    for c in COLS:
        if c not in prev.columns:
            prev[c] = ""
    prev = prev.fillna("")
    prev_rel = {}
    if "Releases" in prev_sheets:
        pr_ = prev_sheets["Releases"].fillna("")
        prev_rel = {r["key"]: r for r in pr_.to_dict("records")}
    ctx = Ctx(prev, prev_rel, a.force)
    rows, failed, ok_keys = [], set(), set()
    funcs = SOURCES_FUNCS + [("FACTS", "Official-document facts (nuclear, closures)", collect_facts)]
    for key, label, fn in funcs:
        if key in skip:
            r = ctx.prev_rows(key)
            rows += r
            ctx.rel[key] = {**prev_rel.get(key, {}), "Source": label, "note": "skipped this run"}
            continue
        r, ok = run_source(ctx, key, label, fn)
        rows += r
        (ok_keys if ok else failed).add(key)
    df = finalize(rows, ctx, failed)
    hist = (prev_sheets["History"].fillna("").to_dict("records") if "History" in prev_sheets else []) + history_rows(df, ctx, ok_keys)
    tried = STATIC_TRIED and [(s, w, r_, TODAY_S) for s, w, r_ in STATIC_TRIED]
    tried = tried + [(t["Source"], t["What"], t["Result"], t["Checked"]) for t in ctx.tried]
    rel = {k: {"Source": v.get("Source", ""), "fingerprint": v.get("fingerprint", ""), "fetched": v.get("fetched", ""), "rows": v.get("rows", ""), "note": v.get("note", "")} for k, v in ctx.rel.items()}
    summ = build_workbook(df, hist, rel, tried, a.out)
    scopes, made = add_chart_sheets(a.out, df)
    pngs = png_charts(scopes, made)
    log("\nRows by country x status (MW, sum_flag Y):")
    d = df[(df["sum_flag"] == "Y") & (df["capacity_mw"] != "")].copy()
    d["capacity_mw"] = d["capacity_mw"].astype(float)
    log(d.pivot_table(index="country", columns="status", values="capacity_mw", aggfunc="sum", fill_value=0).round(0).to_string())
    log("\nRows by source:", df["source_key"].value_counts().to_dict())
    log("PNGs:", pngs)
    log("wrote", a.out, len(df), "rows")


if __name__ == "__main__":
    main()
