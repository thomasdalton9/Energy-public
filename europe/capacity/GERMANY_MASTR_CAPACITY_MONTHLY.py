"""
Germany installed generating capacity by fuel, MONTH-END, rebuilt from the Marktstammdatenregister (MaStR), the official
register of the Bundesnetzagentur (www.marktstammdatenregister.de, 'Gesamtdatenexport', free, no key, licence
Datenlizenz Deutschland - Namensnennung 2.0).

  Source   The daily 'Gesamtdatenauszug' zip (about 3.2 GB; UTF-16 XML). Only the unit files are read, each member streamed
           by HTTP range request straight from the zip (about 1.8 GB compressed) and parsed in parallel, nothing is
           stored on disk: EinheitenSolar_*, EinheitenWind, EinheitenStromSpeicher_*, EinheitenVerbrennung (gas, coal,
           lignite, oil, waste), EinheitenBiomasse, EinheitenKernkraft, EinheitenWasser, EinheitenGeothermie...
           (+ Katalogwerte for the code lists). The download server throttles one stream to about 0.7 MB/s, so
           members are fetched on many connections at once.
  Method   Each unit has a commissioning date (Inbetriebnahmedatum), an optional final shutdown date
           (DatumEndgueltigeStilllegung) and a net rated power (Nettonennleistung, kW; gross = Bruttoleistung). Capacity at the
           end of month M = units commissioned up to M minus units finally shut down up to M. 'In Betrieb' and
           'Vorübergehend stillgelegt' units are installed capacity; 'In Planung' units are not; 'Endgültig stillgelegt'
           units count only if their shutdown date is given (the rest are listed on 'Excluded', never dated by guess).
           Units deleted from the register (wrong registrations) are not in it.
  Caveats  * The register is the stock TODAY: a unit registered late appears in the month it was commissioned, so the latest
             two months are provisional (flag) and every run rebuilds the whole history - there is no vintage.
           * Solar_MW = net rated (AC / inverter) power, Solar_DC_MW = gross (module) power, both from MaStR.
           * Storage_MW = battery power; PumpedStorage_MW from the storage and hydro files; Gas includes other gases
             (blast furnace, coke oven, mine gas); Coal = hard coal + lignite (Lignite_MW is the subset); Oil = all oil-fired units incl.
             emergency-power and small CHP units (Oil_ge10MW_MW = units of 10 MW and over only, closer to ENTSO-E's reporting); Other = non-biogenic waste, geothermal, heat.
           * Germany only (Land = Germany); ENTSO-E's DE-LU zone also holds Luxembourg (about 1.2 GW).

Output   output/Data and Chart Outputs/germany_power_capacity_monthly.xlsx: 'Monthly' (standard capacity columns + memo columns),
         'Detail' (by MaStR category), 'Cross-check' (1 Jan vs ENTSO-E annual A68 and, for renewables/batteries, Energy-Charts
         (Fraunhofer ISE), reference only), 'Excluded', 'Releases'.
Incremental: the export file name carries its date; a run with the same file name as 'Releases' is skipped (use --force).
Usage    python3 GERMANY_MASTR_CAPACITY_MONTHLY.py [--out PATH] [--workers N] [--force] [--only-members REGEX]
"""
import argparse
import codecs
import collections
import concurrent.futures as cf
import io
import os
import re
import sys
import time
import zipfile

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(os.path.dirname(HERE)))
import xlsx_notes  # noqa: E402

OUT = os.path.join("output", "Data and Chart Outputs", "germany_power_capacity_monthly.xlsx")
ENTSOE = os.path.join("output", "Data and Chart Outputs", "germany_power_capacity.xlsx")
PAGE = "https://www.marktstammdatenregister.de/MaStR/Datendownload"
EC_URL = "https://api.energy-charts.info/installed_power?country=de&time_step=monthly&installation_decommission=false"
UA = {"User-Agent": "Mozilla/5.0 (research; energy data)"}
START = pd.Timestamp("2015-01-01")
PROVISIONAL_MONTHS = 2
BASE = 1950 * 12          # month index = year*12 + month-1; anything before 1950 is clamped

MEMBER_RE = re.compile(r"^(Einheiten(Solar|Wind|StromSpeicher|Verbrennung|Biomasse|Kernkraft|Wasser|GeothermieGrubengasDruckentspannung))(_\d+)?\.xml$")
RECORD_TAG = {"Solar": "EinheitSolar", "Wind": "EinheitWind", "StromSpeicher": "EinheitStromSpeicher", "Verbrennung": "EinheitVerbrennung",
              "Biomasse": "EinheitBiomasse", "Kernkraft": "EinheitKernkraft", "Wasser": "EinheitWasser",
              "GeothermieGrubengasDruckentspannung": "EinheitGeothermieGrubengasDruckentspannung"}
FIELDS = ("Inbetriebnahmedatum", "DatumEndgueltigeStilllegung", "EinheitBetriebsstatus", "Energietraeger", "Nettonennleistung",
          "Bruttoleistung", "Land", "Technologie", "WindAnLandOderAufSee", "ArtDerWasserkraftanlage")
FIELD_RE = re.compile(r"<(%s)>([^<]*)<" % "|".join(FIELDS))
# Energietraeger (Verbrennung) catalogue ids -> detail key
VERBRENNUNG = {"2407": "Coal (hard)", "2408": "Lignite", "2409": "Oil", "2410": "Gas (natural)", "2411": "Gas (other gases)",
               "2412": "Other (non-biogenic waste)", "2413": "Other (heat)"}
DETAIL_TO_STD = {"Solar": "Solar", "Wind onshore": "Wind", "Wind offshore": "Wind", "Hydro": "Hydro", "Bioenergy": "Bioenergy",
                 "Nuclear": "Nuclear", "Coal (hard)": "Coal", "Lignite": "Coal", "Oil (>= 10 MW)": "Oil", "Oil (< 10 MW)": "Oil", "Gas (natural)": "Gas",
                 "Gas (other gases)": "Gas", "Other (non-biogenic waste)": "Other", "Other (heat)": "Other",
                 "Other (geothermal)": "Other", "Other (verbrennung, unknown fuel)": "Other"}
STD = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]


def log(*a):
    print(*a, flush=True)


# ---------------------------------------------------------------- HTTP with retries / zip over range requests
def rget(url, tries=8, **kw):
    for k in range(tries):
        try:
            r = requests.get(url, timeout=(30, 180), headers={**UA, **kw.pop("headers", {})}, **kw)
            if r.status_code < 500 and r.status_code != 429:
                return r
        except Exception as e:  # noqa: BLE001
            log("   retry", k, repr(e)[:90])
        time.sleep(4 * (k + 1))
    raise RuntimeError("giving up on " + url)


class RangeFile(io.RawIOBase):
    def __init__(self, url):
        r = rget(url, headers={"Range": "bytes=0-0"})
        self.size = int(r.headers["content-range"].split("/")[1])
        self.pos = 0
        self.url = url

    def seekable(self): return True
    def readable(self): return True
    def tell(self): return self.pos

    def seek(self, off, whence=0):
        self.pos = off if whence == 0 else self.pos + off if whence == 1 else self.size + off
        return self.pos

    def read(self, n=-1):
        if n < 0:
            n = self.size - self.pos
        n = min(n, self.size - self.pos)
        if n <= 0:
            return b""
        d = rget(self.url, headers={"Range": f"bytes={self.pos}-{self.pos + n - 1}"}).content
        self.pos += len(d)
        return d

    def readinto(self, b):
        d = self.read(len(b))
        b[:len(d)] = d
        return len(d)


def open_zip(url):
    return zipfile.ZipFile(io.BufferedReader(RangeFile(url), buffer_size=8 << 20))


def export_url():
    html = rget(PAGE).text
    m = re.search(r'href="(https://download\.marktstammdatenregister\.de/Gesamtdatenexport[^"]+\.zip)"', html)
    if not m:
        raise RuntimeError("no Gesamtdatenexport link on the Datendownload page")
    return m.group(1)


# ---------------------------------------------------------------- catalogue
def decode(b):
    return b.decode("utf-16", "replace") if b[:2] in (b"\xff\xfe", b"\xfe\xff") else b.decode("utf-8", "replace")


def read_catalogue(z):
    with z.open("Katalogwerte.xml") as f:
        kw = decode(f.read())
    with z.open("Katalogkategorien.xml") as f:
        kk = decode(f.read())
    cats = dict(re.findall(r"<Id>(\d+)</Id>\s*<Name>(.*?)</Name>", kk, re.S))
    vals = {}
    for e in re.findall(r"<Katalogwert>(.*?)</Katalogwert>", kw, re.S):
        g = dict(re.findall(r"<(\w+)>(.*?)</\1>", e, re.S))
        vals[g["Id"]] = (g.get("Wert", ""), cats.get(g.get("KatalogKategorieId"), "?"))
    return vals


def catalogue_maps(vals):
    """Id sets resolved by NAME so a renumbered catalogue does not silently misclassify."""
    onshore = {i for i, (w, c) in vals.items() if "wind" in c.lower() and ("land" in w.lower())}
    offshore = {i for i, (w, c) in vals.items() if "wind" in c.lower() and ("see" in w.lower().split() or w.lower().endswith("see"))}
    pumped_hydro = {i for i, (w, c) in vals.items() if "pumpspeicher" in w.lower()}
    battery = {i for i, (w, c) in vals.items() if c == "StromspeicherOhneWasserkraft" and w == "Batterie"}
    other_storage = {i: w for i, (w, c) in vals.items() if c == "StromspeicherOhneWasserkraft" and w != "Batterie"}
    return onshore, offshore, pumped_hydro, battery, other_storage


# ---------------------------------------------------------------- worker: stream + parse one member
def midx(s):
    try:
        y, m = int(s[:4]), int(s[5:7])
    except ValueError:
        return None
    return max(y * 12 + m - 1, BASE)


def num(s):
    try:
        return float(s)
    except (TypeError, ValueError):
        return 0.0


def parse_member(args):
    url, name, kind, maps = args
    onshore, offshore, pumped_hydro, battery, other_storage = maps
    tag = RECORD_TAG[kind]
    close = "</%s>" % tag
    for attempt in range(4):
        t0 = time.time()
        acc = collections.defaultdict(lambda: [0.0, 0.0, 0])      # (detail, 'a'|'r', idx) -> [net kW, gross kW, n]
        ex = collections.defaultdict(lambda: [0.0, 0])            # reason, detail -> [kW, n]
        n = 0
        try:
            z = open_zip(url)
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
                        d = dict(FIELD_RE.findall(rec))
                        if d.get("Land", "84") != "84":
                            continue
                        key = classify(kind, d, onshore, offshore, pumped_hydro, battery, other_storage)
                        net, gross = num(d.get("Nettonennleistung")), num(d.get("Bruttoleistung"))
                        st = d.get("EinheitBetriebsstatus")
                        start, end = midx(d.get("Inbetriebnahmedatum", "")), midx(d.get("DatumEndgueltigeStilllegung", ""))
                        if st == "31":
                            e = ex[("planned (In Planung)", key)]; e[0] += net; e[1] += 1
                            continue
                        if st == "38" and end is None:
                            e = ex[("finally shut down, no shutdown date", key)]; e[0] += net; e[1] += 1
                            continue
                        if start is None:
                            e = ex[("no commissioning date", key)]; e[0] += net; e[1] += 1
                            continue
                        a = acc[(key, "a", start)]; a[0] += net; a[1] += gross; a[2] += 1
                        if end is not None:
                            r = acc[(key, "r", max(end, start))]; r[0] += net; r[1] += gross; r[2] += 1
            return name, dict(acc), dict(ex), n, round(time.time() - t0, 1)
        except Exception as e:  # noqa: BLE001
            log(f"   {name}: attempt {attempt + 1} failed ({repr(e)[:100]})")
            time.sleep(10)
    raise RuntimeError("member failed: " + name)


def classify(kind, d, onshore, offshore, pumped_hydro, battery, other_storage):
    if kind == "Solar":
        return "Solar"
    if kind == "Wind":
        w = d.get("WindAnLandOderAufSee")
        return "Wind offshore" if w in offshore else "Wind onshore"
    if kind == "Wasser":
        return "Pumped storage" if d.get("ArtDerWasserkraftanlage") in pumped_hydro else "Hydro"
    if kind == "Biomasse":
        return "Bioenergy"
    if kind == "Kernkraft":
        return "Nuclear"
    if kind == "GeothermieGrubengasDruckentspannung":
        return "Other (geothermal)"
    if kind == "Verbrennung":
        k = VERBRENNUNG.get(d.get("Energietraeger"), "Other (verbrennung, unknown fuel)")
        if k == "Oil":      # ENTSO-E's Oil is about the units of 10 MW and over: keep the split
            return "Oil (>= 10 MW)" if num(d.get("Nettonennleistung")) >= 10000 else "Oil (< 10 MW)"
        return k
    if kind == "StromSpeicher":
        t = d.get("Technologie")
        if t in battery:
            return "Battery"
        if t in pumped_hydro:
            return "Pumped storage"
        return "Other storage"
    return "Other"


# ---------------------------------------------------------------- assemble
def build_monthly(acc, last_idx):
    keys = sorted({k[0] for k in acc})
    idxs = range(START.year * 12 + START.month - 1, last_idx + 1)
    net = {k: pd.Series(0.0, index=range(BASE, last_idx + 1)) for k in keys}
    gross = {k: pd.Series(0.0, index=range(BASE, last_idx + 1)) for k in keys}
    cnt = {k: pd.Series(0.0, index=range(BASE, last_idx + 1)) for k in keys}
    for (k, kind, i), (n_, g_, c_) in acc.items():
        if i > last_idx:
            continue
        s = 1 if kind == "a" else -1
        net[k].loc[i] += s * n_
        gross[k].loc[i] += s * g_
        cnt[k].loc[i] += s * c_
    dates = [pd.Timestamp(year=i // 12, month=i % 12 + 1, day=1) for i in idxs]
    det = pd.DataFrame({k: net[k].cumsum().loc[list(idxs)].values / 1000.0 for k in keys}, index=dates)
    detg = pd.DataFrame({k: gross[k].cumsum().loc[list(idxs)].values / 1000.0 for k in keys}, index=dates)
    detn = pd.DataFrame({k: cnt[k].cumsum().loc[list(idxs)].values for k in keys}, index=dates)
    for d in (det, detg, detn):
        d.index.name = "date"
    return det, detg, detn


def standard(det, detg):
    m = pd.DataFrame(index=det.index)
    for f in STD:
        m[f"{f}_MW"] = sum(det[k] for k in det.columns if DETAIL_TO_STD.get(k) == f) if any(DETAIL_TO_STD.get(k) == f for k in det.columns) else 0.0
    m["Total_MW"] = m[[f"{f}_MW" for f in STD]].sum(axis=1)
    m["PumpedStorage_MW"] = det.get("Pumped storage", 0.0)
    m["Storage_MW"] = det.get("Battery", 0.0)
    m["Wind_Onshore_MW"] = det.get("Wind onshore", 0.0)
    m["Wind_Offshore_MW"] = det.get("Wind offshore", 0.0)
    m["Lignite_MW"] = det.get("Lignite", 0.0)
    m["Oil_ge10MW_MW"] = det.get("Oil (>= 10 MW)", 0.0)
    m["Solar_DC_MW"] = detg.get("Solar", 0.0)
    m["Provisional"] = ""
    m.iloc[-PROVISIONAL_MONTHS:, m.columns.get_loc("Provisional")] = "late registrations"
    return m.round(1)


def crosscheck(m):
    rows = []
    try:
        e = pd.read_excel(ENTSOE, sheet_name="Monthly", index_col=0)
        e.index = pd.to_datetime(e.index)
    except Exception as ex:  # noqa: BLE001
        log("ENTSO-E annual workbook not read:", repr(ex)[:80])
        e = pd.DataFrame()
    for ts in e.index:
        prev = ts - pd.DateOffset(months=1)         # 1 Jan = end of December
        if prev not in m.index:
            continue
        for f in STD + ["Total", "PumpedStorage", "Storage"]:
            c = f"{f}_MW"
            if c not in e.columns:
                continue
            o, t = float(m.loc[prev, c]), float(e.loc[ts, c])
            rows.append({"1 Jan": ts.date(), "Fuel": f, "MaStR end of previous month (MW)": round(o), "ENTSO-E A68 (MW)": round(t),
                         "Difference (MW)": round(o - t), "Difference (%)": round((o - t) / t * 100, 1) if t else None})
    out = pd.DataFrame(rows)
    # Energy-Charts (Fraunhofer ISE) monthly Germany: reference only, not used in the series
    try:
        d = rget(EC_URL).json()
        idx = {t: i for i, t in enumerate(d["time"])}
        ref = []
        for ts in e.index if len(e) else []:
            t = f"{ts.month:02d}.{ts.year}"
            pt = f"{(ts - pd.DateOffset(months=1)).month:02d}.{(ts - pd.DateOffset(months=1)).year}"
            if pt not in idx or (ts - pd.DateOffset(months=1)) not in m.index:
                continue
            ec = {s["name"]: s["data"][idx[pt]] for s in d["production_types"]}
            prev = ts - pd.DateOffset(months=1)
            for lab, ours, key in [("Wind onshore", m.loc[prev, "Wind_Onshore_MW"], "Wind onshore"), ("Wind offshore", m.loc[prev, "Wind_Offshore_MW"], "Wind offshore"),
                                   ("Solar (AC vs net rated)", m.loc[prev, "Solar_MW"], "Solar AC"), ("Solar (DC vs gross)", m.loc[prev, "Solar_DC_MW"], "Solar DC"),
                                   ("Bioenergy", m.loc[prev, "Bioenergy_MW"], "Biomass"), ("Battery power", m.loc[prev, "Storage_MW"], "Battery storage (power)")]:
                if key in ec and ec[key]:
                    ref.append({"1 Jan": ts.date(), "Item": lab, "MaStR (MW)": round(ours), "Energy-Charts (MW), FALLBACK reference": round(ec[key] * 1000),
                                "Difference (%)": round((ours - ec[key] * 1000) / (ec[key] * 1000) * 100, 1)})
        return out, pd.DataFrame(ref)
    except Exception as ex:  # noqa: BLE001
        log("Energy-Charts reference not read:", repr(ex)[:80])
        return out, pd.DataFrame()


def notes(url, last_month, n_units, secs):
    return ["UNITS", "MW installed at the END of each month (row date = 1st of the month it closes). Total_MW excludes PumpedStorage_MW and Storage_MW (battery power), as in the ENTSO-E annual workbook. Lignite_MW is part of Coal_MW; Solar_DC_MW is a memo (module power).",
            "", "SOURCE", "Bundesnetzagentur Marktstammdatenregister (MaStR) Gesamtdatenexport, www.marktstammdatenregister.de/MaStR/Datendownload, Datenlizenz Deutschland - Namensnennung 2.0. Export read: " + url.rsplit("/", 1)[1] + f" (last complete month {last_month:%b %Y}; {n_units:,} units parsed in {secs:.0f} s).",
            "", "METHOD", "Month-end capacity = sum of Nettonennleistung of units commissioned up to that month minus units finally shut down up to it. In Betrieb and Vorübergehend stillgelegt count as installed; In Planung does not; Endgültig stillgelegt only with a shutdown date (the rest on 'Excluded'). Germany only.",
            "The register is the stock today: late registrations land in the month of commissioning, so the last two months are provisional and the whole history is rebuilt every run (no vintage kept).",
            "", "COVERAGE", "All fuels, from Jan 2015. Oil includes emergency-power diesel generators registered as units. Other = non-biogenic waste, geothermal and heat. Gas includes other gases (blast furnace, coke oven, mine gas).",
            "", "VALIDATION", "'Cross-check' compares end-December with the 1 January ENTSO-E annual values (DE-LU zone, so ENTSO-E also holds Luxembourg) and, for renewables and batteries, with Energy-Charts (Fraunhofer ISE) monthly figures, reference only."]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--workers", type=int, default=14)
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--only-members", default="")
    a = ap.parse_args()
    url = export_url()
    fname = url.rsplit("/", 1)[1]
    log("export:", url)
    try:
        rel = pd.read_excel(a.out, sheet_name="Releases", index_col=0)
        if not a.force and fname in set(rel["File"]) and not a.only_members:
            log("same export already processed:", fname)
            return
    except Exception:  # noqa: BLE001
        rel = pd.DataFrame()
    t0 = time.time()
    z = open_zip(url)
    vals = read_catalogue(z)
    maps = catalogue_maps(vals)
    log("catalogue: %d values; onshore ids %s offshore ids %s pumped-storage ids %s battery ids %s other storage %s" % (
        len(vals), sorted(maps[0]), sorted(maps[1]), sorted(maps[2]), sorted(maps[3]), maps[4]))
    tasks = []
    for i in z.infolist():
        m = MEMBER_RE.match(i.filename)
        if m and (not a.only_members or re.search(a.only_members, i.filename)):
            tasks.append((i.compress_size, (url, i.filename, m.group(2), maps)))
    tasks.sort(key=lambda t: -t[0])
    log(len(tasks), "members,", round(sum(t[0] for t in tasks) / 1e9, 2), "GB compressed")
    acc = collections.defaultdict(lambda: [0.0, 0.0, 0])
    exc = collections.defaultdict(lambda: [0.0, 0])
    units = 0
    with cf.ProcessPoolExecutor(max_workers=a.workers) as pool:
        futs = [pool.submit(parse_member, t[1]) for t in tasks]
        for k, fu in enumerate(cf.as_completed(futs), 1):
            name, ac, ex, n, secs = fu.result()
            units += n
            for key, v in ac.items():
                s = acc[key]; s[0] += v[0]; s[1] += v[1]; s[2] += v[2]
            for key, v in ex.items():
                s = exc[key]; s[0] += v[0]; s[1] += v[1]
            log(f"[{k}/{len(tasks)}] {name}: {n:,} units in {secs}s")
    export_day = pd.Timestamp(re.search(r"_(\d{8})_", fname).group(1))
    last_complete = (export_day.replace(day=1) - pd.DateOffset(months=1))
    last_idx = last_complete.year * 12 + last_complete.month - 1
    det, detg, detn = build_monthly(dict(acc), last_idx)
    m = standard(det, detg)
    xc, ec = crosscheck(m)
    excl = pd.DataFrame([{"Reason": k[0], "Category": k[1], "MW": round(v[0] / 1000, 1), "Units": v[1]} for k, v in sorted(exc.items())])
    rel_new = pd.DataFrame([{"File": fname, "Export date": export_day.date(), "Units parsed": units, "Seconds": round(time.time() - t0),
                             "Last complete month": last_complete.date()}], index=[pd.Timestamp.now(tz="UTC").strftime("%Y-%m-%d %H:%M")])
    rel_all = pd.concat([rel, rel_new]) if len(rel) else rel_new
    sheets = {"Monthly": m, "Detail": det.round(1), "Detail units": detn.astype(int), "Cross-check": xc.set_index("1 Jan") if len(xc) else xc,
              "Cross-check Energy-Charts": ec.set_index("1 Jan") if len(ec) else ec, "Excluded": excl, "Releases": rel_all.tail(60)}
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    lines = notes(url, last_complete, units, time.time() - t0)
    xlsx_notes.write_workbook(a.out, sheets, lines, {"UNITS", "SOURCE", "METHOD", "COVERAGE", "VALIDATION"})
    log("wrote", a.out, "last month", last_complete.date())
    if len(xc):
        log(xc[xc["Fuel"].isin(["Solar", "Wind", "Gas", "Coal", "Total"])].tail(30).to_string())


if __name__ == "__main__":
    main()
