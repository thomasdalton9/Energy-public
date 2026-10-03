"""
Keyed probe for the Europe master: ENTSO-E and GIE (AGSI+/ALSI) with the real
keys, and the accuracy test that decides which source is primary.

Reads ENTSOE_API_KEY / GIE_API_KEY from the environment (GitHub secrets) and
never prints them - every logged string goes through redact().

1. Key check (set / not set, ENTSO-E auth works).
2. ENTSO-E coverage: actual generation per type (A75) for ~20 bidding zones,
   last 3 days: psr types returned, resolution, lag of the last point.
   GB is expected to be empty (no GB data after Brexit).
3. ACCURACY: one settled week (2026-05-04..10, UTC) of daily-total generation by
   fuel group (gas, coal, nuclear, wind, solar, hydro) from ENTSO-E vs
   Energy-Charts for DE/FR/ES/IT/PL/NL/BE, and vs the national raw feed where
   we have one (DE SMARD, FR RTE eCO2mix, ES REE). Prints GWh and % difference
   against ENTSO-E.
4. Other ENTSO-E documents: day-ahead price, actual load, installed capacity,
   hydro reservoir filling.
5. GIE with key: listing, long-history call size/time, ALSI.

Usage: python3 EU_KEYED_PROBE.py   (writes eu_keyed_probe_report.json)
"""
import json
import os
import re
import statistics
import sys
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone

import requests

try:  # local runs: api_keys.py (gitignored); Actions: environment secrets
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
    import api_keys  # noqa: F401
except Exception:
    api_keys = None


def key(name):
    return os.environ.get(name) or (getattr(api_keys, name, "") if api_keys else "") or ""


ENTSOE = key("ENTSOE_API_KEY")
GIE = key("GIE_API_KEY")
SECRETS = [s for s in (ENTSOE, GIE) if s]


def redact(text):
    text = str(text)
    for s in SECRETS:
        text = text.replace(s, "***")
    return re.sub(r"securityToken=[^&\s]+", "securityToken=***", text)


def log(*a):
    print(redact(" ".join(str(x) for x in a)), flush=True)


UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
RESULTS = {}

ENTSOE_URL = "https://web-api.tp.entsoe.eu/api"
ZONES = {
    "DE-LU": "10Y1001A1001A82H", "FR": "10YFR-RTE------C", "ES": "10YES-REE------0", "IT-North": "10Y1001A1001A73I",
    "PL": "10YPL-AREA-----S", "NL": "10YNL----------L", "BE": "10YBE----------2", "AT": "10YAT-APG------L",
    "CH": "10YCH-SWISSGRIDZ", "CZ": "10YCZ-CEPS-----N", "HU": "10YHU-MAVIR----U", "RO": "10YRO-TEL------P",
    "PT": "10YPT-REN------W", "GR": "10YGR-HTSO-----Y", "DK1": "10YDK-1--------W", "SE3": "10Y1001A1001A46L",
    "NO2": "10YNO-2--------T", "FI": "10YFI-1--------U", "IE(SEM)": "10Y1001A1001A59C", "GB": "10YGB----------A",
}
GROUPS = {  # ENTSO-E psrType -> fuel group
    "B04": "gas", "B05": "coal", "B02": "coal", "B14": "nuclear", "B16": "solar", "B18": "wind", "B19": "wind",
    "B11": "hydro", "B12": "hydro",
}
ORDER = ["gas", "coal", "nuclear", "wind", "solar", "hydro"]
W0 = datetime.strptime(os.environ.get("ACC_START") or "2026-05-04", "%Y-%m-%d").replace(tzinfo=timezone.utc)
W1 = W0 + timedelta(days=7)  # [W0, W1)


def stamp(d):
    return d.strftime("%Y%m%d%H%M")


# ------------------------------------------------------------- ENTSO-E -----
def entsoe_get(params, retries=2):
    p = dict(params, securityToken=ENTSOE)
    for _ in range(retries + 1):
        try:
            r = requests.get(ENTSOE_URL, params=p, headers=UA, timeout=(10, 90))
        except requests.RequestException as e:
            return None, redact(f"{type(e).__name__}: {e}")[:200]
        if r.status_code == 429:
            time.sleep(15)
            continue
        return r, None
    return r, None


def strip(tag):
    return tag.split("}")[-1]


def parse_series(xml_text, want_generation=True):
    """-> {psr: {'mwh': float, 'points': n, 'res': minutes, 'last': datetime}} ; or ('ack', reason)."""
    root = ET.fromstring(xml_text)
    if strip(root.tag).startswith("Acknowledgement"):
        reason = " ".join(e.text or "" for e in root.iter() if strip(e.tag) == "text")
        return "ack", reason.strip()
    out = {}
    for ts in root:
        if strip(ts.tag) != "TimeSeries":
            continue
        kids = {strip(c.tag): c for c in ts}
        # Direction tags are not reliable across TSOs (some publish generation as in-, others as out-BiddingZone),
        # so keep both and key by psr:direction; fuel groups only use generation psr types.
        direction = "in" if "inBiddingZone_Domain.mRID" in kids else "out" if "outBiddingZone_Domain.mRID" in kids else ""
        psr = None
        for e in ts.iter():
            if strip(e.tag) == "psrType":
                psr = e.text
        psr = psr or "total"
        if want_generation:
            psr = f"{psr}:{direction}"
        for per in ts:
            if strip(per.tag) != "Period":
                continue
            ti = {strip(c.tag): c.text for c in next(c for c in per if strip(c.tag) == "timeInterval")}
            res = next(c.text for c in per if strip(c.tag) == "resolution")
            m = re.match(r"PT(\d+)M", res)
            step = int(m.group(1)) if m else (60 if res == "PT60M" else 1440 if res == "P1D" else 0)
            if not step:
                step = {"P1Y": 525600, "P7D": 10080, "P1M": 43200}.get(res, 60)
            start = datetime.strptime(ti["start"], "%Y-%m-%dT%H:%MZ").replace(tzinfo=timezone.utc)
            end = datetime.strptime(ti["end"], "%Y-%m-%dT%H:%MZ").replace(tzinfo=timezone.utc)
            n = int((end - start).total_seconds() // 60 // step)
            pts = {int(next(c.text for c in p if strip(c.tag) == "position")): float(next(c.text for c in p if strip(c.tag) in ("quantity", "price.amount")))
                   for p in per if strip(p.tag) == "Point"}
            vals, last = [], None
            for pos in range(1, n + 1):  # A03 curves omit repeated points: fill forward
                if pos in pts:
                    last = pts[pos]
                vals.append(last or 0.0)
            d = out.setdefault(psr, {"mwh": 0.0, "points": 0, "res": step, "last": start})
            d["mwh"] += sum(vals) * step / 60.0
            d["points"] += len(pts)
            d["last"] = max(d["last"], start + timedelta(minutes=step * max(pts) if pts else 0))
    return out


def entsoe_gen(zone, d0, d1):
    r, err = entsoe_get({"documentType": "A75", "processType": "A16", "in_Domain": zone, "periodStart": stamp(d0), "periodEnd": stamp(d1)})
    if r is None:
        return "err", err
    if r.status_code != 200:
        return "http", f"{r.status_code} {redact(r.text[:200])}"
    return parse_series(r.text)


def entsoe_coverage():
    now = datetime.now(timezone.utc).replace(minute=0, second=0, microsecond=0)
    log("\n== ENTSO-E actual generation per type, last 3 days ==")
    cov = {}
    for name, z in ZONES.items():
        res = entsoe_gen(z, now - timedelta(days=3), now)
        if isinstance(res, tuple):
            cov[name] = f"{res[0]}: {res[1][:120]}"
        elif not res:
            cov[name] = "empty"
        else:
            last = max(v["last"] for v in res.values())
            gen = sorted({k.split(":")[0] for k in res if k.split(":")[0] in GROUPS})
            dirs = sorted({k.split(":")[1] for k in res if k.split(":")[0] in GROUPS})
            cov[name] = (f"types={len({k.split(':')[0] for k in res})} fuel-types={gen} dir={dirs} "
                         f"res={sorted({v['res'] for v in res.values()})}min last={last:%Y-%m-%d %H:%M} lag_h={(now - last).total_seconds() / 3600:.0f}")
        log(f"  {name}: {cov[name]}")
        time.sleep(0.4)
    RESULTS["entsoe_coverage"] = cov
    log("  GB history check 2019:")
    res = entsoe_gen(ZONES["GB"], datetime(2019, 5, 4, tzinfo=timezone.utc), datetime(2019, 5, 6, tzinfo=timezone.utc))
    log("   ", res if isinstance(res, tuple) else f"types={len(res)}")
    log("  FR full-year request (size/time):")
    t0 = time.time()
    r, err = entsoe_get({"documentType": "A75", "processType": "A16", "in_Domain": ZONES["FR"], "periodStart": "202501010000", "periodEnd": "202601010000"})
    log("   ", err or f"{r.status_code} {len(r.content) / 1e6:.1f}MB {time.time() - t0:.1f}s")


# ----------------------------------------------------------- comparators ---
def group_sum(d):
    g = {k: 0.0 for k in ORDER}
    for k, v in d.items():
        if k in g:
            g[k] += v
    return g


def entsoe_week(zone):
    """Day-by-day requests (a 7-day request came back empty/errored), summed to fuel groups."""
    g = {k: 0.0 for k in ORDER}
    days_ok = 0
    d = W0
    while d < W1:
        res = entsoe_gen(zone, d, d + timedelta(days=1))
        if isinstance(res, tuple):
            log(f"    ENTSO-E {zone} {d:%Y-%m-%d}: {res[0]} {str(res[1])[:200]}")
        else:
            days_ok += 1
            for k, v in res.items():
                psr = k.split(":")[0]
                if psr in GROUPS:
                    g[GROUPS[psr]] += v["mwh"]
        d += timedelta(days=1)
        time.sleep(0.3)
    return g if days_ok == 7 else None


def energy_charts_week(cc):
    d1 = (W1 - timedelta(days=1)).strftime("%Y-%m-%d")
    for _ in range(3):
        r = requests.get("https://api.energy-charts.info/public_power", params={"country": cc, "start": W0.strftime("%Y-%m-%d"), "end": d1}, headers=UA, timeout=(10, 90))
        if r.status_code == 429:
            time.sleep(20)
            continue
        break
    if r.status_code != 200:
        return None
    d = r.json()
    ts = d["unix_seconds"]
    step_h = statistics.median(b - a for a, b in zip(ts, ts[1:])) / 3600.0
    m = {"Fossil gas": "gas", "Fossil hard coal": "coal", "Fossil brown coal / lignite": "coal", "Nuclear": "nuclear",
         "Wind onshore": "wind", "Wind offshore": "wind", "Solar": "solar", "Hydro Run-of-River": "hydro", "Hydro water reservoir": "hydro"}
    g = {k: 0.0 for k in ORDER}
    for s in d["production_types"]:
        name = s["name"] if isinstance(s["name"], str) else s["name"].get("en")
        if name in m:
            g[m[name]] += sum(x for x in s["data"] if x is not None) * step_h
    return g


def smard_week():
    ids = {1223: "coal", 4069: "coal", 1224: "nuclear", 1225: "wind", 4067: "wind", 4068: "solar", 4071: "gas", 1226: "hydro"}
    g = {k: 0.0 for k in ORDER}
    a, b = int(W0.timestamp() * 1000), int(W1.timestamp() * 1000)
    for fid, grp in ids.items():
        idx = requests.get(f"https://www.smard.de/app/chart_data/{fid}/DE/index_hour.json", headers=UA, timeout=30).json()["timestamps"]
        for i, t in enumerate(idx):
            nxt = idx[i + 1] if i + 1 < len(idx) else 10 ** 15
            if t < b and nxt > a:
                ser = requests.get(f"https://www.smard.de/app/chart_data/{fid}/DE/{fid}_DE_hour_{t}.json", headers=UA, timeout=30).json()["series"]
                g[grp] += sum(v for ms, v in ser if v is not None and a <= ms < b)
    return g


def rte_week():
    ds = "eco2mix-national-cons-def" if W1 <= datetime(2026, 6, 30, tzinfo=timezone.utc) else "eco2mix-national-tr"
    log(f"    RTE dataset: {ds}")
    base = f"https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/{ds}/records"
    cols = {"gaz": "gas", "charbon": "coal", "nucleaire": "nuclear", "eolien": "wind", "solaire": "solar", "hydraulique": "hydro"}
    g = {k: 0.0 for k in ORDER}
    off, n = 0, 0
    where = f"date_heure >= '{W0:%Y-%m-%dT%H:%M:%SZ}' and date_heure < '{W1:%Y-%m-%dT%H:%M:%SZ}'"
    while True:
        r = requests.get(base, params={"where": where, "limit": 100, "offset": off, "order_by": "date_heure"}, headers=UA, timeout=60)
        if r.status_code != 200:
            return None
        rows = r.json()["results"]
        for row in rows:
            if row.get("nature", "").lower().startswith("donn") and row.get("gaz") is not None:
                n += 1
                for c, grp in cols.items():
                    g[grp] += (row.get(c) or 0) * 0.5
        off += 100
        if len(rows) < 100:
            break
    log(f"    RTE rows used: {n}")
    return g


def ree_week():
    r = requests.get("https://apidatos.ree.es/en/datos/generacion/estructura-generacion",
                     params={"start_date": f"{W0:%Y-%m-%d}T00:00", "end_date": f"{(W1 - timedelta(days=1)):%Y-%m-%d}T23:59", "time_trunc": "day"}, headers=UA, timeout=60)
    if r.status_code != 200:
        return None
    m = {"Combined cycle": "gas", "Coal": "coal", "Nuclear": "nuclear", "Wind": "wind", "Solar PV": "solar", "Hydro": "hydro"}
    g = {k: 0.0 for k in ORDER}
    for s in r.json()["included"]:
        t = s["attributes"]["title"]
        if t in m:
            g[m[t]] += sum(v["value"] for v in s["attributes"]["values"])
    return g


def show(label, base, other):
    if base is None or other is None:
        log(f"    {label}: unavailable")
        return
    parts = []
    for k in ORDER:
        b, o = base[k], other[k]
        parts.append(f"{k} {o / 1000:.0f}GWh ({'n/a' if b < 1 else f'{(o - b) / b * 100:+.1f}%'})")
    log(f"    {label}: " + "; ".join(parts))


def accuracy():
    log(f"\n== ACCURACY: week {W0:%Y-%m-%d}..{(W1 - timedelta(days=1)):%Y-%m-%d} (GWh, % vs ENTSO-E) ==")
    ec_codes = {"DE-LU": "de", "FR": "fr", "ES": "es", "IT-North": "it", "PL": "pl", "NL": "nl", "BE": "be", "AT": "at", "CZ": "cz", "PT": "pt"}
    raw = {"DE-LU": ("SMARD", smard_week), "FR": ("RTE eCO2mix", rte_week), "ES": ("REE REData", ree_week)}
    out = {}
    for zone, cc in ec_codes.items():
        base = entsoe_week(ZONES[zone])  # EIC code, not the label
        log(f"  {zone}: ENTSO-E " + ("unavailable" if base is None else "; ".join(f"{k} {base[k] / 1000:.0f}GWh" for k in ORDER)))
        try:
            show("Energy-Charts", base, energy_charts_week(cc))
        except Exception as e:
            log(f"    Energy-Charts failed: {type(e).__name__}: {e}")
        if zone in raw:
            try:
                show(raw[zone][0], base, raw[zone][1]())
            except Exception as e:
                log(f"    {raw[zone][0]} failed: {type(e).__name__}: {e}")
        out[zone] = base
        time.sleep(5)
    log("  note: IT-North is one of six Italian zones, so Energy-Charts (whole Italy) is expected to differ; read it as a coverage check only.")
    RESULTS["accuracy_entsoe_gwh"] = {z: (None if b is None else {k: round(v / 1000, 1) for k, v in b.items()}) for z, b in out.items()}


def other_docs():
    log("\n== ENTSO-E other documents ==")
    de, now = ZONES["DE-LU"], datetime.now(timezone.utc).replace(minute=0, second=0, microsecond=0)
    tests = [
        ("day-ahead price DE-LU", {"documentType": "A44", "in_Domain": de, "out_Domain": de, "periodStart": stamp(now - timedelta(days=3)), "periodEnd": stamp(now)}),
        ("actual load DE-LU", {"documentType": "A65", "processType": "A16", "outBiddingZone_Domain": de, "periodStart": stamp(now - timedelta(days=3)), "periodEnd": stamp(now)}),
        ("installed capacity ES (yearly)", {"documentType": "A68", "processType": "A33", "in_Domain": ZONES["ES"], "periodStart": "202501010000", "periodEnd": "202601010000"}),
        ("hydro reservoir filling NO2 (weekly)", {"documentType": "A72", "processType": "A16", "in_Domain": ZONES["NO2"], "periodStart": stamp(now - timedelta(days=60)), "periodEnd": stamp(now)}),
        ("hydro reservoir filling ES (weekly)", {"documentType": "A72", "processType": "A16", "in_Domain": ZONES["ES"], "periodStart": stamp(now - timedelta(days=60)), "periodEnd": stamp(now)}),
        ("cross-border physical flow FR->DE", {"documentType": "A11", "in_Domain": ZONES["DE-LU"], "out_Domain": ZONES["FR"], "periodStart": stamp(now - timedelta(days=2)), "periodEnd": stamp(now)}),
    ]
    for name, p in tests:
        r, err = entsoe_get(p)
        if r is None:
            log(f"  {name}: ERR {err}")
        elif r.status_code != 200:
            log(f"  {name}: {r.status_code} {redact(r.text[:160])}")
        else:
            try:
                res = parse_series(r.text, want_generation=False)
                if isinstance(res, tuple):
                    log(f"  {name}: no data ({res[1][:100]})")
                else:
                    log(f"  {name}: series={ {k: (v['points'], v['res']) for k, v in list(res.items())[:6]} }")
            except Exception as e:
                log(f"  {name}: 200 but parse failed {type(e).__name__}: {e}; head={redact(r.text[:200])}")
        time.sleep(0.4)


# ------------------------------------------------------------------ GIE ----
def gie():
    log("\n== GIE with key ==")
    h = dict(UA, **{"x-key": GIE})

    def call(name, url, params):
        t0 = time.time()
        try:
            r = requests.get(url, params=params, headers=h, timeout=(10, 120))
        except requests.RequestException as e:
            log(f"  {name}: ERR {redact(e)[:160]}")
            return
        info = ""
        if r.status_code == 200:
            j = r.json()
            data = j.get("data", j) if isinstance(j, dict) else j
            info = f"total={j.get('total') if isinstance(j, dict) else len(j)} last_page={j.get('last_page') if isinstance(j, dict) else ''} rows={len(data) if hasattr(data, '__len__') else '?'}"
        log(f"  {name}: {r.status_code} {len(r.content) / 1e3:.0f}KB {time.time() - t0:.1f}s {info or redact(r.text[:120])}")
        return r

    call("AGSI listing", "https://agsi.gie.eu/api/about", {"show": "listing"})
    call("AGSI DE 2025-10..2026-09 size=400", "https://agsi.gie.eu/api", {"country": "DE", "from": "2025-10-01", "to": "2026-09-30", "size": 400})
    call("AGSI EU 2011..2026 size=400 page1", "https://agsi.gie.eu/api", {"continent": "eu", "from": "2011-01-01", "to": "2026-09-30", "size": 400})
    call("AGSI DE by company", "https://agsi.gie.eu/api", {"country": "DE", "from": "2026-09-25", "to": "2026-09-30", "size": 50})
    call("ALSI ES 2025-10..2026-09 size=400", "https://alsi.gie.eu/api", {"country": "ES", "from": "2025-10-01", "to": "2026-09-30", "size": 400})
    call("ALSI EU aggregate 2016..2026 size=400 page1", "https://alsi.gie.eu/api", {"continent": "eu", "from": "2016-01-01", "to": "2026-09-30", "size": 400})
    call("ALSI UK latest", "https://alsi.gie.eu/api", {"country": "GB", "from": "2026-09-25", "to": "2026-09-30", "size": 20})
    call("AGSI UK latest", "https://agsi.gie.eu/api", {"country": "GB", "from": "2026-09-25", "to": "2026-09-30", "size": 20})
    call("AGSI UA latest", "https://agsi.gie.eu/api", {"country": "UA", "from": "2026-09-25", "to": "2026-09-30", "size": 20})


def main():
    log(f"ENTSOE_API_KEY: {'set (len ' + str(len(ENTSOE)) + ')' if ENTSOE else 'NOT SET'};  GIE_API_KEY: {'set (len ' + str(len(GIE)) + ')' if GIE else 'NOT SET'}")
    parts = (os.environ.get("KEYED_PARTS") or "coverage,accuracy,other,gie").split(",")
    if ENTSOE:
        for name, fn in (("coverage", entsoe_coverage), ("accuracy", accuracy), ("other", other_docs)):
            if name not in parts:
                continue
            try:
                fn()
            except Exception as e:
                log(f"!! {fn.__name__} failed: {type(e).__name__}: {e}")
    if GIE and "gie" in parts:
        try:
            gie()
        except Exception as e:
            log(f"!! gie failed: {type(e).__name__}: {e}")
    with open("eu_keyed_probe_report.json", "w") as f:
        json.dump(json.loads(redact(json.dumps(RESULTS, default=str))), f, indent=1)


if __name__ == "__main__":
    main()
