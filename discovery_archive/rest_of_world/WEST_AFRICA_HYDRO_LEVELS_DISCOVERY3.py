"""
Follow-up 2: Bui Power Authority live reservoir JSON (data/site-content.json), the VRA home-page
widget source, Hydroweb.next HYDROWEB_LAKES_OPE items for West African reservoirs (names, data
asset), SOGEM/OMVS level links, Open-Meteo discharge sample. Writes west_africa_hydro_levels_output3.txt.
"""
import json
import re
import signal
import requests

OUT = open("west_africa_hydro_levels_output3.txt", "w", encoding="utf-8")


def log(*a):
    s = " ".join(str(x) for x in a)
    print(s, flush=True)
    OUT.write(s + "\n")
    OUT.flush()


class Hard(Exception):
    pass


def _alarm(*a):
    raise Hard("hard deadline")


signal.signal(signal.SIGALRM, _alarm)
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def get(url, **kw):
    try:
        signal.alarm(60)
        r = requests.get(url, headers=H, timeout=(10, 40), **kw)
        signal.alarm(0)
        return r
    except BaseException as e:
        signal.alarm(0)
        log("  ERR", url, type(e).__name__, str(e)[:150])
        return None


def section(t):
    log("=" * 100)
    log("##", t)


section("Bui site-content.json")
for u in ["https://www.buipower.com/data/site-content.json", "https://buipower.com/data/site-content.json"]:
    r = get(u)
    if r is not None:
        log(" ", u, r.status_code, r.headers.get("content-type"), len(r.content))
        log("  ", r.text[:2500])
        break

section("VRA home page: widget, json, script references")
r = get("https://vra.com/")
if r is not None:
    t = r.text
    for m in re.finditer(r"(?is)Live reservoir status.{0,1800}", t):
        log("  widget html:", re.sub(r"\s+", " ", m.group(0))[:1500])
        break
    for m in sorted(set(re.findall(r"[\"']([^\"']*\.(?:json|php|ashx|aspx)[^\"']*)[\"']", t))):
        log("  ref:", m[:160])
    for m in re.finditer(r"(?is)<script(?![^>]+src)[^>]*>(.*?)</script>", t):
        b = m.group(1)
        if re.search(r"(?i)reservoir|masl|water.?level|fetch\(|ajax", b):
            log("  inline script:", re.sub(r"\s+", " ", b)[:2500])
    for p in ["data/site-content.json", "data/reservoir.json", "data/water-level.json", "data/lake-level.json"]:
        rr = get("https://vra.com/" + p)
        if rr is not None:
            log(f"  probe {p}: {rr.status_code} {rr.headers.get('content-type')} {rr.text[:300]!r}")

section("Hydroweb.next lake items, West Africa")
base = "https://hydroweb.next.theia-land.fr/api/v1/rs-catalog/stac/search"
for coll in ["HYDROWEB_LAKES_OPE", "HYDROWEB_LAKES_RESEARCH", "HYDROWEB_RIVERS_OPE"]:
    for bbox in ["-5,4,15,15", "-14,4,15,17"]:
        r = get(f"{base}?collections={coll}&bbox={bbox}&limit=50")
        if r is None:
            continue
        log(f"  {coll} bbox {bbox}: {r.status_code} {len(r.content)}")
        if r.ok:
            try:
                j = r.json()
            except Exception:
                log(r.text[:300]); continue
            fs = j.get("features", [])
            log("   features:", len(fs), "matched:", j.get("numberMatched") or j.get("context"))
            for f in fs[:60]:
                p = f.get("properties", {})
                nm = p.get("title") or p.get("name") or f.get("id")
                log("    ", f.get("id"), "|", nm, "|", p.get("datetime") or (p.get("start_datetime"), p.get("end_datetime")))
            if fs:
                f0 = fs[0]
                log("  sample item keys:", list(f0.keys()), "props:", json.dumps(f0.get("properties"))[:1200])
                log("  assets:", json.dumps(f0.get("assets"))[:1500])
            break

section("SOGEM / OMVS links on levels")
for u in ["https://www.sogem-omvs.org/", "https://www.omvs.org/", "https://www.sogem-omvs.org/barrage-de-manantali/"]:
    r = get(u)
    if r is None:
        continue
    t = r.text
    for m in re.finditer(r"(?i)<a[^>]+href=[\"']([^\"']+)[\"'][^>]*>([^<]{0,80})</a>", t):
        if re.search(r"(?i)cote|niveau|hydro|bulletin|retenue|lac|remplissage|donn|stat|rapport", m.group(2) + m.group(1)):
            log("  ", u.split("/")[2], "|", m.group(2).strip()[:70], "|", m.group(1)[:160])

section("Open-Meteo discharge sample")
r = get("https://flood-api.open-meteo.com/v1/flood?latitude=6.30&longitude=0.06&daily=river_discharge&start_date=2026-09-20&end_date=2026-10-08")
if r is not None:
    log(" ", r.text[:600])
r = get("https://flood-api.open-meteo.com/v1/flood?latitude=6.30&longitude=0.06&daily=river_discharge&start_date=2000-01-01&end_date=2000-01-05")
if r is not None:
    log("  2000:", r.text[-250:])
r = get("https://flood-api.open-meteo.com/v1/flood?latitude=6.30&longitude=0.06&daily=river_discharge&start_date=1995-01-01&end_date=1995-01-05")
if r is not None:
    log("  1995:", r.text[-250:])
log("Done")
OUT.close()
