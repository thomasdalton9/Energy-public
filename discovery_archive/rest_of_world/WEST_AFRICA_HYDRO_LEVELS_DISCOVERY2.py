"""
Follow-up to WEST_AFRICA_HYDRO_LEVELS_DISCOVERY.py: (1) find the endpoint behind the live
'Reservoir water level' widgets on vra.com and buipower.com; (2) list Hydroweb.next
collections naming West African reservoirs/lakes and fetch one item; (3) Open-Meteo/GloFAS
discharge history depth; (4) ABN hydrological bulletin PDF content (Niger levels, Kainji);
(5) try to reach G-REALM text files. Writes west_africa_hydro_levels_output2.txt.
"""
import io
import re
import signal
import requests

OUT = open("west_africa_hydro_levels_output2.txt", "w", encoding="utf-8")


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


# 1. VRA / Bui widgets
for name, url in [("VRA", "https://vra.com/"), ("Bui", "https://www.buipower.com/")]:
    section(f"{name} widget endpoints")
    r = get(url)
    if r is None:
        continue
    t = r.text
    for m in re.finditer(r"(fetch|axios\.get|\$\.(?:get|ajax|getJSON)|XMLHttpRequest|\.open)\s*\(\s*[\"'`]([^\"'`]+)", t):
        log("  js call:", m.group(2)[:200])
    for m in re.finditer(r"[\"'](https?://[^\"']*(?:api|json|reservoir|level|lake|scada|data)[^\"']*)[\"']", t, re.I):
        log("  url-ish:", m.group(1)[:200])
    for m in re.finditer(r"<script[^>]+src=[\"']([^\"']+)", t):
        log("  script:", m.group(1)[:160])
    for m in re.finditer(r"(?i)(reservoir[^<]{0,60}|masl[^<]{0,60}|data-[a-z-]*(?:level|reservoir)[a-z-]*=[\"'][^\"']*)", t):
        log("  marker:", m.group(0)[:150])
    # candidate API paths
    base = r.url.rstrip("/")
    for p in ["/api/reservoir", "/api/reservoir-level", "/api/lake-level", "/api/water-level",
              "/wp-json/", "/api/live", "/api/stats", "/api/dashboard", "/api/hydro"]:
        rr = get(base + p)
        if rr is not None:
            log(f"  probe {p}: {rr.status_code} {rr.headers.get('content-type','')} {rr.text[:150]!r}")
    # inline scripts mentioning reservoir
    for m in re.finditer(r"<script(?![^>]+src)[^>]*>(.*?)</script>", t, re.S):
        body = m.group(1)
        if re.search(r"(?i)reservoir|masl|waterlevel|water_level", body):
            log("  inline script excerpt:", re.sub(r"\s+", " ", body)[:1500])

# 2. Hydroweb.next
section("Hydroweb.next collections for West Africa")
r = get("https://hydroweb.next.theia-land.fr/api/v1/rs-catalog/stac/collections")
cols = []
if r is not None and r.ok:
    j = r.json()
    cols = j.get("collections", [])
    log("  total collections", len(cols))
    for c in cols:
        log("   col:", c.get("id"), "|", (c.get("title") or "")[:80], "| ext", (c.get("extent", {}).get("temporal", {}).get("interval")))
names = ["volta", "akosombo", "kainji", "jebba", "shiroro", "manantali", "bui", "kossou", "soubre", "lagdo", "niger", "senegal", "chad", "mount coffee", "kompienga"]
ids = [c.get("id") for c in cols]
for c in cols:
    cid = c.get("id")
    for lk in (c.get("links") or []):
        pass
# search items in collections that look like reservoirs
for cid in ids:
    if cid and re.search(r"(?i)hydrocoast|lake|reserv|river|hydro|swot|altimet", cid):
        q = ("https://hydroweb.next.theia-land.fr/api/v1/rs-catalog/stac/search?collections=%s&limit=3&bbox=-5,4,15,15" % cid)
        rr = get(q)
        if rr is not None:
            log(f"  search {cid}: {rr.status_code} {rr.text[:700]!r}")
        break

# 3. Open-Meteo / GloFAS history depth
section("Open-Meteo flood API history depth (Akosombo)")
r = get("https://flood-api.open-meteo.com/v1/flood?latitude=6.30&longitude=0.06&daily=river_discharge&start_date=1984-01-01&end_date=1984-01-05")
if r is not None:
    log("  1984:", r.status_code, r.text[:400])
r = get("https://flood-api.open-meteo.com/v1/flood?latitude=6.30&longitude=0.06&daily=river_discharge&start_date=2026-09-01&end_date=2026-10-10&past_days=0")
if r is not None:
    log("  2026:", r.status_code, r.text[:300])

# 4. ABN bulletin
section("ABN hydrological bulletin")
r = get("https://www.abn.ne/files/15/Bulletins-hydrologiques-fr-FR/125/BulletinHydroJanvier2026.pdf")
if r is not None:
    log("  status", r.status_code, r.headers.get("content-type"), len(r.content))
    try:
        import subprocess
        open("abn.pdf", "wb").write(r.content)
        subprocess.run(["pip", "install", "-q", "pypdf"], check=False)
        from pypdf import PdfReader
        pdf = PdfReader("abn.pdf")
        log("  pages", len(pdf.pages))
        for i, pg in enumerate(pdf.pages[:6]):
            tx = pg.extract_text() or ""
            log(f"  --- page {i+1} ---\n", tx[:1800])
    except Exception as e:
        log("  pdf err", e)
r = get("https://www.abn.ne/fr/")
if r is not None:
    ls = sorted(set(re.findall(r"https://www\.abn\.ne/files/[^\"' ]+\.pdf", r.text)))
    log("  ABN pdf links:", len(ls))
    for l in ls[:40]:
        log("   ", l)

# 5. G-REALM alternatives
section("G-REALM")
for u in ["https://ipad.fas.usda.gov/cropexplorer/global_reservoir/", "https://ipad.fas.usda.gov/lakes/images/",
          "https://ipad.fas.usda.gov/cropexplorer/global_reservoir/gr_regional_chart.aspx?regionid=afr&ftypeid=22",
          "https://apps.fas.usda.gov/g-realm/", "https://www.fas.usda.gov/data/g-realm"]:
    rr = get(u)
    if rr is not None:
        log(" ", u, rr.status_code, rr.headers.get("content-type"), len(rr.content), re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", rr.text))[:200])

# 6. Other satellite
section("Other")
for u in ["https://dahiti.dgfi.tum.de/en/africa/", "https://www.fao.org/aquastat/en/databases/", 
          "https://zenodo.org/api/records?q=GRLM+reservoir+water+level&size=5",
          "https://zenodo.org/api/records?q=Volta+Lake+water+level&size=5"]:
    rr = get(u)
    if rr is not None:
        txt = rr.text
        titles = re.findall(r'"title": "([^"]+)"', txt)[:8]
        log(" ", u, rr.status_code, len(rr.content), titles or re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", txt))[:200])
log("Done")
OUT.close()
