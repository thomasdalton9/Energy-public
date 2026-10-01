"""
Round 2 of PERU_COES_ECUADOR_CENACE_DISCOVERY.py.

Round 1 found: COES' /Portal/portalinformacion/generacion POST returns,
for ANY date range, the fuel-type half-hourly chart of the FIRST day only
(48 points, MW) plus per-company totals by technology for the whole range
(MWh) - so daily history means one request per day (works back to 2021).
COES labels its 'BAGAZO' and 'SOLAR' fuel series the wrong way round
(the 'BAGAZO' one has the solar daytime shape and equals the SOLAR
technology total). CENACE has no archive: info-operativa/ is 403, the
WordPress media library holds only five 2024 R1/R3 xlsx files.

This round: COES' ExportarGeneracion (the page's Excel button); CENACE
history via the Wayback Machine (InformacionOperativa.htm snapshots -
each carries that day plus month-to-date totals), CENACE's despacho-real
/ biblioteca / indicadores pages, one R1 xlsx, ARCERNNR's site.
"""

print("STARTING", flush=True)

import collections
import io
import json
import os
import re
from urllib.parse import urljoin

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
OUT = "ecpe_raw"
os.makedirs(OUT, exist_ok=True)
S = requests.Session()
S.headers["User-Agent"] = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                           "Chrome/124.0 Safari/537.36")


def get(url, name=None, **kw):
    try:
        r = S.get(url, timeout=kw.pop("timeout", 90), verify=kw.pop("verify", True), **kw)
    except requests.RequestException as e:
        print(f"== GET {url}: FAILED {type(e).__name__}: {str(e)[:200]}", flush=True)
        return None
    print(f"== GET {url}: {r.status_code} {r.headers.get('Content-Type', '')} {len(r.content):,} bytes", flush=True)
    if name:
        open(os.path.join(OUT, name), "wb").write(r.content)
    return r


def show_xlsx(content):
    try:
        xl = pd.ExcelFile(io.BytesIO(content))
    except Exception as e:  # noqa: BLE001
        print("  not excel:", e, content[:200], flush=True)
        return
    for s in xl.sheet_names[:6]:
        d = pd.read_excel(xl, s, header=None)
        print(f"  sheet {s!r} shape={d.shape}", flush=True)
        print(d.head(25).iloc[:, :14].to_string(max_colwidth=28), flush=True)


# ---------------------------------------------------------------- PERU
print("\n############ PERU COES export", flush=True)
for label, a, b in [("1day", "01/09/2026", "01/09/2026"), ("1month", "01/08/2026", "31/08/2026")]:
    r = get("https://www.coes.org.pe/Portal/portalinformacion/ExportarGeneracion",
            f"pe_export_{label}.bin", params={"fechaInicial": a, "fechaFinal": b, "indicador": 0}, timeout=240)
    if r is not None and r.ok:
        print("  headers:", dict(r.headers), flush=True)
        show_xlsx(r.content)

# ---------------------------------------------------------------- ECUADOR
print("\n############ ECUADOR wayback", flush=True)
r = get("https://web.archive.org/cdx/search/cdx", "ec_wayback_cdx.json",
        params={"url": "cenace.gob.ec/info-operativa/InformacionOperativa.htm", "output": "json",
                "from": "2019", "filter": "statuscode:200"}, timeout=180)
snaps = []
if r is not None and r.ok:
    try:
        rows = r.json()
        snaps = [dict(zip(rows[0], x)) for x in rows[1:]]
    except ValueError:
        print(r.text[:500])
print(f"  {len(snaps)} snapshots", flush=True)
per_month = collections.Counter(s["timestamp"][:6] for s in snaps)
print("  per month:", dict(sorted(per_month.items())), flush=True)
days = sorted({s["timestamp"][:8] for s in snaps})
print(f"  distinct days: {len(days)} first={days[:3]} last={days[-3:]}", flush=True)
# fetch the oldest snapshot and one from 2023 / 2025, print section headers + dates
picks = []
for yr in ("2021", "2022", "2023", "2024", "2025"):
    s = next((x for x in snaps if x["timestamp"].startswith(yr)), None)
    if s:
        picks.append(s)
for s in picks:
    u = f"https://web.archive.org/web/{s['timestamp']}id_/{s['original']}"
    rr = get(u, f"ec_wayback_{s['timestamp']}.html", timeout=120)
    if rr is not None and rr.ok:
        t = re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", rr.text, flags=re.S))
        for m in re.finditer(r"INFORMACI[ÓO]N OPERATIVA \w+", t):
            print("   ", t[m.start():m.start() + 400], flush=True)
        print("    plotly:", rr.text.count("Plotly.newPlot"), flush=True)

print("\n############ ECUADOR CENACE pages", flush=True)
for u in ["https://www.cenace.gob.ec/elaboracion-del-despacho-real/", "https://www.cenace.gob.ec/biblioteca/",
          "https://www.cenace.gob.ec/indicadores/", "https://www.cenace.gob.ec/transparencia/",
          "https://www.controlrecursosyenergia.gob.ec/"]:
    rr = get(u, verify=False)
    if rr is None or not rr.ok:
        continue
    for href, label in re.findall(r'<a[^>]+href=["\']([^"\'#]+)["\'][^>]*>(.*?)</a>', rr.text, re.S | re.I):
        label = re.sub(r"<[^>]+>|\s+", " ", label).strip()
        if re.search(r"estad|operat|operac|despacho|produc|generac|anual|mensual|diari|balance|xls|csv|informe|"
                     r"bi\.|powerbi|reporte|sirio|bosni", href + " " + label, re.I):
            print(f"    LINK {label[:70]!r} -> {urljoin(u, href)}", flush=True)
    for src in re.findall(r'<iframe[^>]+src=["\']([^"\']+)', rr.text, re.I):
        print("    IFRAME", src, flush=True)

rr = get("https://www.cenace.gob.ec/wp-content/uploads/2024/03/R1_2024-03-09.xlsx", "ec_R1_2024-03-09.xlsx",
         verify=False)
if rr is not None and rr.ok:
    show_xlsx(rr.content)

print("\nDONE", flush=True)
