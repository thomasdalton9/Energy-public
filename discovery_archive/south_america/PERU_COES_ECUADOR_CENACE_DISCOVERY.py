"""
One-off probe: find keyless, scriptable sources of DAILY generation by
type for Peru (COES) and Ecuador (CENACE history), so
PERU_COES_GENERATION.py / ECUADOR_POWER_DAILY.py can be written against
real responses.

Peru: COES' Portal de Informacion 'Generacion' page posts a date range to
/Portal/portalinformacion/generacion and gets back JSON with half-hourly
MW by fuel ('GraficoTipoCombustible'). Probes a 1-day, 1-month and 1-year
range and an early-2021 day, and saves the page + its scripts.

Ecuador: CENACE's InformacionOperativa.htm only shows the last day. Looks
for an archive: the info-operativa/ folder listing, dated variants of the
page, the WordPress media/pages API (PDF/xlsx daily reports), and
ARCERNNR's statistics pages.

Raw responses are saved under ecpe_raw/ (pushed to a throwaway branch by
the workflow).
"""

print("STARTING", flush=True)

import json
import os
import re
from urllib.parse import urljoin

import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
OUT = "ecpe_raw"
os.makedirs(OUT, exist_ok=True)
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36"}
S = requests.Session()
S.headers.update(H)


def save(name, content):
    with open(os.path.join(OUT, name), "wb") as f:
        f.write(content if isinstance(content, bytes) else content.encode("utf-8"))


def get(url, name=None, method="GET", **kw):
    try:
        r = S.request(method, url, timeout=kw.pop("timeout", 90), verify=kw.pop("verify", True), **kw)
    except requests.RequestException as e:
        print(f"== {method} {url}: FAILED {type(e).__name__}: {str(e)[:200]}", flush=True)
        return None
    print(f"== {method} {url}: {r.status_code} {r.headers.get('Content-Type', '')} {len(r.content):,} bytes",
          flush=True)
    if name:
        save(name, r.content)
    return r


# ---------------------------------------------------------------- PERU
print("\n############ PERU COES", flush=True)
page = get("https://www.coes.org.pe/Portal/portalinformacion/generacion", "pe_generacion_page.html")
if page is not None and page.ok:
    for s in sorted(set(re.findall(r'<script[^>]+src=["\']([^"\']+)["\']', page.text))):
        if "jquery" in s.lower() and "portal" not in s.lower():
            continue
        u = urljoin(page.url, s)
        print("  SCRIPT", u, flush=True)
        if "portal" in u.lower() or "informacion" in u.lower() or "generacion" in u.lower():
            get(u, "pe_js_" + re.sub(r"[^A-Za-z0-9.]+", "_", s)[-80:])
    for m in re.findall(r'(?:url|action|href)\s*[:=]\s*["\']([^"\']*(?:portalinformacion|Export|export|Descarg|descarg)[^"\']*)["\']',
                        page.text):
        print("  ENDPOINT-REF", m, flush=True)
    for m in re.findall(r'<(?:input|select)[^>]+(?:id|name)=["\']([^"\']+)["\']', page.text):
        print("  FIELD", m, flush=True)

PE_URL = "https://www.coes.org.pe/Portal/portalinformacion/generacion"
for label, a, b in [("1day", "01/09/2026", "01/09/2026"), ("1month", "01/08/2026", "31/08/2026"),
                    ("1year", "01/01/2025", "31/12/2025"), ("2021day", "01/01/2021", "01/01/2021"),
                    ("2021q1", "01/01/2021", "31/03/2021")]:
    r = get(PE_URL, f"pe_gen_{label}.json", method="POST",
            data={"fechaInicial": a, "fechaFinal": b, "indicador": 0}, timeout=240)
    if r is None or not r.ok:
        continue
    try:
        j = r.json()
    except ValueError:
        print("  not JSON:", r.text[:300], flush=True)
        continue
    print(f"  [{label}] top keys: {list(j)[:30]}", flush=True)
    for k, v in j.items():
        if isinstance(v, dict):
            print(f"    {k}: dict keys {list(v)[:15]}", flush=True)
            ser = v.get("Series")
            if isinstance(ser, list):
                for s in ser:
                    d = s.get("Data") or []
                    print(f"      series {s.get('Name')!r} n={len(d)} first={d[:1]} last={d[-1:]} "
                          f"other={ {kk: vv for kk, vv in s.items() if kk != 'Data'} }", flush=True)
        elif isinstance(v, list):
            print(f"    {k}: list n={len(v)} first={str(v[:2])[:400]}", flush=True)
        else:
            print(f"    {k}: {str(v)[:200]}", flush=True)

# other COES portal-informacion pages that might give daily energy by resource
for path in ["Portal/portalinformacion/demanda", "Portal/portalinformacion/Index",
             "Portal/PostOperacion/Reportes/Ieod", "Portal/portalinformacion/produccion"]:
    get("https://www.coes.org.pe/" + path, "pe_" + path.replace("/", "_") + ".html")

# ---------------------------------------------------------------- ECUADOR
print("\n############ ECUADOR CENACE", flush=True)
for u in ["https://www.cenace.gob.ec/info-operativa/", "https://www.cenace.gob.ec/info-operativa/InformacionOperativa/",
          "https://www.cenace.gob.ec/info-operativa/InformacionOperativa_20260901.htm",
          "https://www.cenace.gob.ec/info-operativa/2026-09-01.htm"]:
    r = get(u, verify=False)
    if r is not None and r.ok:
        save("ec_" + re.sub(r"[^A-Za-z0-9.]+", "_", u)[-60:] + ".html", r.content)
        print("   ", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))[:600], flush=True)

for api in ["wp-json/wp/v2/media?per_page=100&search=diari",
            "wp-json/wp/v2/media?per_page=100&search=operaci",
            "wp-json/wp/v2/media?per_page=100&search=producci",
            "wp-json/wp/v2/media?per_page=100&mime_type=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "wp-json/wp/v2/media?per_page=100&mime_type=application/vnd.ms-excel",
            "wp-json/wp/v2/pages?per_page=100&_fields=link,title",
            "wp-json/wp/v2/posts?per_page=50&search=operaci&_fields=link,title,date"]:
    r = get("https://www.cenace.gob.ec/" + api, verify=False)
    if r is None or not r.ok:
        continue
    save("ec_" + re.sub(r"[^A-Za-z0-9]+", "_", api)[:80] + ".json", r.content)
    try:
        items = r.json()
    except ValueError:
        print("  not JSON", r.text[:200], flush=True)
        continue
    print(f"  total={r.headers.get('X-WP-Total')} pages={r.headers.get('X-WP-TotalPages')}", flush=True)
    for it in items[:100]:
        title = it.get("title", {}).get("rendered") if isinstance(it.get("title"), dict) else it.get("title")
        print(f"    {it.get('date', '')[:10]} {title!r} {it.get('source_url') or it.get('link')}", flush=True)

for u in ["https://www.controlrecursosyenergia.gob.ec/estadisticas-del-sector-electrico-ecuatoriano/",
          "https://www.controlrecursosyenergia.gob.ec/balance-nacional-de-energia-electrica/",
          "https://www.cenace.gob.ec/informes-anuales/", "https://www.cenace.gob.ec/informacion-operativa/"]:
    r = get(u, "ec_" + re.sub(r"[^A-Za-z0-9]+", "_", u)[8:70] + ".html", verify=False)
    if r is not None and r.ok:
        for l in sorted(set(re.findall(r'href=["\']([^"\']+\.(?:xlsx?|csv|pdf|zip)[^"\']*)["\']', r.text, re.I)))[:60]:
            print("    LINK", l, flush=True)
        for l in sorted(set(re.findall(r'href=["\']([^"\']*(?:balance|estad|diari|operativ|bi\b|powerbi|tableau)[^"\']*)["\']',
                                       r.text, re.I)))[:60]:
            print("    PAGE", l, flush=True)

print("\nDONE", flush=True)
