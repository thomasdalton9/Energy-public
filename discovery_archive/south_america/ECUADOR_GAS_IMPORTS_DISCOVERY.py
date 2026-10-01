"""
Ecuador natural gas IMPORTS: does Ecuador import LNG / natural gas, and
where is it published monthly? (ECUADOR_GAS.py covers domestic Amistad
gas only.)

Probes:
  1. UN Comtrade public preview API - Ecuador (reporter 218) monthly
     imports of HS 271111 (LNG) and 271121 (natural gas, gaseous), 2021+,
     plus mirror exports to Ecuador from Panama, Peru, Colombia, USA,
     Trinidad. Comtrade carries Ecuador's customs (SENAE/BCE) returns.
  2. Banco Central del Ecuador: comercio exterior pages, Informacion
     Estadistica Mensual, the quarterly oil-sector report (ASP) gas pages.
  3. SENAE (aduana.gob.ec) statistics pages.
  4. datosabiertos.gob.ec CKAN: import / gas datasets.
  5. ARCERNNR / ARCONEL electricity statistics (fuel use by thermal plants),
     Ministerio de Energia y Minas energy balance.
  6. Petroecuador YTD statistical report: import tables mentioning gas.
Not reachable from the editing sandbox; runs in GitHub Actions.
"""
import io
import json
import re
import time
from urllib.parse import urljoin

import pdfplumber
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 90)
LINK_PAT = r"estad|comercio|import|subpartida|mensual|boletin|bolet|balance|xlsx|xls|csv|pdf|gas|combustib|anual"


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=T, **kw)
        out(f"GET {r.url[:200]} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)}B")
        return r
    except Exception as e:
        out(f"GET {url} -> ERR {type(e).__name__}: {str(e)[:150]}")
        return None


def links(r, pat=LINK_PAT, limit=60):
    seen = set()
    n = 0
    for h, t in re.findall(r'<a[^>]+href=["\']([^"\'#]+)["\'][^>]*>(.*?)</a>', r.text, re.S | re.I):
        t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
        u = urljoin(r.url, h.replace("&amp;", "&"))
        if u in seen or not re.search(pat, u + " " + t, re.I):
            continue
        seen.add(u)
        out(f"    LINK {t[:80]!r} -> {u[:200]}")
        n += 1
        if n >= limit:
            break
    for src in re.findall(r'<iframe[^>]+src=["\']([^"\']+)', r.text, re.I):
        out("    IFRAME", src[:200])


def pdf_gas(content, pat=r"gas natural|GNL|licuado|import", max_lines=50):
    if content[:4] != b"%PDF":
        out("  not a PDF:", content[:80])
        return
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        out(f"  {len(pdf.pages)} pages")
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if not (re.search(r"gas natural|GNL|licuado", text, re.I) and re.search(r"import", text, re.I)):
                continue
            out(f"  --- page {i + 1}")
            for ln in text.splitlines()[:max_lines]:
                if re.search(pat, ln, re.I) or len(ln) < 60:
                    out("     ", ln[:200])


# ------------------------------------------------------------ 1. Comtrade
out("=================== 1. UN COMTRADE public preview")
CT = "https://comtradeapi.un.org/public/v1/preview/C/{freq}/HS"


def comtrade(freq, **params):
    for attempt in range(3):
        r = get(CT.format(freq=freq), params=params)
        if r is not None and r.status_code == 429:
            time.sleep(10)
            continue
        time.sleep(2)
        if r is None or r.status_code != 200:
            if r is not None:
                out("   body:", r.text[:300])
            return []
        try:
            j = r.json()
        except Exception:
            out("   not json:", r.text[:200])
            return []
        if j.get("error"):
            out("   error:", j.get("error"))
        return j.get("data") or []
    return []


def show(rows):
    for d in sorted(rows, key=lambda d: (str(d.get("period")), str(d.get("cmdCode")), d.get("partnerCode") or 0)):
        out(f"     {d.get('period')} rep={d.get('reporterCode')} flow={d.get('flowCode')} cmd={d.get('cmdCode')} "
            f"partner={d.get('partnerCode')} p2={d.get('partner2Code')} mot={d.get('motCode')} "
            f"cust={d.get('customsCode')} netWgt={d.get('netWgt')} qty={d.get('qty')} {d.get('qtyUnitAbbr')} "
            f"altQty={d.get('altQty')} {d.get('altQtyUnitAbbr')} cif={d.get('cifvalue')} "
            f"value={d.get('primaryValue')}")


for y in range(2021, 2027):
    periods = ",".join(f"{y}{m:02d}" for m in range(1, 13))
    out(f"\n -- Ecuador monthly imports {y}, HS 271111/271121 (all partners)")
    rows = comtrade("M", reporterCode=218, period=periods, cmdCode="271111,271121", flowCode="M")
    out(f"   {len(rows)} rows")
    show(rows)
    out(f" -- sanity: Ecuador monthly imports {y}, HS 2711 / 271112 / 271113 / 271119, partner World")
    rows = comtrade("M", reporterCode=218, period=periods, cmdCode="2711,271112,271113,271119", flowCode="M",
                    partnerCode=0)
    periods_seen = sorted({r.get("period") for r in rows})
    out(f"   {len(rows)} rows; periods with data: {periods_seen}")
    show([r for r in rows if r.get("cmdCode") in ("2711",) and r.get("motCode") in (0, None)
          and r.get("customsCode") in ("C00", None) and r.get("partner2Code") in (0, None)][:14])

out("\n -- Ecuador annual imports 2021-2025, HS 271111/271121")
show(comtrade("A", reporterCode=218, period="2021,2022,2023,2024,2025", cmdCode="271111,271121", flowCode="M"))

for name, code in (("Panama", 591), ("Peru", 604), ("Colombia", 170), ("USA", 842), ("Trinidad", 780),
                   ("Chile", 152)):
    out(f"\n -- mirror: {name} exports to Ecuador HS 271111/271121, annual 2021-2025")
    show(comtrade("A", reporterCode=code, period="2021,2022,2023,2024,2025", cmdCode="271111,271121",
                  flowCode="X", partnerCode=218))
for name, code in (("Panama", 591), ("Peru", 604)):
    for y in (2022, 2024, 2025, 2026):
        periods = ",".join(f"{y}{m:02d}" for m in range(1, 13))
        out(f" -- mirror monthly: {name} exports to Ecuador {y}")
        show(comtrade("M", reporterCode=code, period=periods, cmdCode="271111,271121", flowCode="X",
                      partnerCode=218))

# ------------------------------------------------------------ 2. BCE
out("\n=================== 2. BANCO CENTRAL DEL ECUADOR")
for u in ["https://www.bce.fin.ec/", "https://www.bce.fin.ec/comercio-exterior/",
          "https://www.bce.fin.ec/informacion-estadistica-mensual/",
          "https://www.bce.fin.ec/estadisticas-de-comercio-exterior/",
          "https://www.bce.fin.ec/estadisticas-economicas/",
          "https://contenido.bce.fin.ec/home1/estadisticas/bolmensual/IEMensual.jsp",
          "https://contenido.bce.fin.ec/documentos/Estadisticas/SectorExterno/BalanzaPagos/balanzaComercial/ebc202601.pdf",
          "https://www.bce.fin.ec/sector-petrolero/",
          "https://www.bce.fin.ec/informacion-del-sector-petrolero/"]:
    r = get(u)
    if r is not None and r.ok and "html" in r.headers.get("content-type", ""):
        links(r, pat=r"comercio|import|subpartida|petrol|hidrocarb|mensual|IEM|balanza|xlsx|pdf|estad", limit=40)
out("\n -- BCE oil-sector report (ASP), latest editions: gas + import pages")
for code in ("ASP202602", "ASP202601", "ASP202504", "ASP202503"):
    r = get(f"https://contenido.bce.fin.ec/documentos/Estadisticas/Hidrocarburos/{code}.pdf")
    if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
        pdf_gas(r.content)
        break

# ------------------------------------------------------------ 3. SENAE
out("\n=================== 3. SENAE")
for u in ["https://www.aduana.gob.ec/", "https://www.aduana.gob.ec/estadisticas/",
          "https://www.aduana.gob.ec/sistema-de-estadisticas/", "https://portal.aduana.gob.ec/",
          "https://www.produccion.gob.ec/", "https://www.produccion.gob.ec/estadisticas-de-comercio-exterior/"]:
    r = get(u)
    if r is not None and r.ok and "html" in r.headers.get("content-type", ""):
        links(r, pat=r"estad|comercio|import|subpartida|consulta|xlsx|csv", limit=30)

# ------------------------------------------------------------ 4. datosabiertos
out("\n=================== 4. datosabiertos.gob.ec")
for q in ("importaciones", "importaciones subpartida", "gas natural", "comercio exterior", "consumo combustible"):
    r = get("https://www.datosabiertos.gob.ec/api/3/action/package_search", params={"q": q, "rows": 12})
    if r is None or r.status_code != 200:
        continue
    try:
        for p in r.json()["result"]["results"]:
            out(f"  [{q}] {p['name']}: {p.get('title', '')[:90]} ({p.get('organization', {}).get('title', '')})")
            for res in p.get("resources", [])[:4]:
                out(f"       {res.get('format')} {res.get('name', '')[:60]!r} {res.get('url')}")
    except Exception as e:
        out("  parse error", e, r.text[:200])

# ------------------------------------------------------------ 5. electricity regulator / ministry
out("\n=================== 5. ARCERNNR / ARCONEL / Ministerio de Energia")
for u in ["https://www.controlrecursosyenergia.gob.ec/estadisticas-del-sector-electrico-ecuatoriano/",
          "https://www.controlrecursosyenergia.gob.ec/estadistica-del-sector-electrico/",
          "https://www.controlrecursosyenergia.gob.ec/boletines-estadisticos/",
          "https://www.controlrecursosyenergia.gob.ec/",
          "https://www.arconel.gob.ec/", "https://www.arconel.gob.ec/estadisticas/",
          "https://www.recursosyenergia.gob.ec/", "https://www.recursosyenergia.gob.ec/balance-energetico-nacional/",
          "https://www.energia.gob.ec/", "https://www.recursosyenergia.gob.ec/biblioteca/"]:
    r = get(u)
    if r is not None and r.ok and "html" in r.headers.get("content-type", ""):
        links(r, pat=r"estad|boletin|bolet|balance|mensual|anual|multianual|combustib|xlsx|pdf", limit=40)

# ------------------------------------------------------------ 6. Petroecuador imports
out("\n=================== 6. PETROECUADOR YTD report: import pages mentioning gas")
r = get("https://www.eppetroecuador.ec/?p=3721")
if r is not None and r.ok:
    html = r.text
    i = html.find("Mensuales")
    block = html[i:i + 4000] if i > 0 else ""
    for h, t in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', block, re.S | re.I)[:2]:
        t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
        out("  report:", t, h)
        rr = get(urljoin(r.url, h.replace("&amp;", "&")))
        if rr is not None and rr.ok:
            pdf_gas(rr.content, pat=r"gas|GNL|import|total", max_lines=60)
        break
out("DONE")
