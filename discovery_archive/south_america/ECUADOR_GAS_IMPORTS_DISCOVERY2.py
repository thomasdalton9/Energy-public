"""
Round 2 of ECUADOR_GAS_IMPORTS_DISCOVERY.py.

Round 1 found:
  - UN Comtrade public preview allows ONE period per call
    ("Maximum number of periods for preview is 1").
  - BCE quarterly oil-sector report (ASP, Q2 2026, p.27/32): Petroecuador's
    imported derivatives are naphtha, diesel and LPG only - no natural gas.
  - Petroecuador's Jan-Aug 2026 statistical report: no page mentions both
    natural gas and imports.
  - BCE: comercio-exterior BI page (bi_fw.html), BCEData portal, IEM list.
  - SENAE portal.aduana.gob.ec and ARCERNNR controlrecursosyenergia.gob.ec
    fail TLS verification (incomplete chain?); datosabiertos API 403;
    recursosyenergia.gob.ec / energia.gob.ec / arconel.gob.ec don't resolve.

This round:
  1. Comtrade monthly, one period per call: Ecuador imports HS 2711 (all
     petroleum gases - reported every month, mostly LPG), 271111 (LNG),
     271121 (gaseous natural gas), Jan 2021 onwards, partner detail. A month
     with a 2711 row but no 271111/271121 row = reported, zero gas imports.
     Annual mirror exports to Ecuador (Panama, Peru, USA, Colombia, Trinidad).
  2. ARCERNNR statistics pages via the Wayback Machine (normal TLS), and
     a CDX listing of its uploaded statistics files.
  3. BCE BCEData and comercio-exterior BI pages, IEM publication list.
  4. Ministerio de Energia y Minas hosts.
"""
import re
import time
from urllib.parse import urljoin

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 90)


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=kw.pop("timeout", T), **kw)
        out(f"GET {r.url[:200]} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)}B")
        return r
    except Exception as e:
        out(f"GET {url} -> ERR {type(e).__name__}: {str(e)[:220]}")
        return None


def links(r, pat, limit=60):
    seen = set()
    n = 0
    for h, t in re.findall(r'<a[^>]+href=["\']([^"\'#]+)["\'][^>]*>(.*?)</a>', r.text, re.S | re.I):
        t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
        u = urljoin(r.url, h.replace("&amp;", "&"))
        if u in seen or not re.search(pat, u + " " + t, re.I):
            continue
        seen.add(u)
        out(f"    LINK {t[:80]!r} -> {u[:220]}")
        n += 1
        if n >= limit:
            break
    for src in re.findall(r'<iframe[^>]+src=["\']([^"\']+)', r.text, re.I):
        out("    IFRAME", src[:220])


# ------------------------------------------------------------ 1. Comtrade
out("=================== 1. UN COMTRADE monthly, one period per call")
CT = "https://comtradeapi.un.org/public/v1/preview/C/{freq}/HS"


def comtrade(freq, **params):
    for attempt in range(4):
        try:
            r = requests.get(CT.format(freq=freq), params=params, headers=H, timeout=T)
        except Exception as e:
            out("   ERR", type(e).__name__, str(e)[:100])
            time.sleep(5)
            continue
        if r.status_code == 429:
            time.sleep(8 * (attempt + 1))
            continue
        time.sleep(1.2)
        if r.status_code != 200:
            out("   HTTP", r.status_code, r.text[:200])
            return None
        j = r.json()
        if j.get("error"):
            out("   error:", j.get("error"))
        return j.get("data") or []
    return None


months = [f"{y}{m:02d}" for y in range(2021, 2027) for m in range(1, 13) if f"{y}{m:02d}" <= "202608"]
summary = []
for p in months:
    rows = comtrade("M", reporterCode=218, period=p, cmdCode="2711,271111,271121", flowCode="M")
    if rows is None:
        summary.append((p, "request failed"))
        continue
    agg = [r for r in rows if r.get("partnerCode") == 0 and r.get("motCode") in (0, None)
           and r.get("customsCode") in ("C00", None) and r.get("partner2Code") in (0, None)]
    tot2711 = [r for r in agg if r.get("cmdCode") == "2711"]
    gas = [r for r in rows if r.get("cmdCode") in ("271111", "271121")]
    line = (f"{p}: rows={len(rows)} 2711_world_kg={tot2711[0].get('netWgt') if tot2711 else None} "
            f"2711_usd={tot2711[0].get('primaryValue') if tot2711 else None}")
    for g in gas:
        line += (f" | {g.get('cmdCode')} partner={g.get('partnerCode')} mot={g.get('motCode')} "
                 f"cust={g.get('customsCode')} kg={g.get('netWgt')} qty={g.get('qty')}{g.get('qtyUnitAbbr')} "
                 f"usd={g.get('primaryValue')}")
    out("  ", line)
    summary.append((p, line))
out("\n -- annual Ecuador imports + mirror exports to Ecuador")
for y in range(2021, 2026):
    rows = comtrade("A", reporterCode=218, period=y, cmdCode="2711,271111,271121", flowCode="M")
    for r in rows or []:
        if r.get("cmdCode") in ("271111", "271121") or r.get("partnerCode") == 0:
            out(f"   EC {y} {r.get('cmdCode')} partner={r.get('partnerCode')} mot={r.get('motCode')} "
                f"kg={r.get('netWgt')} usd={r.get('primaryValue')}")
    for name, code in (("Panama", 591), ("Peru", 604), ("USA", 842), ("Colombia", 170), ("Trinidad", 780)):
        rows = comtrade("A", reporterCode=code, period=y, cmdCode="271111,271121", flowCode="X", partnerCode=218)
        for r in rows or []:
            out(f"   mirror {name}->EC {y} {r.get('cmdCode')} kg={r.get('netWgt')} usd={r.get('primaryValue')} "
                f"mot={r.get('motCode')}")
        if rows == []:
            out(f"   mirror {name}->EC {y}: no rows")

# ------------------------------------------------------------ 2. ARCERNNR / SENAE via the Wayback Machine
out("\n=================== 2. ARCERNNR / SENAE pages via web.archive.org (their own TLS chains fail)")
for u in ["https://www.controlrecursosyenergia.gob.ec/estadisticas-del-sector-electrico-ecuatoriano/",
          "https://www.controlrecursosyenergia.gob.ec/boletines-estadisticos/",
          "https://www.controlrecursosyenergia.gob.ec/estadistica-del-sector-electrico/",
          "https://www.controlrecursosyenergia.gob.ec/"]:
    r = get(f"https://web.archive.org/web/2026/{u}")
    if r is not None and r.ok:
        links(r, r"estad|bolet|anual|mensual|multianual|combustib|xlsx|pdf", limit=60)
r = get("https://web.archive.org/cdx/search/cdx", params={
    "url": "controlrecursosyenergia.gob.ec/wp-content/uploads/*", "output": "json", "limit": 400,
    "filter": "original:.*(?i)(estad|bolet|combust).*", "collapse": "urlkey", "from": "2023"})
if r is not None and r.ok:
    try:
        for row in r.json()[1:200]:
            out("    CDX", row[1], row[2], row[4])
    except Exception as e:
        out("  cdx parse", e, r.text[:200])

# ------------------------------------------------------------ 3. BCE
out("\n=================== 3. BCE BCEData / comercio exterior / IEM")
for u in ["https://contenido.bce.fin.ec/bcedata/",
          "https://contenido.bce.fin.ec/documentos/PublicacionesNotas/bi_fw.html",
          "https://contenido.bce.fin.ec/iem-publicaciones/",
          "https://contenido.bce.fin.ec/estadisticas-del-sector-externo-d/"]:
    r = get(u)
    if r is not None and r.ok:
        links(r, r"comercio|import|subpartida|nandina|producto|IEM|xlsx|xls|csv|externo|balanza", limit=40)
        for s in re.findall(r'(?:src|href)=["\']([^"\']*(?:powerbi|tableau|app\.|api|json)[^"\']*)', r.text, re.I)[:15]:
            out("    EMBED", s[:220])

# ------------------------------------------------------------ 4. Ministry
out("\n=================== 4. Ministerio de Energia y Minas")
for u in ["https://www.energiayminas.gob.ec/", "https://www.recursosyenergia.gob.ec/", "https://recursosyenergia.gob.ec/",
          "https://www.ministeriodeenergia.gob.ec/", "https://www.energia.gob.ec/"]:
    r = get(u, timeout=(8, 40))
    if r is not None and r.ok:
        links(r, r"balance|estad|bolet|biblioteca|gas natural", limit=30)
out("DONE")
