"""
Follow-up to SA_GAS_DEMAND_DISCOVERY.py.
Brazil: the MME monthly gas bulletin ("Boletim Mensal de Acompanhamento
  da Industria de Gas Natural") has per-year pages and an "anexos" page -
  list every file link on those (looking for xlsx annexes with demand by
  segment), then download the newest xlsx and dump its sheet names and
  the rows that mention consumo/demanda/segmento.
Colombia: the Gestor del Mercado de Gas "BI Gas > Demanda" page is a
  Power BI embed - capture the embed-token call and the QueryExecution
  POST bodies + responses so a scraper can replay them; also print the
  datos.gov.co catalog hits for gas demand.
"""
import io
import json
import re

import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 60)
BASE = ("https://www.gov.br/mme/pt-br/assuntos/secretarias/petroleo-gas-natural-e-biocombustiveis/publicacoes-1/"
        "boletim-mensal-de-acompanhamento-da-industria-de-gas-natural")


def out(*a):
    print(*a, flush=True)


def files_on(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    except Exception as e:
        out(f"  ERR {url}: {e}")
        return []
    out(f"-- {url} -> {r.status_code}")
    fl = sorted(set(re.findall(r'href="([^"]+?(?:\.xlsx?|\.pdf|\.csv|\.zip|/@@download/file)[^"]*)"', r.text, re.I)))
    for f in fl:
        out("   ", f)
    subs = sorted(set(re.findall(r'href="(' + re.escape(BASE) + r'/[^"#?]+)"', r.text)))
    return fl, subs


out("################ BRAZIL: MME gas bulletin ################")
all_files = []
for page in [f"{BASE}/anexos", f"{BASE}/2024", f"{BASE}/2025", f"{BASE}/2026",
             "https://www.gov.br/mme/pt-br/assuntos/secretarias/petroleo-gas-natural-e-biocombustiveis/dgn/2025",
             "https://www.gov.br/mme/pt-br/assuntos/secretarias/petroleo-gas-natural-e-biocombustiveis/dgn/2026"]:
    res = files_on(page)
    if res:
        all_files += res[0]
xls = [f for f in all_files if re.search(r"\.xlsx?($|[/?])", f, re.I) or ("anexo" in f.lower() and "download" in f.lower())]
out(f"\n{len(xls)} spreadsheet-looking links")
for f in xls[-3:]:
    url = f if f.startswith("http") else "https://www.gov.br" + f
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        out(f"\n== {url} -> {r.status_code} {r.headers.get('content-type','')[:50]} {len(r.content)}B")
        import pandas as pd
        xl = pd.ExcelFile(io.BytesIO(r.content))
        out("   sheets:", xl.sheet_names)
        for s in xl.sheet_names:
            df = xl.parse(s, header=None)
            hits = df[df.astype(str).apply(lambda row: row.str.contains("consumo|demanda|segmento|industrial|termel|automotiv|residencial", case=False, regex=True)).any(axis=1)]
            if not hits.empty:
                out(f"   -- sheet {s!r} shape {df.shape}; matching rows:")
                out(hits.head(25).to_string(max_colwidth=28)[:4000])
    except Exception as e:
        out(f"   ERR {type(e).__name__}: {e}")

out("\n################ COLOMBIA: datos.gov.co ################")
for q in ["gas natural demanda", "consumo gas natural", "gas natural sector"]:
    try:
        r = requests.get("https://api.us.socrata.com/api/catalog/v1", headers=HEADERS, timeout=TIMEOUT,
                         params={"domains": "www.datos.gov.co", "q": q, "limit": 12})
        for res in r.json().get("results", []):
            x = res["resource"]
            out(f"   [{x['id']}] {x['name']} | updated {str(x.get('data_updated_at',''))[:10]} | cols {x.get('columns_name', [])[:10]}")
    except Exception as e:
        out("  ERR", e)

out("\n################ COLOMBIA: BI Gas demanda (Power BI) ################")
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    b = p.chromium.launch()
    ctx = b.new_context(user_agent=HEADERS["User-Agent"], viewport={"width": 1600, "height": 1000})
    page = ctx.new_page()
    reqs = []

    def on_req(req):
        if "execute-api" in req.url or "querydata" in req.url.lower() or "QueryExecutionService" in req.url:
            reqs.append(req)
    page.on("request", on_req)
    for url in ["https://www.bmcbec.com.co/bi-gas/demanda"]:
        out(f"\n== {url}")
        try:
            page.goto(url, timeout=90000, wait_until="load")
            page.wait_for_timeout(20000)
            out("   title:", page.title())
            for fr in page.frames:
                out("   frame:", fr.url[:200])
        except Exception as e:
            out("   goto error:", type(e).__name__, str(e)[:200])
    out(f"\n{len(reqs)} Power BI requests captured")
    for i, rq in enumerate(reqs[:12]):
        out(f"\n--- [{i}] {rq.method} {rq.url[:220]}")
        hdr = {k: v for k, v in rq.headers.items() if k.lower() in ("authorization", "x-powerbi-resourcekey", "activityid", "requestid", "content-type")}
        out("   headers:", json.dumps({k: (v[:60] + "...") if len(v) > 60 else v for k, v in hdr.items()}))
        if rq.post_data:
            out("   body:", rq.post_data[:2500])
        try:
            resp = rq.response()
            if resp:
                body = resp.text()
                out(f"   resp {resp.status} {len(body)}B:", body[:2500])
        except Exception as e:
            out("   resp err", e)
    ctx.close()
    b.close()
