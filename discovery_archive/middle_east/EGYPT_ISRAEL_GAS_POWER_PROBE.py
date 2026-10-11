"""
Discovery probe: raw gas and power data for EGYPT and ISRAEL (Oct 2026).

Round 1 (requests only): fetch the candidate official sites, print status,
content type, size, <title>, and the links on each page that look like data
(statistics / reports / xlsx / csv / pdf / api), plus any API-looking URLs in
inline scripts. Also queries Israel's CKAN open-data portal (data.gov.il) for
gas and electricity datasets.

RUN 1 (2026-10-11, requests): ISRAEL - noga-iso.co.il reachable (200, Hebrew
only, /en/ paths 404); links: /statistics/, /pdt/highvoltageproduction/,
/pdt/power-consumption/, /systemoperationunit/demand-curve/,
/statistics/annual-consumption-data/, /statistics/annual-hourly-consumption-data/;
sitemap.xml 64 KB. apim-portal.noga-iso.co.il = Azure APIM sign-in (SMP API,
developer guide PDF /media/smsdo1bd/api-smp-developer-guide-v2.pdf). gov.il
(Ministry of Energy, Electricity Authority) and energean.com: 403 Cloudflare
"Just a moment". iec.co.il 200 (JS shell). cbs.gov.il Energy page 200.
data.gov.il CKAN works but has no gas/electricity datasets. newmedenergy.com
200 (quarterly/financial reports, Leviathan bond reports).
EGYPT - eehc.gov.eg 200 (tiny shell page); moee.gov.eg http 200 (2 KB), https
refused; eetc.net.eg connect TIMEOUT (both http/https); egyptera.org 200 with
ElectricityObservatory.aspx, GenEvaluation.aspx, annualReport.aspx,
PeriodicalBooks2022.aspx; egas.com.eg 200 (Liferay, no statistics links);
petroleum.gov.eg 200 (SharePoint; gas.aspx); capmas.gov.eg and idsc.gov.eg
200 but 1.4 KB JS shells; egypt-odp.portaljs.com has no energy datasets.

Workflow (manual): discovery_archive/workflows/egypt_israel_gas_power_probe.yml.
"""

import re
import sys
import json
from urllib.parse import urljoin, quote

import requests

H = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9,he;q=0.8,ar;q=0.7",
}
T = (12, 40)
KEY = re.compile(r"(stat|report|data|xls|csv|pdf|api|json|download|annual|month|gas|product|generat|"
                 r"consum|demand|load|export|lng|bulletin|%D7|%D8|natural|electric|energy|"
                 r"גז|חשמל|إحصا|تقارير)", re.I)


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=T, allow_redirects=True, **kw)
        return r
    except requests.RequestException as e:
        out(f"  ERROR {type(e).__name__}: {str(e)[:160]}")
        return None


def page(url, max_links=60):
    out(f"\n=== {url}")
    r = get(url)
    if r is None:
        return None
    ct = r.headers.get("content-type", "")
    title = ""
    m = re.search(r"<title[^>]*>(.*?)</title>", r.text[:20000], re.S | re.I) if "html" in ct else None
    if m:
        title = " ".join(m.group(1).split())[:100]
    out(f"  {r.status_code} {ct[:40]} {len(r.content)}B final={r.url[:100]} title={title!r}")
    if r.status_code != 200 or "html" not in ct:
        if r.status_code != 200:
            out("  body:", " ".join(r.text[:200].split()))
        return r
    links = []
    for href, txt in re.findall(r'<a[^>]+href=["\']([^"\'#]+)["\'][^>]*>(.*?)</a>', r.text, re.S | re.I):
        t = " ".join(re.sub(r"<[^>]+>", " ", txt).split())[:60]
        full = urljoin(r.url, href)
        if KEY.search(full) or KEY.search(t):
            links.append((full, t))
    seen = set()
    n = 0
    for full, t in links:
        if full in seen:
            continue
        seen.add(full)
        out(f"   - {full[:150]} | {t}")
        n += 1
        if n >= max_links:
            out("   ...")
            break
    apis = sorted(set(re.findall(r'["\'](/?(?:[\w\-/]*/)?(?:api|umbraco/api|Umbraco/Surface)[^"\'\s<>]*)["\']',
                                 r.text, re.I)))[:20]
    if apis:
        out("  api-like strings:", apis)
    return r


def ckan(base, q):
    url = f"{base}/api/3/action/package_search?q={quote(q)}&rows=15"
    out(f"\n=== CKAN {url}")
    r = get(url)
    if r is None:
        return
    out(f"  {r.status_code} {r.headers.get('content-type','')[:30]} {len(r.content)}B")
    try:
        js = r.json()
    except Exception:
        out("  not json:", " ".join(r.text[:200].split()))
        return
    res = js.get("result", {})
    out(f"  count={res.get('count')}")
    for p in res.get("results", []):
        out(f"  * {p.get('name')} | {p.get('title')} | org={((p.get('organization') or {}).get('title'))} "
            f"| modified={p.get('metadata_modified','')[:10]}")
        for rs in p.get("resources", [])[:6]:
            out(f"      {rs.get('format')} {rs.get('last_modified') or rs.get('created','')}"
                f" {str(rs.get('url'))[:140]}")


ISRAEL = [
    "https://www.noga-iso.co.il/",
    "https://www.noga-iso.co.il/en/",
    "https://www.noga-iso.co.il/en/pdt/highvoltageproduction/",
    "https://www.noga-iso.co.il/en/statistics/annual-hourly-consumption-data/",
    "https://www.noga-iso.co.il/en/statistics/annual-consumption-data/",
    "https://www.noga-iso.co.il/sitemap.xml",
    "https://apim-portal.noga-iso.co.il/",
    "https://apim-portal.noga-iso.co.il/apis",
    "https://www.iec.co.il/en/about/statistics",
    "https://www.iec.co.il/statistics",
    "https://www.gov.il/en/departments/ministry_of_energy/govil-landing-page",
    "https://www.gov.il/he/departments/topics/natural_gas_sector/govil-landing-page",
    "https://www.gov.il/en/pages/natural_gas_economy",
    "https://www.gov.il/he/departments/electricity_authority/govil-landing-page",
    "https://www.gov.il/en/departments/electricity_authority/govil-landing-page",
    "https://www.cbs.gov.il/en/subjects/Pages/Energy.aspx",
    "https://www.cbs.gov.il/he/subjects/Pages/%D7%90%D7%A0%D7%A8%D7%92%D7%99%D7%94.aspx",
    "https://www.energean.com/operations/israel/",
    "https://www.newmedenergy.com/",
]
EGYPT = [
    "http://www.eehc.gov.eg/eehcportal/",
    "https://www.eehc.gov.eg/",
    "http://www.eehc.gov.eg/eehcportal/YearlyReport.aspx",
    "http://www.moee.gov.eg/",
    "https://www.moee.gov.eg/english_new/report.aspx",
    "http://www.eetc.net.eg/",
    "https://www.eetc.net.eg/statistics.html",
    "http://egyptera.org/en/",
    "https://egyptera.org/en/",
    "https://egyptera.org/ar/",
    "https://www.egas.com.eg/",
    "https://egas.com.eg/",
    "https://www.petroleum.gov.eg/en/",
    "https://www.petroleum.gov.eg/ar/",
    "https://www.capmas.gov.eg/",
    "https://www.capmas.gov.eg/Pages/Publications.aspx?page_id=5104",
    "https://egypt-odp.portaljs.com/",
    "https://egypt.opendataforafrica.org/",
    "https://www.idsc.gov.eg/",
]

def round1():
    out("##### ISRAEL")
    for u in ISRAEL:
        page(u)
    for q in ["natural gas", "גז טבעי", "חשמל",
              "electricity", "ייצור חשמל"]:
        ckan("https://data.gov.il", q)
    out("\n##### EGYPT")
    for u in EGYPT:
        page(u)
    for q in ["electricity", "natural gas"]:
        ckan("https://egypt-odp.portaljs.com", q)


# ---------------------------------------------------------------- round 2
def body(url, n=600):
    out(f"\n=== BODY {url}")
    r = get(url)
    if r is None:
        return None
    r.encoding = r.apparent_encoding or r.encoding
    out(f"  {r.status_code} {r.headers.get('content-type','')[:40]} {len(r.content)}B final={r.url[:120]}")
    out("  ", " ".join(r.text[:n].split()))
    for m in re.findall(r'(?:location(?:\.href)?\s*=|href=|src=|url=)\s*["\']?([^"\'\s>;]+)', r.text, re.I)[:15]:
        out("   ->", m[:150])
    return r


def sitemap_filter(url, pat):
    out(f"\n=== SITEMAP {url}")
    r = get(url)
    if r is None:
        return
    locs = re.findall(r"<loc>(.*?)</loc>", r.text)
    out(f"  {len(locs)} urls")
    from urllib.parse import unquote
    for l in locs:
        u = unquote(l)
        if re.search(pat, u, re.I):
            out("   ", u[:170])


def pw_pages(urls, show_links=True):
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        ctx = b.new_context(user_agent=H["User-Agent"], locale="en-US")
        for url in urls:
            pg = ctx.new_page()
            xhr = []
            pg.on("response", lambda resp, xhr=xhr: xhr.append(resp)
                  if resp.request.resource_type in ("xhr", "fetch") else None)
            out(f"\n=== PW {url}")
            try:
                pg.goto(url, wait_until="domcontentloaded", timeout=60000)
                try:
                    pg.wait_for_load_state("networkidle", timeout=25000)
                except Exception:
                    pg.wait_for_timeout(8000)
                out(f"  title={pg.title()[:90]!r} final={pg.url[:120]}")
                txt = " ".join(pg.inner_text("body")[:1500].split())
                out("  text:", txt[:900])
                if show_links:
                    links = pg.eval_on_selector_all(
                        "a[href]", "els => els.map(e => [e.href, (e.innerText||'').trim().slice(0,60)])")
                    seen = set()
                    k = 0
                    for h, t in links:
                        if h in seen or not (KEY.search(h) or KEY.search(t)):
                            continue
                        if re.search(r"facebook|twitter|linkedin|youtube|instagram|whatsapp|mailto", h):
                            continue
                        seen.add(h)
                        out(f"   - {h[:160]} | {' '.join(t.split())}")
                        k += 1
                        if k > 70:
                            break
                for resp in xhr[:25]:
                    try:
                        b2 = resp.body()
                    except Exception:
                        b2 = b""
                    out(f"  XHR {resp.status} {resp.request.method} {resp.url[:170]} "
                        f"{resp.headers.get('content-type','')[:30]} {len(b2)}B")
                    if b2 and len(b2) < 400000 and resp.status == 200:
                        out("      ", " ".join(b2[:300].decode('utf-8', 'replace').split()))
            except Exception as e:
                out(f"  PW ERROR {type(e).__name__}: {str(e)[:200]}")
            pg.close()
        b.close()


def pdf_text(url, pages=4):
    out(f"\n=== PDF {url}")
    r = get(url)
    if r is None:
        return
    out(f"  {r.status_code} {r.headers.get('content-type','')[:30]} {len(r.content)}B")
    if r.status_code != 200 or not r.content.startswith(b"%PDF"):
        return
    import io
    import pdfplumber
    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
        out(f"  pages={len(pdf.pages)}")
        for i, pgx in enumerate(pdf.pages[:pages]):
            out(f"  -- p{i+1}:", " ".join((pgx.extract_text() or "")[:1500].split()))


def round2():
    out("##### ISRAEL")
    sitemap_filter("https://www.noga-iso.co.il/sitemap.xml",
                   r"stat|pdt|production|consum|demand|data|report|xls|csv|media/|systemoperation")
    pdf_text("https://www.noga-iso.co.il/media/smsdo1bd/api-smp-developer-guide-v2.pdf", pages=6)
    pw_pages([
        "https://www.noga-iso.co.il/pdt/highvoltageproduction/",
        "https://www.noga-iso.co.il/pdt/power-consumption/",
        "https://www.noga-iso.co.il/statistics/",
        "https://www.noga-iso.co.il/statistics/annual-consumption-data/",
        "https://www.noga-iso.co.il/statistics/annual-hourly-consumption-data/",
        "https://www.noga-iso.co.il/systemoperationunit/demand-curve/",
        "https://www.gov.il/en/departments/ministry_of_energy/govil-landing-page",
        "https://www.gov.il/he/departments/topics/natural_gas_sector/govil-landing-page",
        "https://www.iec.co.il/en/about/statistics",
        "https://www.cbs.gov.il/en/subjects/Pages/Energy.aspx",
    ])
    out("\n##### EGYPT")
    for u in ["https://www.eehc.gov.eg/eehcportal/", "https://www.eehc.gov.eg/", "http://www.moee.gov.eg/",
              "https://www.capmas.gov.eg/", "https://www.idsc.gov.eg/"]:
        body(u)
    pw_pages([
        "https://www.eehc.gov.eg/eehcportal/",
        "http://www.moee.gov.eg/",
        "http://www.eetc.net.eg/",
        "https://egyptera.org/en/ElectricityObservatory.aspx",
        "https://egyptera.org/ar/ElectricityObservatory.aspx",
        "https://egyptera.org/en/GenEvaluation.aspx",
        "https://egyptera.org/en/annualReport.aspx",
        "https://egyptera.org/en/PeriodicalBooks2022.aspx",
        "https://www.egas.com.eg/w/production-and-development",
        "https://www.egas.com.eg/egas-company-profile",
        "https://www.petroleum.gov.eg/en/gas-and-petrol/discovery-search-production/Pages/gas.aspx",
        "https://www.capmas.gov.eg/",
    ])


if __name__ == "__main__":
    round2()
