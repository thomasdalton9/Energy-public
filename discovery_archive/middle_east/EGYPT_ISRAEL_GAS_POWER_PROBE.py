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
RUN 2 (Playwright): Noga sitemap has /systemoperationunit/demand-curve/,
/piechartspage/, /co2/, /demandforecastnew/. Pages call JSON endpoints
/umbraco/Noga/DemandGraph/Get?date=YYYY-MM-DD, /umbraco/Noga/CityTotalConsumption/Get,
/umbraco/Noga/CityElectricityUsage/Get - all 403 (1.5 KB html) even in the
browser (reCAPTCHA v3 key on the page). SMP API (APIM) needs sign-up + token,
Israeli users only. gov.il: Cloudflare challenge / "you have been blocked" in
Playwright too; page calls openapi-gc.digital.gov.il/pub/cio/govil/rest/landingpage/...
iec.co.il statistics -> pageNotFound. CBS energy page reachable (SharePoint).
EGYPT: EgyptERA ElectricityObservatory.aspx = MONTHLY "electricity observatory"
PDFs, Aug 2020 .. May 2023 (AR page; EN page to Aug 2022), URLs like
/ar/download/Report/Elec_ob_May2023.pdf (irregular names). Annual reports
2011-2021 PDFs (/en/download/report/AnnualRep2021.pdf). EEHC portal is a
landing page -> /CMSeehc/. EETC times out again. CAPMAS is a React app over
a JSON API at https://www.capmas.gov.eg:8080/api/ (Subject, PublicationDetail,
Calendar ...). EGAS: only a 2025 company-profile PDF, no statistics.
RUN 3: Noga's JS (sb/noga-js-bundle.js, production-mix-js-bundle.js) calls
/umbraco/Noga/ProductionMixGraph/Get?date=, DemandGraph/Get?date=, Co2Graph/Get?date=
and Excel exports /umbraco/Noga/<graphType>/Get?fromDate=&toDate= (blob), all
with header X-Recaptcha-Token (reCAPTCHA v3). From Actions they return 403
(same IIS page) with and without a fresh token in headless Chromium.
CBS series API (apis.cbs.gov.il/series) reachable: catalog level 1 lists
subject 42 "Energy (data from 2013 on)". gov.il CollectorsWebApi: Cloudflare 403;
openapi-gc.digital.gov.il landing-page API: 500.
EGYPT: moee.gov.eg/english_new/home.aspx (http) shows "Grid Loads <date>:
Yesterday Peak Load 33300 MW, Min Load 23129 MW, Temperature, Today Expected
Evening Peak" (daily snapshot) and links Statistics (St_main.aspx) and EEHC
annual reports (report.aspx). EgyptERA observatory PDFs are scanned images
(no text layer, 10 MB, 20 pages). CAPMAS data menu has subSubject 30 Energy
(/data/mainSubject/3/subSubject/30). EETC still times out.
RUN 4: Noga 403 body = Azure Front Door WAF "Service unavailable - The request
is blocked" (x-azure-ref header): the data endpoints are blocked for GitHub
runner IPs (pages themselves load) -> Noga NOT usable from Actions.
CBS series catalog under 42: 42,1 electricity (production by fuel: total, diesel,
fossil total, fuel oil, coal, natural gas, renewables, wind, hydro, biogas,
solar, oil shale, waste heat), 42,2 energy balance, 42,3 indicators, 42,4
sources & uses.
EGYPT: MOEE statistics pages (http, English): ST_load (peak load MW),
ST_taka (energy consumption), ST_energy2 (energy sold by purpose), ST_fuel
(fuel consumption ktoe), ST_capacity, ST_energy (generation GWh by type),
ST_gas (gas share of fuel). CAPMAS energy sub-subject 30 lists ~70 indicator
visualisations via api/Subject/SubSubjectWithIndicator/30 and
api/Indicator/IndicatorDetails.
RUN 5: MOEE statistics pages are plain HTML tables (pandas.read_html works),
ANNUAL by Egyptian fiscal year (Jul-Jun), latest 2023/2022 (FY2022/23):
ST_energy generation GWh Thermal/Hydro/Wind/Solar/total from 1999/2000;
ST_load peak load MW from 1995/96 (gaps 2018-2020); ST_taka energy consumption
GWh; ST_energy2 energy sold by purpose (industry, agriculture, residential,
commercial, government, interconnection) from ~2000; ST_fuel fuel used in power
(HFO kt, natural gas million m3, ...) from 1998/99 = gas demand of the power
sector; ST_gas gas share (to 2020/21); ST_capacity MW by type to 2021/22.
report.aspx no longer lists EEHC annual report PDFs (only NEAP 2026-29).
CAPMAS api/Subject/SubSubjectWithIndicator/30 -> key publicationWithIndicators
(structure to inspect); api/Indicator/IndicatorDetails/<id> 404.
CBS series API works fully: catalog/level?id=3&subject=42 lists 42,1,1 electricity
production by producer and energy type, 42,1,2 fuel use for power, 42,1,3
supply by use, 42,1,5 capacity, 42,1,6 peak; 42,2 energy balance (natural gas
42,2,2,11). data/path?id=42,1,1,1,4060 -> JSON series 650046 (GWh, gas-fired
generation) ANNUAL, 2013-2023, updated 2025-02-25; energy balance to 2024
(updated 2026-01-19). Annual and ~1-2 years behind.

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


def js_endpoints(page_url, pat):
    """Fetch a page, then its scripts, and print lines matching pat."""
    out(f"\n=== JS-SCAN {page_url}")
    r = get(page_url)
    if r is None:
        return
    srcs = re.findall(r'<script[^>]+src=["\']([^"\']+)', r.text, re.I)
    inline = re.findall(r"<script[^>]*>(.*?)</script>", r.text, re.S | re.I)
    blobs = [("inline", "\n".join(inline))]
    for src in srcs:
        u = urljoin(r.url, src)
        if re.search(r"google|jquery|bootstrap|gtag|recaptcha|cookie|nagich|userway", u, re.I):
            continue
        rr = get(u)
        if rr is not None and rr.status_code == 200:
            blobs.append((u, rr.text))
    for name, txt in blobs:
        hits = [m.start() for m in re.finditer(pat, txt)]
        if hits:
            out(f"  [{name[:110]}] {len(txt)}B {len(hits)} hits")
        for h in hits[:8]:
            out("     ...", " ".join(txt[max(0, h - 250):h + 350].split()))


def try_api(url, headers=None, n=500):
    hh = dict(H)
    hh.update(headers or {})
    try:
        r = requests.get(url, headers=hh, timeout=T)
    except requests.RequestException as e:
        out(f"  API {url[:150]} ERROR {type(e).__name__}")
        return None
    out(f"  API {r.status_code} {r.headers.get('content-type','')[:30]} {len(r.content)}B {url[:160]}")
    out("      ", " ".join(r.text[:n].split()))
    return r


def noga_pw():
    """Accept cookies, then load the data pages and log every XHR (incl. request headers)."""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        ctx = b.new_context(user_agent=H["User-Agent"], locale="he-IL")
        for url in ["https://www.noga-iso.co.il/systemoperationunit/demand-curve/",
                    "https://www.noga-iso.co.il/systemoperationunit/piechartspage/",
                    "https://www.noga-iso.co.il/systemoperationunit/co2/",
                    "https://www.noga-iso.co.il/systemoperationunit/demandforecastnew/"]:
            pg = ctx.new_page()
            xhr = []
            pg.on("response", lambda resp, xhr=xhr: xhr.append(resp)
                  if "noga-iso" in resp.url and resp.request.resource_type in ("xhr", "fetch") else None)
            out(f"\n=== NOGA-PW {url}")
            try:
                pg.goto(url, wait_until="domcontentloaded", timeout=60000)
                for sel in ["text=קבלת עוגיות", "button:has-text('קבל')"]:
                    try:
                        pg.click(sel, timeout=4000)
                        out("  clicked cookie accept")
                        break
                    except Exception:
                        pass
                pg.wait_for_timeout(9000)
                txt = " ".join(pg.inner_text("main")[:1200].split()) if pg.query_selector("main") else ""
                out("  main text:", txt[:700])
                for resp in xhr[:12]:
                    try:
                        bb = resp.body()
                    except Exception:
                        bb = b""
                    rh = resp.request.headers
                    out(f"  XHR {resp.status} {resp.request.method} {resp.url[:160]} {len(bb)}B "
                        f"reqhdr={ {k: v[:60] for k, v in rh.items() if k.lower() not in ('user-agent','accept-language','sec-ch-ua','sec-ch-ua-mobile','sec-ch-ua-platform')} }")
                    out("      ", " ".join(bb[:400].decode("utf-8", "replace").split()))
            except Exception as e:
                out(f"  ERROR {type(e).__name__}: {str(e)[:200]}")
            pg.close()
        b.close()


def round3():
    out("##### ISRAEL")
    js_endpoints("https://www.noga-iso.co.il/systemoperationunit/demand-curve/", r"umbraco/Noga|recaptcha|grecaptcha\.execute")
    js_endpoints("https://www.noga-iso.co.il/systemoperationunit/piechartspage/", r"umbraco/Noga")
    for path in ["DemandGraph/Get?date=2026-10-10", "PieCharts/Get?date=2026-10-10",
                 "ProductionMix/Get?date=2026-10-10"]:
        try_api("https://www.noga-iso.co.il/umbraco/Noga/" + path,
                {"Referer": "https://www.noga-iso.co.il/systemoperationunit/demand-curve/",
                 "X-Requested-With": "XMLHttpRequest", "Accept": "application/json"})
    noga_pw()
    out("\n--- CBS series API")
    for u in ["https://apis.cbs.gov.il/series/catalog/level?id=1&format=json&download=false",
              "https://apis.cbs.gov.il/series/catalog/level?id=2&format=json&download=false&subject=",
              "https://apis.cbs.gov.il/series/catalog/subjects?format=json"]:
        try_api(u, n=1500)
    out("\n--- gov.il open API host")
    for u in ["https://openapi-gc.digital.gov.il/pub/cio/govil/rest/landingpage/v1/he/api/topics/get/natural_gas_sector",
              "https://openapi-gc.digital.gov.il/pub/cio/govil/rest/landingpage/v1/en/api/departments/get/ministry_of_energy",
              "https://www.gov.il/CollectorsWebApi/api/DataCollector/GetResults?CollectorType=reports&officeId=4c3b0a6f-5ac8-4b6c-a0a0-0b2b1d2d5b5c&skip=0&limit=10"]:
        try_api(u, n=800)

    out("\n##### EGYPT")
    pdf_text("https://egyptera.org/ar/download/Report/Elec_ob_May2023.pdf", pages=5)
    pdf_text("https://egyptera.org/en/download/Report/Elec_Ob_August2022.pdf", pages=5)
    pdf_text("https://egyptera.org/en/download/report/AnnualRep2021.pdf", pages=0)
    for u in ["https://www.capmas.gov.eg:8080/api/Subject",
              "https://www.capmas.gov.eg:8080/api/PublicationDetail/HomePublicationDetail"]:
        try_api(u, n=3000)
    js_endpoints("https://www.capmas.gov.eg/", r"[\"'`]/?api/[A-Za-z/]+")
    pw_pages(["https://www.eehc.gov.eg/CMSeehc/", "http://www.moee.gov.eg/english_new/home.aspx",
              "https://www.capmas.gov.eg/data"])
    for u in ["https://www.eetc.net.eg/", "http://eetc.net.eg/"]:
        page(u)


def cbs_walk():
    base = "https://apis.cbs.gov.il/series/catalog/"
    out("\n--- CBS energy catalog (subject 42)")
    def lvl(level, subject):
        u = f"{base}level?id={level}&subject={subject}&format=json&download=false"
        try:
            r = requests.get(u, headers=H, timeout=T)
            js = r.json()
        except Exception as e:
            out(f"  {u} -> {type(e).__name__}")
            return []
        cat = (js.get("catalogs") or {}).get("catalog") or []
        return cat
    for c2 in lvl(2, "42"):
        p2 = ",".join(map(str, c2["path"]))
        out(f"  L2 {p2} {c2['name']}")
        for c3 in lvl(3, p2):
            p3 = ",".join(map(str, c3["path"]))
            out(f"    L3 {p3} {c3['name']}")
            for c4 in lvl(4, p3)[:20]:
                p4 = ",".join(map(str, c4["path"]))
                out(f"      L4 {p4} {c4['name']}")
    for path in ["42,1", "42,1,1", "42,2,1", "42,3,1"]:
        u = f"{base}path?id={path}&format=json&download=false"
        try:
            r = requests.get(u, headers=H, timeout=T)
            out(f"  PATH {path}: {r.status_code} {len(r.content)}B", " ".join(r.text[:1500].split()))
            sid = re.findall(r'"id"\s*:\s*"?(\d{4,})', r.text)
            if sid:
                du = (f"https://apis.cbs.gov.il/series/data/list?id={sid[0]}&startPeriod=01-2013"
                      f"&endPeriod=12-2026&format=json&download=false")
                try_api(du, n=1200)
        except Exception as e:
            out(f"  PATH {path} {type(e).__name__}")


def round4():
    out("##### ISRAEL")
    r = get("https://www.noga-iso.co.il/umbraco/Noga/ProductionMixGraph/Get?date=2026-10-10")
    if r is not None:
        out("  Noga 403 page text:", " ".join(re.sub(r"<[^>]+>", " ", r.text).split())[:600])
        out("  headers:", dict(r.headers))
    cbs_walk()
    out("\n##### EGYPT")
    r = get("http://www.moee.gov.eg/english_new/home.aspx")
    if r is not None:
        i = r.text.find("Peak Load")
        out("  MOEE grid-load html:", " ".join(r.text[max(0, i - 1500):i + 1500].split()))
    for u in ["http://www.moee.gov.eg/english_new/St_main.aspx", "http://www.moee.gov.eg/english_new/report.aspx",
              "http://www.moee.gov.eg/arabic/home.aspx", "http://www.moee.gov.eg/test_new/home.aspx"]:
        page(u, max_links=80)
    js_endpoints("https://www.capmas.gov.eg/", r"[A-Z_]+:\"api/(?:Publication|Subject|Indicator|Data|Bulletin|Release)[A-Za-z/]*\"")
    pw_pages(["https://www.capmas.gov.eg/data/mainSubject/3/subSubject/30"])


def html_tables(url, maxrows=25):
    out(f"\n=== TABLES {url}")
    r = get(url)
    if r is None:
        return
    out(f"  {r.status_code} {len(r.content)}B")
    import pandas as pd
    from io import StringIO
    try:
        tabs = pd.read_html(StringIO(r.text))
    except Exception as e:
        out(f"  no tables: {e}")
        return
    for i, t in enumerate(tabs):
        if t.shape[0] < 2 and t.shape[1] < 2:
            continue
        out(f"  table {i} shape={t.shape}")
        out(t.head(maxrows).to_string(max_colwidth=28)[:3500])


def capmas():
    api = "https://www.capmas.gov.eg:8080/api/"
    r = get(api + "Subject/SubSubjectWithIndicator/30")
    if r is None:
        return
    js = r.json()
    d = js.get("data") or {}
    keys = list(d.keys())
    out("  keys:", keys)
    inds = d.get("indicators") or d.get("publicationIndicators") or []
    if not inds:
        for k in keys:
            if isinstance(d[k], list) and d[k] and isinstance(d[k][0], dict):
                inds = d[k]
                out("  using list key", k)
                break
    out(f"  {len(inds)} indicators; first keys: {list(inds[0].keys()) if inds else None}")
    for it in inds[:90]:
        tr = it.get("indicatorTranslations") or it.get("publicationIndicatorTranslations") or []
        en = next((t.get("title") or t.get("name") for t in tr if t.get("locale") == "en"), "")
        out(f"   {it.get('id')} | {(it.get('title') or it.get('name') or '')[:70]} | {en[:90]} | "
            f"{ {k: it.get(k) for k in ('periodicity','periodic','unit','studyPeriod','year','lastUpdate') if k in it} }")
    for iid in ["3092", "9487", "3153", "4800"]:
        try_api(api + "Indicator/IndicatorDetails/" + iid, n=2500)


def round5():
    out("##### EGYPT")
    for st in ["ST_energy", "ST_load", "ST_taka", "ST_energy2", "ST_fuel", "ST_gas", "ST_capacity"]:
        html_tables(f"http://www.moee.gov.eg/english_new/{st}.aspx")
    r = get("http://www.moee.gov.eg/english_new/report.aspx")
    if r is not None:
        out("  report.aspx pdf/doc refs:", sorted(set(re.findall(r'["\']([^"\']+\.(?:pdf|PDF|xlsx?|docx?))["\']', r.text)))[:60])
    r = get("http://www.moee.gov.eg/test_new/report.aspx")
    if r is not None:
        out("  test_new/report.aspx refs:", sorted(set(re.findall(r'["\']([^"\']+\.(?:pdf|PDF|xlsx?))["\']', r.text)))[:60])
    capmas()
    out("\n##### ISRAEL (CBS)")
    base = "https://apis.cbs.gov.il/series/"
    for u in [base + "catalog/level?id=3&subject=42&format=json&download=false",
              base + "catalog/level?id=4&subject=42&format=json&download=false",
              base + "catalog/path?id=42,1,1,1,4060&format=json&download=false",
              base + "data/path?id=42,1,1,1,4060&startPeriod=01-2020&endPeriod=12-2026&format=json&download=false",
              base + "data/path?id=42,2,1,1,4173&startPeriod=01-2013&endPeriod=12-2026&format=json&download=false",
              base + "catalog/path?id=42,4&format=json&download=false"]:
        try_api(u, n=2500)


def round6():
    out("##### EGYPT CAPMAS")
    r = get("https://www.capmas.gov.eg:8080/api/Subject/SubSubjectWithIndicator/30")
    if r is not None:
        pubs = (r.json().get("data") or {}).get("publicationWithIndicators") or []
        out(f"  {len(pubs)} publications")
        for pb in pubs[:30]:
            if not isinstance(pb, dict):
                continue
            small = {k: v for k, v in pb.items() if not isinstance(v, (list, dict))}
            out("  PUB", json.dumps(small, ensure_ascii=False)[:400])
            for k, v in pb.items():
                if isinstance(v, list):
                    for it in v[:40]:
                        if isinstance(it, dict):
                            sm = {kk: vv for kk, vv in it.items() if not isinstance(vv, (list, dict))}
                            out("     -", k, json.dumps(sm, ensure_ascii=False)[:260])
                elif isinstance(v, dict):
                    out("     *", k, json.dumps(v, ensure_ascii=False)[:300])
    pw_pages(["https://www.capmas.gov.eg/data/mainSubject/3/subSubject/30/data-visualization/3092",
              "https://www.capmas.gov.eg/data/mainSubject/3/subSubject/30/data-visualization/9487"],
             show_links=False)

    out("\n##### ISRAEL")
    out("--- CBS: search all subjects (levels 2-3) for electricity / gas")
    base = "https://apis.cbs.gov.il/series/catalog/"
    r = get(base + "level?id=1&format=json&download=false")
    subs = (r.json().get("catalogs") or {}).get("catalog") or [] if r is not None else []
    pat = re.compile("\u05d7\u05e9\u05de\u05dc|\u05d2\u05d6|\u05d0\u05e0\u05e8\u05d2|\u05d3\u05dc\u05e7")
    for sb in subs:
        sid = sb["path"][0]
        if sid == 42:
            continue
        for lv in (2, 3, 4):
            try:
                rr = requests.get(f"{base}level?id={lv}&subject={sid}&format=json&download=false",
                                  headers=H, timeout=T)
                cat = (rr.json().get("catalogs") or {}).get("catalog") or []
            except Exception:
                cat = []
            for c in cat:
                if pat.search(c.get("name") or ""):
                    out(f"  subj {sid} {sb['name'][:30]} L{lv} {c['path']} {c['name']}")
    for path in ["42,1,3", "42,1,6"]:
        try_api(f"https://apis.cbs.gov.il/series/data/path?id={path}&startPeriod=01-2013&endPeriod=12-2026"
                f"&format=json&download=false", n=700)
    out("\n--- gov.il files / alternatives")
    for u in ["https://www.gov.il/BlobFolder/reports/natural_gas_sector_data/he/test.pdf",
              "https://www.gov.il/he/api/PublicationApi/Index?limit=10&OfficeId=ministry_of_energy"]:
        try_api(u, n=200)
    for u in ["https://newmedenergy.com/investor-relations/financial-reports/",
              "https://newmedenergy.com/investor-relations/presentations/",
              "https://newmedenergy.com/investor-relations/leviathan-bond-reports/",
              "https://www.energean.com/investors/results-reports-and-presentations/",
              "https://www.tamarpetroleum.co.il/en/"]:
        page(u, max_links=25)


if __name__ == "__main__":
    round6()
