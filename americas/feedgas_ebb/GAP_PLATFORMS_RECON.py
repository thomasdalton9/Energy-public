"""
Recon for the feedgas gap platforms, with network capture - these are
single-page apps whose data comes from JSON calls the page makes, so
the most robust scraper is usually to call those endpoints directly.

  - Cheniere LNG Connection (lngconnection.cheniere.com): its home page
    dashboard already showed Creole Trail's "CT200111-CREOLE TRAIL-
    SPLIQ-D" delivery into Sabine Pass Liquefaction, but not which gas
    day. Here: open Capacity > Operationally Available for Creole Trail
    (company 200) and Corpus Christi (company 400), dump the tables and
    controls, and log every JSON/XHR response URL.
  - TC Energy (ANR, Columbia Gulf): the first attempt got connection
    resets from ebb.anrpl.com / ebb.tceconnects.com - retried here with
    the TCeConnects URL form found on the web, plus plain requests.
  - BHE GT&S (Cove Point): first attempt was cut off by a redirect from
    the page before; retried on its own.

Operator-hosted postings only (see EBB_PLATFORM_RECON.py). Outputs:
gap_platforms_recon_output/.
"""

import json
import os
import re
import sys

import requests

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("Playwright isn't installed: pip install playwright && playwright install chromium", file=sys.stderr)
    sys.exit(1)

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gap_platforms_recon_output")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

TABLES_JS = """() => [...document.querySelectorAll('table')].map(t =>
    [...t.rows].slice(0, 25).map(r => [...r.cells].map(c => c.innerText.replace(/\\s+/g, ' ').trim())))
    .filter(t => t.length > 1)"""
CONTROLS_JS = """() => [...document.querySelectorAll('input,select,button')].filter(e => e.type !== 'hidden')
    .map(e => `${e.tagName} id=${e.id} name=${e.name || ''} type=${e.type || ''} value=${(e.value || '').slice(0, 30)}`
              + (e.tagName === 'SELECT' ? ' options=' + [...e.options].map(o => o.text).slice(0, 10).join('|') : ''))
    .slice(0, 40)"""


def log(msg=""):
    print(msg, flush=True)


def dump_page(page, tag):
    page.screenshot(path=os.path.join(OUTPUT_DIR, f"{tag}.png"), full_page=True)
    with open(os.path.join(OUTPUT_DIR, f"{tag}.html"), "w", encoding="utf-8") as f:
        f.write(page.content())
    log(f"  url: {page.url}")
    log(f"  text: {' '.join(page.inner_text('body').split())[:500]!r}")
    for c in page.evaluate(CONTROLS_JS):
        log(f"  CONTROL {c}")
    for t in page.evaluate(TABLES_JS):
        log(f"  TABLE ({len(t)} rows shown)")
        for r in t[:14]:
            log(f"    {r}")


def cheniere(context):
    page = context.new_page()
    api_calls = []

    def on_response(resp):
        ct = resp.headers.get("content-type", "")
        if "json" in ct or "/api/" in resp.url.lower() or resp.request.resource_type in ("xhr", "fetch"):
            try:
                body = resp.text()[:400]
            except Exception:
                body = "?"
            api_calls.append((resp.status, resp.request.method, resp.url, ct, body))

    page.on("response", on_response)
    log("\n==================== Cheniere LNG Connection ====================")
    page.goto("https://lngconnection.cheniere.com/", wait_until="networkidle", timeout=60000)
    page.wait_for_timeout(4000)
    if page.locator("#cookieAcceptBtn").count():
        page.locator("#cookieAcceptBtn").click()
    for company, label in (("200", "creole_trail"), ("400", "corpus_christi")):
        log(f"\n--- {label} (companyradio={company})")
        try:
            page.locator(f"input[name='companyradio'][value='{company}']").check(force=True)
            page.wait_for_timeout(3000)
            dump_page(page, f"cheniere_{label}_home")
            page.locator("a", has_text="Operationally Available").first.click(force=True)
            page.wait_for_load_state("networkidle", timeout=60000)
            page.wait_for_timeout(5000)
            dump_page(page, f"cheniere_{label}_oac")
        except Exception as e:
            log(f"  FAILED: {type(e).__name__}: {str(e)[:300]}")
    log(f"\n--- Cheniere XHR/JSON responses ({len(api_calls)})")
    seen = set()
    for status, method, url, ct, body in api_calls:
        key = (method, re.sub(r"\d{4,}", "N", url))
        if key in seen:
            continue
        seen.add(key)
        log(f"  {status} {method} {url} [{ct}]")
        log(f"      {body[:300]!r}")
    with open(os.path.join(OUTPUT_DIR, "cheniere_api_calls.json"), "w") as f:
        json.dump(api_calls, f, indent=1)
    page.close()


def plain(url):
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=30)
        log(f"  requests GET {url} -> {r.status_code}, {len(r.content):,} bytes, title "
            f"{(re.search(r'<title>(.*?)</title>', r.text, re.S | re.I) or [None, ''])[1].strip()[:80]!r}")
    except requests.RequestException as e:
        log(f"  requests GET {url} FAILED: {type(e).__name__}: {str(e)[:150]}")


def simple(context, name, urls):
    log(f"\n==================== {name} ====================")
    for url in urls:
        plain(url)
        page = context.new_page()
        log(f"\n--- browser {url}")
        try:
            page.goto(url, wait_until="networkidle", timeout=60000)
            page.wait_for_timeout(4000)
            dump_page(page, f"{name}_{re.sub(r'[^A-Za-z0-9]+', '_', url)[8:80]}")
            links = page.evaluate("""() => [...document.querySelectorAll('a[href]')]
                .filter(a => /capacit|operational|avail|location/i.test(a.innerText + a.href))
                .map(a => a.innerText.trim().slice(0, 60) + ' -> ' + a.href).slice(0, 30)""")
            for l in links:
                log(f"  LINK {l}")
        except Exception as e:
            log(f"  FAILED: {type(e).__name__}: {str(e)[:300]}")
        page.close()


TC_BASE = "https://ebb.tceconnects.com/infopost/"
# Second pass findings: TC eConnects reports are SQL Server Reporting
# Services (ReportViewer.aspx?/InfoPost/<Report>&pAssetNbr=N), which can
# export CSV with rs:Format=CSV. Asset numbers from the menu: ANR 3005,
# Columbia Gulf 14, Columbia Gas (TCO) 51.
TC_REPORTS = [
    ("ANR", "OperationallyAvailableCapacityANR", "pAssetNbr=3005"),
    ("CGT", "OperationallyAvailableCapacity", "pAssetNbr=14"),
    ("TCO", "OperationallyAvailableCapacity", "pAssetNbr=51"),
    ("ANR_power", "ScheduledQtyForPowerPlants", "pAssetNbr=3005"),
    ("ANR_locations", "LocationDataDownload", "assetNbr=3005"),
]
LNG_RE = re.compile(r"LNG|LIQ|CAMERON|TRANSCAM|CALCAS|VENTURE|GATOR|PLAQUEM|COVE|SABINE|CREOLE|GILLIS", re.I)


def tc_csv(context):
    log("\n==================== TC Energy CSV exports ====================")
    page = context.new_page()
    page.goto(TC_BASE + "TCeConnects.aspx?v=1.3&SID=67&info=Y&assetid=3005", wait_until="networkidle", timeout=60000)
    for tag, report, asset in TC_REPORTS:
        url = f"{TC_BASE}ReportViewer.aspx?/InfoPost/{report}&{asset}&rs:Format=CSV"
        log(f"\n--- {tag}: {url}")
        try:
            r = context.request.get(url, timeout=120000)
            text = r.text()
            log(f"  HTTP {r.status}, {len(text):,} chars, {r.headers.get('content-type')}")
            with open(os.path.join(OUTPUT_DIR, f"tc_{tag}.csv"), "w", encoding="utf-8") as f:
                f.write(text)
            lines = text.splitlines()
            for line in lines[:6]:
                log(f"    {line[:300]}")
            hits = [l for l in lines if LNG_RE.search(l)]
            log(f"  {len(lines)} lines, {len(hits)} LNG-ish:")
            for l in hits[:25]:
                log(f"    {l[:300]}")
        except Exception as e:
            log(f"  FAILED: {type(e).__name__}: {str(e)[:300]}")
    page.close()


def bhe(context):
    simple(context, "bhe2", [
        "https://infopost.bhegts.com/cpl",
        "https://ebb.bhegts.com/InformationalPostings/default.aspx/",
    ])


def capture_page(context, tag, url, clicks=()):
    """Load a (JS app) page, optionally click through, dump tables and
    every JSON/XHR response - the data endpoints behind the page."""
    page = context.new_page()
    calls = []

    def on_response(resp):
        if resp.request.resource_type in ("xhr", "fetch") or "json" in resp.headers.get("content-type", ""):
            try:
                body = resp.text()[:500]
            except Exception:
                body = "?"
            calls.append((resp.status, resp.request.method, resp.url, body, resp.request.post_data))

    page.on("response", on_response)
    log(f"\n--- {tag}: {url}")
    try:
        page.goto(url, wait_until="networkidle", timeout=90000)
        page.wait_for_timeout(5000)
        for sel in clicks:
            page.locator(sel).first.click(timeout=15000)
            page.wait_for_load_state("networkidle", timeout=60000)
            page.wait_for_timeout(3000)
        dump_page(page, tag)
    except Exception as e:
        log(f"  FAILED: {type(e).__name__}: {str(e)[:300]}")
    log(f"  {len(calls)} XHR/JSON responses:")
    for status, method, u, body, post in calls[:30]:
        log(f"    {status} {method} {u}")
        if post:
            log(f"      POST {post[:300]!r}")
        log(f"      {body[:300]!r}")
    with open(os.path.join(OUTPUT_DIR, f"{tag}_calls.json"), "w") as f:
        json.dump(calls, f, indent=1)
    page.close()


def tc_params(context):
    """ReportViewer HTML for the OA reports - its parameter area shows the
    parameter names needed to ask for a specific gas day/cycle (the CSV
    export defaults to the NEXT gas day's Timely cycle). Also Columbia
    Gulf's location file, to confirm who is behind Wilkinson Bayou and
    'Cameron LNG'."""
    log("\n==================== TC Energy report parameters ====================")
    page = context.new_page()
    page.goto(TC_BASE + "TCeConnects.aspx?v=1.3&SID=67&info=Y&assetid=3005", wait_until="networkidle", timeout=60000)
    for tag, report, asset in (("ANR", "OperationallyAvailableCapacityANR", "pAssetNbr=3005"),
                               ("CGT", "OperationallyAvailableCapacity", "pAssetNbr=14")):
        url = f"{TC_BASE}ReportViewer.aspx?/InfoPost/{report}&{asset}"
        log(f"\n--- {tag} viewer: {url}")
        try:
            page.goto(url, wait_until="networkidle", timeout=120000)
            page.wait_for_timeout(3000)
            h = page.content()
            with open(os.path.join(OUTPUT_DIR, f"tc_{tag}_viewer.html"), "w", encoding="utf-8") as f:
                f.write(h)
            names = sorted(set(re.findall(r'(?:name|id)="([^"]*(?:Param|param|ctl\d+_ctl\d+_ctl\d+_txtValue|ddValue)[^"]*)"', h)))
            log(f"  param-ish ids: {names[:40]}")
            labels = page.evaluate("""() => [...document.querySelectorAll('td.ParamLabelCell, label, span')]
                .map(e => e.innerText.trim()).filter(t => t && t.length < 40).slice(0, 60)""")
            log(f"  labels: {labels}")
            for c in page.evaluate(CONTROLS_JS):
                log(f"  CONTROL {c}")
        except Exception as e:
            log(f"  FAILED: {type(e).__name__}: {str(e)[:300]}")
    url = f"{TC_BASE}ReportViewer.aspx?/InfoPost/LocationDataDownload&assetNbr=14&rs:Format=CSV&rc:NoHeader=true"
    r = context.request.get(url, timeout=120000)
    text = r.text()
    with open(os.path.join(OUTPUT_DIR, "tc_CGT_locations.csv"), "w", encoding="utf-8") as f:
        f.write(text)
    log(f"\n--- CGT locations: HTTP {r.status}, {len(text):,} chars")
    for line in text.splitlines():
        if re.search(r"^[^,]*,[^,]*,[^,]*,[^,]*,(4267|4246|801)\b|WILKINSON|CAMERON LNG|GATOR|PLAQUEM|VENTURE", line, re.I):
            log(f"    {line[:400]}")
    page.close()


def enbridge_units(context):
    """Enbridge's portal lists business-unit codes in businessunits.json -
    Algonquin's is 'AG', not 'AGT', so BIG's rtba code is probably not
    'BIG' either (bu=BIG returned the generic error page)."""
    log("\n==================== Enbridge business units ====================")
    r = context.request.get("https://infopost.enbridge.com/assets/data/businessunits.json", timeout=60000)
    units = r.json().get("pipelines", [])
    for u in units:
        log(f"  {u}")
    page = context.new_page()
    for u in units:
        if "BIG" not in (u.get("abbreviation", "") + u.get("name", "")).upper():
            continue
        for code in {u.get("code"), u.get("abbreviation")} - {None}:
            url = f"https://rtba.enbridge.com/InformationalPosting/Default.aspx?bu={code}&Type=OA"
            log(f"\n--- BIG as {code}: {url}")
            try:
                page.goto(url, wait_until="networkidle", timeout=60000)
                log(f"  text: {' '.join(page.inner_text('body').split())[:300]!r}")
                if page.locator("a:has-text('Downloadable Format')").count():
                    with page.expect_download(timeout=90000) as info:
                        page.locator("a:has-text('Downloadable Format')").first.click()
                    text = open(info.value.path(), encoding="utf-8", errors="replace").read()
                    with open(os.path.join(OUTPUT_DIR, f"big_{code}_oa.csv"), "w", encoding="utf-8") as f:
                        f.write(text)
                    for line in text.splitlines()[:40]:
                        log(f"    {line[:250]}")
            except Exception as e:
                log(f"  FAILED: {type(e).__name__}: {str(e)[:300]}")
    page.close()


def gulfsouth_ui(context):
    """Boardwalk's GasQuest app signs its API calls with AWS Cognito
    credentials, so drive the UI instead: pick Gulf South from the TSP
    dropdown (a JS click - the menu item is hidden until the dropdown
    opens), open Operationally Available, capture tables and the API calls
    (with bodies) behind them - Gulf South's Stratton Ridge delivery is one
    of Freeport LNG's three interstate feeds."""
    page = context.new_page()
    calls = []

    def on_response(resp):
        if resp.request.resource_type in ("xhr", "fetch"):
            try:
                body = resp.text()[:1500] if "bwpmlp" in resp.url else ""
            except Exception:
                body = ""
            calls.append((resp.status, resp.request.method, resp.url, resp.request.post_data, body))
    page.on("response", on_response)
    log("\n==================== Gulf South (Boardwalk GasQuest) ====================")
    try:
        page.goto("https://infopost.bwpipelines.com/", wait_until="networkidle", timeout=90000)
        page.wait_for_timeout(4000)
        page.evaluate("""() => { const t = document.querySelector('[data-cy=tsp-wrapper-div] button, [data-cy=tsp-wrapper-div] .dropdown-toggle');
                                  if (t) t.click(); }""")
        page.wait_for_timeout(1500)
        page.evaluate("""() => document.querySelector('[data-cy^="078444247"]').click()""")
        page.wait_for_load_state("networkidle", timeout=60000)
        page.wait_for_timeout(5000)
        dump_page(page, "gulfsouth_tsp")
        links = page.evaluate("() => [...document.querySelectorAll('a')].map(a => [a.innerText.trim().slice(0, 60), a.href]).filter(x => x[0])")
        for t, h in links[:80]:
            log(f"    LINK {t!r} -> {h}")
        # "Operationally Available" sits under the Capacity menu
        try:
            page.get_by_text("Capacity", exact=True).first.click(timeout=10000)
            page.wait_for_timeout(1500)
            page.get_by_text(re.compile("Operationally Available", re.I)).first.click(timeout=10000)
        except Exception as e:
            log(f"  menu click failed ({type(e).__name__}); trying direct URLs")
            for path in ("capacity/operationally-available", "operationally-available",
                         "capacity/operationally-available-capacity", "operationally-available-capacity"):
                page.goto(f"https://www.gasquest.com/informational-posting/{path}", wait_until="networkidle", timeout=60000)
                page.wait_for_timeout(4000)
                if re.search("operationally available", page.inner_text("body"), re.I) and "404" not in page.inner_text("body")[:200]:
                    log(f"  direct URL worked: {page.url}")
                    break
        page.wait_for_load_state("networkidle", timeout=90000)
        page.wait_for_timeout(10000)
        dump_page(page, "gulfsouth_oac")
        menu = page.evaluate("() => [...document.querySelectorAll('a, button, li')].map(e => (e.innerText || '').trim()).filter(t => /capacity|available/i.test(t)).slice(0, 30)")
        log(f"  capacity menu items: {menu}")
        rows = page.evaluate("""() => [...document.querySelectorAll('tr')].map(tr => [...tr.querySelectorAll('td,th')].map(c => c.innerText.trim()))""")
        log(f"  {len(rows)} rows")
        for r in rows[:3]:
            log(f"    {r}")
        for r in rows:
            if re.search(r"STRATTON|FREEPORT|LNG", " ".join(r), re.I):
                log(f"    ROW {r}")
    except Exception as e:
        log(f"  FAILED: {type(e).__name__}: {str(e)[:300]}")
    for status, method, url, post, body in calls[-30:]:
        log(f"  CALL {status} {method} {url[:200]}")
        if post:
            log(f"    POST {post[:500]}")
        if body:
            log(f"    {body[:600]!r}")
    page.close()


def tc_param_probe(context):
    """Find the SSRS parameter names that set the gas day and cycle - the
    CSV's own EffGasDay/Cycle columns show whether a guess took effect."""
    log("\n==================== TC Energy parameter probe ====================")
    page = context.new_page()
    page.goto(TC_BASE + "TCeConnects.aspx?v=1.3&SID=67&info=Y&assetid=3005", wait_until="networkidle", timeout=60000)
    base = f"{TC_BASE}ReportViewer.aspx?/InfoPost/OperationallyAvailableCapacityANR&pAssetNbr=3005&rs:Format=CSV"
    day = "09/27/2026"
    guesses = [
        f"&pEffGasDay={day}&pCycle=ID3", f"&pGasDay={day}&pCycle=ID3", f"&pEffGasDate={day}&pCycle=ID3",
        f"&pBeginDate={day}", f"&pDate={day}", f"&EffGasDay={day}&Cycle=ID3", f"&pEffGasDay={day}",
    ]
    for g in guesses:
        try:
            r = context.request.get(base + g, timeout=120000)
            lines = r.text().splitlines()
            first = lines[1][:120] if len(lines) > 1 else lines[:1]
            log(f"  {g!r}: HTTP {r.status}, {len(lines)} lines, first row: {first!r}")
        except Exception as e:
            log(f"  {g!r}: FAILED {type(e).__name__}: {str(e)[:150]}")
    # the date-range variant, in case it takes begin/end dates
    for g in (f"&pBeginDate={day}&pEndDate={day}", f"&pStartDate={day}&pEndDate={day}", f"&pFromDate={day}&pToDate={day}"):
        url = f"{TC_BASE}ReportViewer.aspx?/InfoPost/OperationallyAvailableCapacityByDateRangeANR&pAssetNbr=3005&rs:Format=CSV{g}"
        try:
            r = context.request.get(url, timeout=120000)
            lines = r.text().splitlines()
            log(f"  daterange {g!r}: HTTP {r.status}, {len(lines)} lines, first row: {(lines[1][:120] if len(lines) > 1 else lines[:1])!r}")
        except Exception as e:
            log(f"  daterange {g!r}: FAILED {type(e).__name__}: {str(e)[:150]}")
    page.close()


def tc_param_probe2(context):
    """Wider search for the OA report's gas-day parameter (an unknown name
    makes SSRS error, so the right one is the guess that returns CSV for
    09/27), then open the viewer, set the gas day and print what its
    report area actually says (the daily run saw it land in an error)."""
    log("\n==================== TC Energy parameter probe 2 ====================")
    page = context.new_page()
    page.goto(TC_BASE + "TCeConnects.aspx?v=1.3&SID=67&info=Y&assetid=3005", wait_until="domcontentloaded", timeout=90000)
    page.wait_for_timeout(3000)
    base = f"{TC_BASE}ReportViewer.aspx?/InfoPost/OperationallyAvailableCapacityANR&pAssetNbr=3005&rs:Format=CSV"
    names = ["pGasDate", "GasDate", "pGasDt", "pEffDate", "pEffectiveDate", "pFlowDate", "pPostDate", "pPostingDate",
             "pReportDate", "pRptDate", "pStartDate", "pBegDate", "pGasFlowDate", "pDay", "pEffDt", "pEffGasDt",
             "pGasDateTime", "pEffGasDate", "EffGasDate", "pDate"]
    for name in names:
        for day in ("09/27/2026", "2026-09-27", "9/27/2026"):
            try:
                r = context.request.get(f"{base}&{name}={day}", timeout=120000)
                t = r.text().lstrip("\ufeff")
                if t.startswith("TSPName"):
                    rows = t.splitlines()
                    eff = rows[1].split(",")[2] if len(rows) > 1 else "?"
                    cyc = rows[1].split(",")[3] if len(rows) > 1 else "?"
                    log(f"  {name}={day}: CSV, EffGasDay {eff}, cycle {cyc}")
                else:
                    log(f"  {name}={day}: HTTP {r.status}, not CSV")
            except Exception as e:
                log(f"  {name}={day}: FAILED {type(e).__name__}: {str(e)[:120]}")
    # the viewer itself, with a date change
    page.goto(f"{TC_BASE}ReportViewer.aspx?/InfoPost/OperationallyAvailableCapacityANR&pAssetNbr=3005",
              wait_until="domcontentloaded", timeout=180000)
    page.locator("#txtGasDate").wait_for(timeout=120000)
    page.wait_for_timeout(20000)
    body = page.evaluate("() => document.body.innerText").split()
    log(f"  viewer before date change: {' '.join(body)[:1500]!r}")
    box = page.locator("#txtGasDate")
    box.fill("09/27/2026")
    box.dispatch_event("change")
    page.wait_for_timeout(45000)
    body = page.evaluate("() => document.body.innerText").split()
    log(f"  viewer after date change: {' '.join(body)[:1500]!r}")
    with open(os.path.join(OUTPUT_DIR, "tc_ANR_viewer_after_date.html"), "w", encoding="utf-8") as f:
        f.write(page.content())
    page.screenshot(path=os.path.join(OUTPUT_DIR, "tc_ANR_viewer_after_date.png"), full_page=True)
    page.close()


def tc_session_diag(context):
    """Why the TC viewer says "ASP.NET session has expired": log every
    Set-Cookie and each request the viewer makes (with its status), then
    the cookies the browser holds for the site."""
    log("\n==================== TC Energy session diagnostics ====================")
    page = context.new_page()
    def on_response(r):
        if "tceconnects" in r.url:
            sc = r.headers.get("set-cookie", "")
            log(f"  {r.status} {r.request.method} {r.url[:150]}" + (f"  SET-COOKIE: {sc[:300]!r}" if sc else ""))
    page.on("response", on_response)
    page.goto(TC_BASE + "TCeConnects.aspx?v=1.3&SID=67&info=Y&assetid=3005", wait_until="domcontentloaded", timeout=90000)
    page.wait_for_timeout(3000)
    log(f"  cookies after menu: {[(c['name'], c['domain'], c['path'], c.get('sameSite'), c.get('secure')) for c in context.cookies()]}")
    page.goto(f"{TC_BASE}ReportViewer.aspx?/InfoPost/OperationallyAvailableCapacityANR&pAssetNbr=3005",
              wait_until="domcontentloaded", timeout=180000)
    page.wait_for_timeout(20000)
    log(f"  cookies after viewer: {[(c['name'], c['domain'], c['path'], c.get('sameSite'), c.get('secure')) for c in context.cookies()]}")
    log(f"  viewer text: {' '.join(page.evaluate('() => document.body.innerText').split())[:400]!r}")
    page.close()


def transco_oac(context):
    """Transco's OA report is an SPA (1line.williams.com/#/oac-reports/
    transco). Capture its data calls, keep full JSON/CSV bodies, and grep
    them for LNG terminal delivery points (Sabine Pass via Gulf Trace,
    Cameron, Cove Point, Corpus Christi...)."""
    log("\n==================== Transco (Williams 1Line) OA ====================")
    page = context.new_page()
    bodies = []

    def on_response(resp):
        ct = resp.headers.get("content-type", "")
        if resp.request.resource_type in ("xhr", "fetch") or "json" in ct or "csv" in ct:
            try:
                bodies.append((resp.status, resp.request.method, resp.url, ct, resp.text(), resp.request.post_data))
            except Exception:
                pass
    context.on("response", on_response)  # popups too
    popups, downloads = [], []
    context.on("page", lambda pg: popups.append(pg))
    page.on("download", lambda d: downloads.append(d))
    lng = re.compile(r"SABINE|CHENIERE|GULF ?TRACE|LIQUEF|\bLNG\b|CAMERON|COVE POINT|CORPUS|PLAQUEMINES|CALCASIEU|VENTURE|GOLDEN PASS|ELBA", re.I)
    try:
        page.goto("https://www.1line.williams.com/#/oac-reports/transco", wait_until="networkidle", timeout=120000)
        page.wait_for_timeout(10000)
        dump_page(page, "transco_oac")
        n_before = len(bodies)
        page.get_by_text("View Report", exact=False).first.click(timeout=20000)
        page.wait_for_load_state("networkidle", timeout=180000)
        page.wait_for_timeout(15000)
        log(f"  View Report -> {len(bodies) - n_before} new data responses, {len(popups)} popups, {len(downloads)} downloads")
        for pg in popups:
            try:
                pg.wait_for_load_state("networkidle", timeout=120000)
            except Exception:
                pass
            log(f"  POPUP {pg.url}")
            text = pg.evaluate("() => document.body ? document.body.innerText : ''")
            log(f"    {' '.join(text.split())[:800]!r}")
            try:
                pg.screenshot(path=os.path.join(OUTPUT_DIR, "transco_popup.png"))
                with open(os.path.join(OUTPUT_DIR, "transco_popup.html"), "w", encoding="utf-8") as f:
                    f.write(pg.content())
            except Exception as e:
                log(f"    popup save failed: {e}")
        for d in downloads:
            path = os.path.join(OUTPUT_DIR, "transco_download_" + d.suggested_filename)
            d.save_as(path)
            log(f"  DOWNLOAD {d.url} -> {path}")
            with open(path, encoding="utf-8", errors="replace") as f:
                text = f.read()
            log(f"    {text[:400]!r}")
            for l in text.splitlines():
                if lng.search(l):
                    log(f"    LNG: {l[:300]}")
        page.screenshot(path=os.path.join(OUTPUT_DIR, "transco_oac_report.png"), full_page=False)
        # any download/export control
        for label in ("CSV", "Download", "Export", "Excel"):
            loc = page.get_by_text(label, exact=False)
            if loc.count():
                log(f"  control with text {label!r}: {loc.count()}")
    except Exception as e:
        log(f"  FAILED: {type(e).__name__}: {str(e)[:300]}")
    log(f"  {len(bodies)} data responses:")
    for i, (status, method, url, ct, body, post) in enumerate(bodies):
        log(f"    [{i}] {status} {method} {url[:200]} ({ct}, {len(body):,} chars)")
        if post:
            log(f"      POST {post[:400]!r}")
        log(f"      {body[:300]!r}")
        hits = [l for l in re.split(r"[\n}]", body) if lng.search(l)]
        for h in hits[:15]:
            log(f"      LNG: {h[:400]}")
        with open(os.path.join(OUTPUT_DIR, f"transco_body_{i}.txt"), "w", encoding="utf-8") as f:
            f.write(f"{method} {url}\n{post or ''}\n\n{body}")
    page.close()


def vg_gp_recon(context):
    """Whole-plant feed meters on the terminals' own interstate pipelines:
    Golden Pass Pipeline (gasnom.com, like Cameron Interstate), and Venture
    Global's TransCameron (Calcasieu Pass) and Gator Express (Plaquemines)
    on Quorum's myquorumcloud informational postings."""
    log("\n==================== Golden Pass Pipeline (gasnom) ====================")
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import lng_feedgas_daily as lfd
    from datetime import date
    page = context.new_page()
    for pipe in ("goldenpass", "GOLDENPASS"):
        for day in (date(2026, 9, 27), date(2026, 6, 1), date(2025, 12, 1)):
            try:
                df = lfd.fetch_gasnom(page, pipe, day)
                log(f"  {pipe} {day}: {len(df)} delivery points")
                for loc, r in df.sort_values("scheduled_dth", ascending=False).iterrows():
                    log(f"    {loc} {r['name'][:60]!r} {r['scheduled_dth']:,.0f}")
            except Exception as e:
                log(f"  {pipe} {day}: FAILED {type(e).__name__}: {str(e)[:200]}")
        if "df" in dir():
            break
    page.close()
    log("\n==================== Venture Global pipelines (Quorum) ====================")
    for tag, url in [
        ("transcameron_home", "https://web-prd.myquorumcloud.com/VGPPA1IPWS/?tspno=10"),
        ("transcameron_oa", "https://web-prd.myquorumcloud.com/VGL17IPWS/OpAvailPosting?tspno=10"),
        ("gatorexpress_home", "https://web-prd.myquorumcloud.com/VGPPB1IPWS/?tspno=2"),
        ("gatorexpress_oa", "https://web-prd.myquorumcloud.com/VGPPB1IPWS/OpAvailPosting?tspno=2"),
    ]:
        capture_page(context, tag, url)
        pg = context.new_page()
        try:
            pg.goto(url, wait_until="networkidle", timeout=90000)
            pg.wait_for_timeout(4000)
            links = pg.evaluate("() => [...document.querySelectorAll('a[href]')].map(a => [a.innerText.trim().slice(0, 60), a.href])")
            for t, h in links[:80]:
                log(f"    LINK {t!r} -> {h}")
            rows = pg.evaluate(lfd.TABLE_ROWS_JS)
            log(f"    {len(rows)} table rows; first: {rows[:3]}")
            for r in rows:
                if re.search(r"LNG|CALCAS|PLAQUE|TERMINAL|DELIVERY|Del", " ".join(r), re.I):
                    log(f"    ROW {r[:14]}")
        except Exception as e:
            log(f"  {tag} FAILED {type(e).__name__}: {str(e)[:200]}")
        pg.close()


def quorum_history(context):
    """Venture Global plant delivery meters (Quorum) and Golden Pass
    Pipeline's terminal meter (gasnom) across the gas days we need."""
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import lng_feedgas_daily as lfd
    from datetime import date
    page = context.new_page()
    page.goto("https://web-prd.myquorumcloud.com/VGPPB1IPWS/OpAvailPosting?tspno=2", wait_until="domcontentloaded", timeout=90000)
    days = [date(2026, 9, 27), date(2026, 9, 1), date(2026, 8, 1), date(2026, 7, 1), date(2026, 6, 1), date(2026, 1, 15), date(2025, 6, 1), date(2024, 6, 1), date(2023, 6, 1)]
    for pipe, locs in (("VGL17IPWS:10", ["VGCPD", "ANR", "TTC", "BRL"]), ("VGPPB1IPWS:2", ["VGPQD", "TGP", "TETCO", "CGT"])):
        for day in days:
            try:
                f = lfd.fetch_quorum(page, pipe, day)
                log(f"  {pipe} {day}: " + ", ".join(f"{l}={f.at[l, 'scheduled_dth']:,.0f}" for l in locs if l in f.index))
            except Exception as e:
                log(f"  {pipe} {day}: FAILED {type(e).__name__}: {str(e)[:200]}")
    for day in (date(2026, 9, 27), date(2026, 8, 1), date(2026, 7, 1)):
        try:
            f = lfd.fetch_gasnom(page, "goldenpass", day)
            log(f"  goldenpass {day}: Terminal={f.at['1097217', 'scheduled_dth']:,.0f}")
        except Exception as e:
            log(f"  goldenpass {day}: FAILED {type(e).__name__}: {str(e)[:200]}")
    page.close()


def gulfsouth_csv(context):
    """Click one Gulf South Operational Capacity CSV download (27 Sep
    Intraday 3), capture the request it makes and the file; list Stratton
    Ridge / Freeport rows."""
    log("\n==================== Gulf South OA CSV ====================")
    page = context.new_page()
    reqs = []
    page.on("request", lambda r: reqs.append((r.method, r.url, r.post_data)) if "bwpmlp" in r.url or "download" in r.url.lower() else None)
    try:
        page.goto("https://infopost.bwpipelines.com/", wait_until="networkidle", timeout=90000)
        page.wait_for_timeout(3000)
        page.evaluate("""() => { const t = document.querySelector('[data-cy=tsp-wrapper-div] button, [data-cy=tsp-wrapper-div] .dropdown-toggle'); if (t) t.click(); }""")
        page.wait_for_timeout(1000)
        page.evaluate("""() => document.querySelector('[data-cy^="078444247"]').click()""")
        page.wait_for_timeout(4000)
        page.goto("https://www.gasquest.com/capacity/operationally-available", wait_until="networkidle", timeout=90000)
        page.wait_for_timeout(6000)
        row = page.locator("tr", has_text="09/27/2026 Intraday 3").first
        icons = row.locator("[title], a, button, i, span")
        log(f"  row found: {row.count()}; clickable bits: {[ (icons.nth(i).get_attribute('title'), icons.nth(i).get_attribute('class')) for i in range(min(icons.count(), 12))]}")
        target = row.locator("[title*=csv i], [title*=CSV]").first
        if not target.count():
            target = row.locator("[title]").first
        n0 = len(reqs)
        with page.expect_download(timeout=60000) as info:
            target.click()
        d = info.value
        path = os.path.join(OUTPUT_DIR, "gulfsouth_" + d.suggested_filename)
        d.save_as(path)
        log(f"  DOWNLOAD {d.url} -> {path}")
        for r in reqs[n0:]:
            log(f"  REQ {r[0]} {r[1][:250]} {(r[2] or '')[:400]}")
        text = open(path, encoding="utf-8", errors="replace").read()
        lines = text.splitlines()
        log(f"  {len(lines)} lines; header: {lines[:3]}")
        for l in lines:
            if re.search(r"STRATTON|FREEPORT|LNG|BRAZORIA|ANGLETON", l, re.I):
                log(f"    ROW {l[:300]}")
    except Exception as e:
        log(f"  FAILED: {type(e).__name__}: {str(e)[:300]}")
        for r in reqs[-15:]:
            log(f"  REQ {r[0]} {r[1][:250]} {(r[2] or '')[:300]}")
    page.close()


def covepoint_csv(context):
    """Newest Cove Point OA CSV for yesterday's gas day (the listing page
    links each posting's CSV at /docs/cpl/postings/{id}/0/cpl-{id}.csv)."""
    log("\n==================== Cove Point CSV ====================")
    page = context.new_page()
    page.goto("https://infopost.bhegts.com/cpl/postings/capacity-operationally-available", wait_until="networkidle", timeout=90000)
    page.wait_for_timeout(4000)
    rows = page.evaluate("""() => [...document.querySelectorAll('[role=row]')].map(r => ({
        text: r.innerText.replace(/\\s+/g, ' ').trim(),
        csv: (r.querySelector('a[href$=".csv"]') || {}).href}))""")
    for r in rows[:12]:
        log(f"  {r}")
    target = next((r for r in rows if r.get("csv") and "09/27/2026 Intraday 3" in r["text"]), None) or \
        next((r for r in rows if r.get("csv")), None)
    if target:
        text = context.request.get(target["csv"], timeout=60000).text()
        with open(os.path.join(OUTPUT_DIR, "covepoint_oa.csv"), "w", encoding="utf-8") as f:
            f.write(text)
        log(f"\n  {target['csv']}:")
        for line in text.splitlines()[:60]:
            log(f"    {line[:250]}")
    page.close()


def main():
    wanted = sys.argv[1:] or ["cheniere", "tc", "bhe"]
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(ignore_https_errors=True, accept_downloads=True, user_agent=UA, viewport={"width": 1600, "height": 1200})
        if "cheniere" in wanted:
            cheniere(context)
        if "tc" in wanted:
            tc_csv(context)
        if "bhe" in wanted:
            bhe(context)
        if "tc2" in wanted:
            tc_params(context)
        if "covepoint" in wanted:
            capture_page(context, "covepoint_oac", "https://infopost.bhegts.com/cpl/postings/capacity-operationally-available")
        if "big" in wanted:
            # bu=BIG&Type=OA returned Enbridge's generic error page; try
            # the other URL shapes, and the LINK home page's own route.
            for i, url in enumerate([
                "https://rtba.enbridge.com/InformationalPosting/Default.aspx?bu=BIG",
                "https://rtba.enbridge.com/InformationalPosting/Default.aspx?bu=BIG&Type=OAC",
                "https://infopost.enbridge.com/infopost/BIGHome.asp?Pipe=BIG",
            ]):
                capture_page(context, f"big_{i}", url)
        if "gulfsouth" in wanted:
            capture_page(context, "gulfsouth_home", "https://infopost.bwpipelines.com/")
        if "enbridge_units" in wanted:
            enbridge_units(context)
        if "gulfsouth_ui" in wanted:
            gulfsouth_ui(context)
        if "tc_params" in wanted:
            tc_param_probe(context)
        if "tc_params2" in wanted:
            tc_param_probe2(context)
        if "tc_session" in wanted:
            tc_session_diag(context)
        if "transco" in wanted:
            transco_oac(context)
        if "vg_gp" in wanted:
            vg_gp_recon(context)
        if "quorum_hist" in wanted:
            quorum_history(context)
        if "gulfsouth_csv" in wanted:
            gulfsouth_csv(context)
        if "covepoint_csv" in wanted:
            covepoint_csv(context)
        browser.close()
    log(f"\nDONE. Outputs in {OUTPUT_DIR}/")


if __name__ == "__main__":
    main()
