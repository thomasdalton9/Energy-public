"""
Malaysia + Philippines natural gas discovery (run from GitHub Actions; the sandbox is blocked).
Wanted: production, consumption by sector (power, industry, feedstock, res/com, LNG plant fuel), LNG trade
(MY exports from Bintulu, imports at Melaka/Pengerang; PH imports at Batangas since 2023, Malampaya production).
Earlier rounds (GAS_SSEA_DISCOVERY1-3.py): Malaysia = DOSM quarterly only; Philippines = nothing monthly.

Candidates checked here:
  MY  data.gov.my catalogue (gas / LNG / trade datasets + API), MEIH (meih.st.gov.my) statistics pages,
      DOSM mining of petroleum & natural gas release, METS online (HS-level trade), Petronas IR
  PH  DOE energy-statistics pages (natural gas section), PSA OpenSTAT PXWeb API, PSA trade releases
  Both  UN Comtrade public preview API (HS 271111 monthly; MY exports 458, PH imports 608) as an official
        mirror of the national customs data
Round 1 functions: malaysia(), philippines(), comtrade(); round 2: meih2, dosm2, mets2, petronas2, comtrade2,
doe2, doe2_render; round 3: comtrade3, opendosm3, meih3, doe3, psa3; round 4: doe4, psa4, ipi4, bnm4; round 5: psa5, myes5; round 6: psa6 (run with function names as arguments; default = latest round).
Prints status, size, data-like links and text around gas keywords. Nothing is written.

FINDINGS (runs 1-6, Oct 2026, all from GitHub Actions; workflow discovery_archive/workflows/malaysia_philippines_gas_probe.yml)
RUN 1
  - UN Comtrade public preview API (comtradeapi.un.org/public/v1/preview/C/M/HS, no key) answers from Actions, but
    only ONE period per call ("Maximum number of periods for preview is 1") and ~1 call / 3 s (HTTP 429 otherwise).
    getDA (data availability): Malaysia (458) monthly HS 2010-01..2026-08; Philippines (608) 2010-01..2026-06.
  - data.gov.my API: trade_sitc_1d only (SITC section 3 = all mineral fuels, no LNG split); ipi / ipi_1d monthly.
  - MEIH (meih.st.gov.my/statistics): Liferay portlet, plain HTML, reachable. Natural gas menu = Reserves,
    Production, Import & Export, Final Consumption (by sector) + Gas Distribution Statistics (consumption, consumers).
  - DOSM 'Mining of Petroleum and Natural Gas Statistics' release page reachable (quarterly).
  - METS online (metsonline.dosm.gov.my, HS-level trade) is a React SPA; bundle has no usable API endpoint (run 2).
  - Petronas IR: quarterly/half-year 'Financial & Operational Report' PDFs, direct links on
    /investor-relations/financial-results.
  - DOE (doe.gov.ph) energy-statistics: only 'Electric Power Industry'; every natural-gas URL guessed renders the
    Next.js shell (no content). PSA web pages (psa.gov.ph) 403 from Actions; OpenSTAT PXWeb API answers.
RUN 2
  - MEIH tables are annual only, years 1978/1980 .. 2021 (latest option). Form POSTs to
    /STOASPublicPortlet/energystatistic/searchStatistic.oas; excel export = same URL + ?rptOutput=excel&<form>.
  - MyEnergyStats (myenergystats.st.gov.my) natural-gas page is narrative only; 'Energy Database' = GSO-style
    electricity dashboard (2025), no gas data.
  - DOSM Q2 2026 release PDF (2 pages, text extractable): gas production 763.5 bcf in Q2 2026 (+19.1% y/y),
    LNG exports RM16.1 bn, LNG imports RM1.6 bn - quarterly totals in text only, no monthly table.
  - Petronas FY2025 FOR PDF (16 pages, text extractable): group-level gross LNG sales 36.62 Mt (FY), cargo counts,
    sales gas 2,171 MMscfd Peninsular - annual/half-year group KPIs, not a national series.
  - Comtrade: one period per call works (rows per partner, netWgt in kg = LNG tonnes x1000).
  - DOE Payload CMS media API (d24qbtp4vooyzi.cloudfront.net/api/media?where[filename][like]=..) works: only 3
    'Natural Gas' files (auction notices, NGMD org chart). No gas statistics files.
RUN 3
  - Comtrade HS 271111 rows: partnerCode 0 = world total, netWgt (kg), primaryValue (USD), FOB/CIF. Several
    reporters / flows per call allowed (reporterCode=458,608 & flowCode=X,M) - still one period.
      MY exports Aug 2026 2.38 Mt (USD 1.50 bn; CN, JP, KR, TH, TW, ...), Jun 2026 2.78 Mt; Jan 2015 2.43 Mt;
      Jan 2010 has value but no weight. MY imports Aug 2026 0.127 Mt (AU), Jun 2026 0.246 Mt.
      PH imports Jun 2026 0.264 Mt (AU, CA, ID, OM); first LNG cargo Apr 2023 (66.2 kt, UAE), May 2023 nil,
      then ~1 cargo a month in 2023 (ID, MY, AU, OM, US), 2 in Dec 2023. PH reports no LNG exports.
  - OpenDOSM energy balance (storage.data.gov.my/energy/*.csv, Energy Commission data, ktoe, ANNUAL 1978/1980-2024):
    energy_supply_commodity (natural_gas primary supply), energy_consumption_commodity (natural_gas final
    consumption), energy_consumption_sector (all fuels by sector, no gas split).
  - IPI 5-digit (storage.dosm.gov.my/ipi/ipi_5d.csv): item 06201 'natural gas' monthly volume INDEX (2015=100),
    2015-01..2026-08 (built from Petronas monthly production data); 06101 crude oil.
  - DOE media search for supply/demand/statistics/energy plan/Malampaya files: nothing gas-statistical.
  - PSA OpenSTAT tree: 3G/ELE/ES Table 14.1a TPES by source (natural gas indigenous) and 14.1b TFEC by sector,
    annual 2014-2022/24 (kTOE); 2L/IMT monthly CIF value (CIF) and gross kilo (GKI) of imports by commodity and
    country, 1991..2026 (2026 table updated 2026-09-30, Jan-Aug).
RUN 4
  - DOE Payload pages API (/api/pages?where[title][like]=..) works: no natural gas statistics page exists on the
    new site (only power statistics, NGMD division page, press releases e.g. Malampaya maintenance).
  - BNM (bnm.gov.my) 403 from Actions.
RUN 5-6
  - PSA OpenSTAT GKI/CIF tables: the 'Commodity Code' variable lists 0 values in the metadata and item queries for
    27111100 / 2711110000 / 271111 return nothing (11-digit PSCC codes, hidden dimension) - not usable via API.
  - st.gov.my publications pages render empty (3 links) - no Energy Commission handbook/NEB file found.

CONCLUSIONS
  MY production   monthly: DOSM IPI item 06201 index (OpenDOSM CSV, 2015-01.., ~6 weeks lag) - index only; scale
                  to DOSM quarterly bcf (release text) or to OpenDOSM annual ktoe. Annual: OpenDOSM energy balance.
  MY consumption  by sector: annual only (MEIH to 2021 with sectors; OpenDOSM to 2024 totals). Nothing monthly.
  MY LNG exp/imp  UN Comtrade HS 271111 (Malaysian customs/DOSM data) monthly 2010-.. (tonnes from ~2011),
                  current to 2026-08, by partner.
  PH LNG imports  UN Comtrade HS 271111 reporter 608 monthly, Apr 2023 .., current to 2026-06, by origin.
  PH production / consumption (Malampaya): no DOE gas statistics on the current site; PSA annual TPES (natural
                  gas indigenous) via OpenSTAT; gas-for-power is already covered by IEMOP plant dispatch.
  Fallback for both: JODI Gas (jodidb.org Beyond 20/20, see discovery_archive/rest_of_world/JODI_LNG_LAG_PULL4.py).
"""
import io
import json
import re
import sys
from urllib.parse import urljoin

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "en"}
T = (15, 90)
LINK = re.compile(r"""(?:href|src|data-src|data-url)=["']([^"']+)["']""", re.I)
S = requests.Session()
S.headers.update(H)


def out(s=""):
    print(s, flush=True)


def get(u, quiet=False, **kw):
    try:
        r = S.get(u, timeout=T, verify=False, **kw)
        if not quiet:
            out(f"\n### {u}\n  -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}b final={r.url} "
                f"LM={r.headers.get('last-modified')}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"\n### {u}\n  ERROR {type(e).__name__}: {str(e)[:200]}")
        return None


def plain(html):
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", html, flags=re.S | re.I)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))


def links(r, pat, n=60):
    ls = []
    for l in LINK.findall(r.text):
        if re.search(pat, l, re.I):
            a = urljoin(r.url, l.replace("&amp;", "&"))
            if a not in ls:
                ls.append(a)
    out(f"  links /{pat}/ ({len(ls)}):")
    for l in ls[:n]:
        out(f"    {l}")
    return ls


def ctx(text, pat, n=5, w=250):
    for i, m in enumerate(re.finditer(pat, text, re.I)):
        if i >= n:
            break
        out(f"  CTX: ...{text[max(0, m.start() - w):m.end() + w]}...")


def pdf_text(content, pat, maxp=4, chars=3000):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        out(f"  pdf pages: {len(pdf.pages)}")
        hits = 0
        for i, p in enumerate(pdf.pages):
            t = p.extract_text() or ""
            if pat is None or re.search(pat, t, re.I):
                out(f"  ---- page {i + 1} ----\n{t[:chars]}")
                hits += 1
                if hits >= maxp:
                    break


GAS = r"natural gas|\bLNG\b|liquefied natural"
DATA = r"\.(xlsx?|csv|pdf|zip|json|parquet)(\?|$)|download|api|statistic|dataset|data-catalogue"


# ------------------------------------------------------------------ Malaysia
def malaysia():
    out("\n================ MALAYSIA ================")
    r = get("https://api.data.gov.my/data-catalogue?meta=true")
    if r is not None and r.ok:
        try:
            cat = r.json()
            out(f"  catalogue entries: {len(cat)}; sample keys {list(cat[0].keys()) if cat else None}")
            for c in cat:
                s = json.dumps(c).lower()
                if re.search(r"gas|lng|petrol|crude|energy|mining|trade_|export|import|hs", s):
                    out(f"    {json.dumps(c)[:300]}")
        except Exception as e:  # noqa: BLE001
            out(f"  not json: {e}; {r.text[:300]}")
    for ds in ("trade_hs2", "trade_hs", "trade_sitc_1d", "trade_sitc_2d", "ipi_1d", "ipi_2d", "ipi_5d", "ipi"):
        r = get(f"https://api.data.gov.my/data-catalogue?id={ds}&limit=3&sort=-date")
        if r is not None:
            out(f"  {r.text[:400]}")
    for u in ("https://meih.st.gov.my/", "https://meih.st.gov.my/statistics", "https://meih.st.gov.my/database",
              "https://meih.st.gov.my/publications", "https://myenergystats.st.gov.my/",
              "https://www.st.gov.my/en/web/general/details/publications"):
        r = get(u)
        if r is not None and r.ok:
            t = plain(r.text)
            out(f"  title: {re.search(r'<title>(.*?)</title>', r.text, re.S).group(1).strip()[:100] if '<title>' in r.text else ''}")
            links(r, r"statistic|gas|energy-balance|download|\.xlsx?|\.pdf|portlet|p_p_id")
            ctx(t, GAS, n=4)
    r = get("https://www.dosm.gov.my/portal-main/release-content/mining-of-petroleum-and-natural-gas-statistics-q22026")
    if r is not None and r.ok:
        t = plain(r.text)
        ctx(t, r"natural gas|LNG", n=6)
        links(r, r"download|\.pdf|\.xlsx")
    for u in ("https://metsonline.dosm.gov.my/", "https://metsonline.dosm.gov.my/tradev2/product-coderesult",
              "https://www.matrade.gov.my/en/malaysian-exporters/services-for-exporters/trade-market-information/trade-statistics"):
        r = get(u)
        if r is not None and r.ok:
            links(r, r"hs|search|api|product|download|\.xlsx|\.csv", n=40)
            ctx(plain(r.text), r"HS|harmoni", n=3, w=150)
    for u in ("https://www.petronas.com/investor-relations", "https://www.petronas.com/investor-relations/financial-results",
              "https://www.petronas.com/investor-relations/quarterly-reports"):
        r = get(u)
        if r is not None and r.ok:
            links(r, r"\.pdf|result|quarter|activity|report", n=40)


# ------------------------------------------------------------------ Philippines
def philippines():
    out("\n================ PHILIPPINES ================")
    seen = set()
    for u in ("https://doe.gov.ph/data-and-prices/energy-statistics",
              "https://doe.gov.ph/data-and-prices/energy-statistics/natural-gas",
              "https://doe.gov.ph/data-and-prices/energy-statistics/natural-gas-industry",
              "https://doe.gov.ph/data-and-prices/energy-statistics/downstream-natural-gas",
              "https://doe.gov.ph/data-and-prices/energy-statistics/energy-supply-and-demand",
              "https://doe.gov.ph/data-and-prices/energy-statistics/key-energy-statistics",
              "https://doe.gov.ph/data-and-prices/energy-statistics/philippine-energy-situationer",
              "https://doe.gov.ph/data-and-prices/energy-statistics/energy-balance-table",
              "https://doe.gov.ph/sitemap.xml"):
        r = get(u)
        if r is None or not r.ok:
            continue
        ls = links(r, r"energy-statistics/|gas|situationer|balance|\.pdf|\.xlsx?", n=80)
        seen.update(l for l in ls if "energy-statistics/" in l)
        ctx(plain(r.text), GAS, n=3, w=200)
    for u in sorted(seen)[:25]:
        if re.search(r"gas|situation|balance|supply|key-energy", u, re.I):
            r = get(u)
            if r is not None and r.ok:
                ls = links(r, r"\.pdf|\.xlsx?|\.csv", n=40)
                ctx(plain(r.text), GAS, n=2, w=200)

    r = get("https://openstat.psa.gov.ph/PXWeb/api/v1/en/DB")
    if r is not None and r.ok:
        out(f"  {r.text[:1500]}")
        try:
            for db in r.json():
                rr = get(f"https://openstat.psa.gov.ph/PXWeb/api/v1/en/DB/{db['dbid']}", quiet=True)
                if rr is not None and rr.ok:
                    for it in rr.json():
                        if re.search(r"trade|energy|import|export|commod", it.get("text", ""), re.I):
                            out(f"    {db['dbid']}/{it.get('id')}: {it.get('text')}")
        except Exception as e:  # noqa: BLE001
            out(f"  parse: {e}")
    for q in ("natural gas", "liquefied", "petroleum gas"):
        r = get(f"https://openstat.psa.gov.ph/PXWeb/api/v1/en/DB?query={q}")
        if r is not None:
            out(f"  {r.text[:1200]}")
    for u in ("https://psa.gov.ph/statistics/export-import", "https://psa.gov.ph/statistics/export-import/monthly",
              "https://psa.gov.ph/statistics/foreign-trade"):
        r = get(u)
        if r is not None and r.ok:
            links(r, r"\.xlsx?|\.pdf|import|trade", n=40)


# ------------------------------------------------------------------ UN Comtrade (both)
def comtrade():
    out("\n================ UN COMTRADE preview (HS 271111) ================")
    base = "https://comtradeapi.un.org/public/v1/preview/C/M/HS"
    for rep, flow in (("458", "X"), ("458", "M"), ("608", "M")):
        for per in ("202601,202602,202603,202604,202605,202606,202607,202608", "202301,202401,202501", "201001,201501"):
            r = get(f"{base}?reporterCode={rep}&period={per}&cmdCode=271111&flowCode={flow}&partnerCode=0")
            if r is None or not r.ok:
                if r is not None:
                    out(f"  {r.text[:300]}")
                continue
            try:
                d = r.json().get("data", [])
                out(f"  rows {len(d)}")
                for x in d[:10]:
                    out(f"    {x.get('period')} {x.get('reporterDesc')} {x.get('flowCode')} {x.get('partnerDesc')} "
                        f"kg={x.get('netWgt')} qty={x.get('qty')} {x.get('qtyUnitAbbr')} USD={x.get('primaryValue')}")
            except Exception as e:  # noqa: BLE001
                out(f"  {e}: {r.text[:300]}")
    # data availability: latest month each reporter has submitted
    for rep in ("458", "608"):
        r = get(f"https://comtradeapi.un.org/public/v1/getDA/C/M/HS?reporterCode={rep}")
        if r is not None:
            out(f"  {r.text[:800]}")



# ================================================================== ROUND 2
def anchors(r):
    """(text, absolute url) for every <a> on a page."""
    res = []
    for m in re.finditer(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", r.text, re.S | re.I):
        t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", m.group(2))).strip()
        res.append((t, urljoin(r.url, m.group(1).replace("&amp;", "&"))))
    return res


def tables(html, n=3, rows=14):
    import pandas as pd
    try:
        ts = pd.read_html(io.StringIO(html))
    except Exception as e:  # noqa: BLE001
        out(f"  no tables ({type(e).__name__})")
        return
    out(f"  {len(ts)} html tables")
    for t in ts[:n]:
        out(f"  table {t.shape}:\n{t.head(rows).to_string()[:2500]}")


def meih2():
    out("\n================ MEIH statistics menu (round 2) ================")
    r = get("https://meih.st.gov.my/statistics")
    if r is None or not r.ok:
        return
    stat = [(t, u) for t, u in anchors(r) if "Eng_Statistic_WAR" in u]
    for t, u in stat:
        out(f"  [{t}] EVT=" + u.split("_eventId=")[-1])
    gas = [(t, u) for t, u in stat if re.search(r"categoryId=(2|6|8)&", u)]
    for t, u in gas[:12]:
        rr = get(u)
        if rr is None or not rr.ok:
            continue
        out(f"  page [{t}]")
        txt = plain(rr.text)
        ctx(txt, r"natural gas|year|unit|ktoe|mmscf", n=2, w=300)
        tables(rr.text, n=2)
        for at, au in anchors(rr):
            if re.search(r"export|excel|xls|csv|download|print", at + au, re.I):
                out(f"    dl? [{at}] {au[:250]}")
        for f in re.findall(r"<form[^>]*>", rr.text)[:5]:
            out(f"    form: {f[:300]}")
        for sel in re.findall(r"<select[^>]*name=[\"']([^\"']+)[\"'][^>]*>(.*?)</select>", rr.text, re.S)[:6]:
            opts = re.findall(r"<option[^>]*value=[\"']([^\"']*)[\"']", sel[1])
            out(f"    select {sel[0]}: {opts[:40]}")
    r = get("https://myenergystats.st.gov.my/natural-gas1")
    if r is not None and r.ok:
        for t, u in anchors(r):
            if re.search(r"gas|lng|database|download|\.xlsx|\.pdf|\.csv", t + u, re.I):
                out(f"    [{t[:80]}] {u[:200]}")
        ctx(plain(r.text), r"LNG|mmscf|production|consumption", n=6, w=200)
    r = get("https://myenergystats.st.gov.my/energy-database")
    if r is not None and r.ok:
        for t, u in anchors(r):
            if re.search(r"gas|database|download|\.xlsx|\.pdf|\.csv|balance", t + u, re.I):
                out(f"    [{t[:80]}] {u[:200]}")


def dosm2():
    out("\n================ DOSM petroleum & gas release PDF / OpenDOSM catalogue ================")
    r = get("https://www.dosm.gov.my/site/downloadrelease?id=mining-of-petroleum-and-natural-gas-statistics-q22026&lang=English&admin_view=")
    if r is not None and r.ok and r.content[:4] == b"%PDF":
        pdf_text(r.content, r"natural gas|LNG", maxp=3, chars=2500)
    for u in ("https://open.dosm.gov.my/data-catalogue", "https://data.gov.my/data-catalogue"):
        r = get(u)
        if r is None or not r.ok:
            continue
        ids = sorted(set(re.findall(r'"id"\s*:\s*"([a-z0-9_]+)"', r.text)) | set(re.findall(r"data-catalogue/([a-z0-9_]+)", r.text)))
        out(f"  {len(ids)} ids; energy/trade-like: {[i for i in ids if re.search(r'gas|lng|petrol|fuel|energy|trade|mining|ipi|electric', i)]}")
    for ds in ("trade_sitc_2d", "trade_sitc_3d", "trade_sitc_5d", "trade_commodity", "ipi_5d", "ipi_mining",
               "mining", "fuel_production", "natural_gas", "lng", "trade_sitc_1d_country", "trade_hs_2d"):
        r = get(f"https://api.data.gov.my/data-catalogue?id={ds}&limit=2", quiet=True)
        out(f"  id {ds}: {r.status_code if r is not None else None} {(r.text[:200] if r is not None and r.ok else '')}")


def mets2():
    out("\n================ METS online bundle (API endpoints?) ================")
    r = get("https://metsonline.dosm.gov.my/")
    if r is None:
        return
    out(r.text[:1700])
    for src in re.findall(r"""src=["']([^"']+\.js)["']""", r.text)[:4]:
        js = get(urljoin(r.url, src))
        if js is None or not js.ok:
            continue
        eps = sorted(set(re.findall(r"""["'`](https?://[^"'`\s]+|/api/[^"'`\s]+|/tradev2/[^"'`\s]+)["'`]""", js.text)))
        out(f"  endpoints in {src} ({len(eps)}): {eps[:80]}")


def petronas2():
    out("\n================ Petronas Financial & Operational Report ================")
    u = ("https://www.petronas.com/sites/default/files/uploads/content/2026/"
         "PETRONAS%20Group%20FRA%20FY2025%20-%20Financial%20Operational%20Report.pdf")
    r = get(u)
    if r is not None and r.ok and r.content[:4] == b"%PDF":
        pdf_text(r.content, r"LNG|natural gas|sales volume|production", maxp=4, chars=2500)


def comtrade2():
    import time
    out("\n================ UN Comtrade preview, one period per call ================")
    base = "https://comtradeapi.un.org/public/v1/preview/C/M/HS"
    for rep in ("458", "608"):
        r = get(f"https://comtradeapi.un.org/public/v1/getDA/C/M/HS?reporterCode={rep}", quiet=True)
        if r is not None and r.ok:
            per = sorted(d["period"] for d in r.json().get("data", []))
            out(f"  reporter {rep} monthly HS periods: {len(per)} {per[:2]}..{per[-6:]}")
        time.sleep(4)
    for rep, flow, per in (("458", "X", "202606"), ("458", "M", "202606"), ("608", "M", "202606"),
                           ("608", "M", "202305"), ("458", "X", "201001")):
        time.sleep(4)
        r = get(f"{base}?reporterCode={rep}&period={per}&cmdCode=271111&flowCode={flow}")
        if r is None:
            continue
        if not r.ok:
            out(f"  {r.text[:200]}")
            continue
        d = r.json().get("data", [])
        out(f"  rows {len(d)}")
        for x in sorted(d, key=lambda x: -(x.get("primaryValue") or 0))[:8]:
            out(f"    {x.get('period')} {x.get('reporterDesc')} {x.get('flowCode')} {x.get('partnerDesc')} "
                f"kg={x.get('netWgt')} qty={x.get('qty')} {x.get('qtyUnitAbbr')} USD={x.get('primaryValue')}")


def doe2():
    out("\n================ DOE Payload CMS media / pages search ================")
    import urllib.parse
    host = "https://d24qbtp4vooyzi.cloudfront.net"
    found = {}
    for term in ("Natural Gas", "natgas", "Malampaya", "LNG", "Situationer", "Energy Balance", "Supply and Demand",
                 "Key Energy", "Gas Production", "Gas Consumption"):
        for h in (host, "https://doe.gov.ph"):
            q = urllib.parse.urlencode({"limit": 100, "depth": 0, "where[filename][like]": term, "sort": "-updatedAt"})
            r = get(f"{h}/api/media?{q}", quiet=True)
            ok = r is not None and r.ok and "json" in (r.headers.get("content-type") or "")
            out(f"  media like '{term}' @ {h}: {r.status_code if r is not None else None} {'json' if ok else ''}")
            if ok:
                j = r.json()
                out(f"    totalDocs {j.get('totalDocs')}")
                for d in j.get("docs") or []:
                    found[d.get("filename")] = (d.get("updatedAt") or "")[:10]
                break
    for fn, t in sorted(found.items(), key=lambda x: x[1], reverse=True)[:120]:
        out(f"    MEDIA {t} {fn}")
    for coll in ("pages", "posts", "articles", "statistics", "energy-statistics"):
        q = urllib.parse.urlencode({"limit": 10, "depth": 0, "where[slug][like]": "gas"})
        r = get(f"https://doe.gov.ph/api/{coll}?{q}", quiet=True)
        out(f"  /api/{coll}?slug~gas: {r.status_code if r is not None else None} "
            f"{(r.text[:600] if r is not None and r.ok and 'json' in (r.headers.get('content-type') or '') else '')}")


def doe2_render():
    out("\n================ DOE energy-statistics rendered (Playwright) ================")
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(user_agent=H["User-Agent"])
        for u in ("https://doe.gov.ph/data-and-prices/energy-statistics",
                  "https://doe.gov.ph/data-and-prices"):
            try:
                pg.goto(u, timeout=60000, wait_until="networkidle")
            except Exception as e:  # noqa: BLE001
                out(f"  {u}: {e}")
                continue
            hrefs = pg.eval_on_selector_all("a", "els => els.map(e => [e.innerText.trim(), e.href])")
            out(f"\n### rendered {u}: {len(hrefs)} links")
            for t, h in hrefs:
                if re.search(r"statistic|gas|situationer|balance|supply|\.pdf|\.xlsx", t + h, re.I):
                    out(f"    [{t[:70]}] {h[:200]}")
        b.close()



# ================================================================== ROUND 3
def comtrade3():
    import time
    out("\n================ UN Comtrade preview: fields, partners, PH start, multi-flow calls ================")
    base = "https://comtradeapi.un.org/public/v1/preview/C/M/HS"
    keys = ("period", "reporterCode", "flowCode", "partnerCode", "partner2Code", "motCode", "customsCode", "cmdCode",
            "netWgt", "qty", "qtyUnitCode", "altQty", "primaryValue", "cifvalue", "fobvalue", "isAggregate")
    calls = [("458", "X", "202608", "271111"), ("458", "X", "202606", "271111"), ("458", "M", "202608", "271111"),
             ("608", "M", "202606", "271111"), ("458", "X", "202001", "271111"), ("458", "X", "201501", "271111")]
    calls += [("608", "M", f"2023{m:02d}", "271111") for m in range(3, 13)]
    calls += [("608", "M", "202306", "2711"), ("608", "M", "202306", "271121"), ("458", "X,M", "202606", "271111"),
              ("458,608", "M", "202606", "271111"), ("608", "X", "202606", "271111")]
    for rep, flow, per, cmd in calls:
        time.sleep(3.5)
        r = get(f"{base}?reporterCode={rep}&period={per}&cmdCode={cmd}&flowCode={flow}", quiet=True)
        if r is None:
            continue
        if not r.ok:
            out(f"  {rep} {flow} {per} {cmd}: HTTP {r.status_code} {r.text[:150]}")
            continue
        d = r.json().get("data", [])
        tot = [x for x in d if x.get("partnerCode") == 0]
        out(f"  {rep} {flow} {per} {cmd}: rows {len(d)}; partners {sorted(x.get('partnerCode') for x in d)}; "
            f"mot {sorted(set(x.get('motCode') for x in d))}; total kg {[x.get('netWgt') for x in tot]} "
            f"USD {[x.get('primaryValue') for x in tot]}")
        if per in ("202608", "202606") and cmd == "271111" and "," not in rep + flow:
            for x in d[:12]:
                out("    " + " ".join(f"{k}={x.get(k)}" for k in keys))
    r = get("https://comtradeapi.un.org/files/v1/app/reference/partnerAreas.json", quiet=True)
    if r is not None and r.ok:
        out(f"  partnerAreas.json ok ({len(r.content)}b)")


def opendosm3():
    out("\n================ OpenDOSM energy datasets ================")
    import pandas as pd
    for ds in ("energy_supply_commodity", "energy_consumption_sector", "energy_consumption_commodity", "ipi_5d",
               "ipi_3d", "trade_sitc_1d"):
        r = get(f"https://api.data.gov.my/data-catalogue?id={ds}&limit=5&sort=-date", quiet=True)
        out(f"\n  api {ds}: {r.status_code if r is not None else None} {(r.text[:500] if r is not None else '')}")
        pg = get(f"https://open.dosm.gov.my/data-catalogue/{ds}", quiet=True)
        if pg is None or not pg.ok:
            continue
        files = sorted(set(re.findall(r"https://storage\.(?:dosm|data)\.gov\.my/[^\"'\\\s]+\.(?:csv|parquet)", pg.text)))
        out(f"  files: {files}")
        ctx(plain(pg.text), r"frequency|data as of|next update|monthly|annual", n=3, w=150)
        csvs = [f for f in files if f.endswith(".csv")] or files
        if csvs:
            try:
                f = csvs[0]
                df = pd.read_parquet(f) if f.endswith(".parquet") else pd.read_csv(f)
                out(f"  {f}: shape {df.shape}; cols {list(df.columns)}")
                if "date" in df:
                    out(f"  dates {df['date'].min()}..{df['date'].max()}")
                for c in df.columns:
                    if df[c].dtype == object and c != "date":
                        u = df[c].unique()
                        out(f"  {c} values ({len(u)}): {list(u)[:40]}")
                if ds in ("ipi_5d", "ipi_3d"):
                    col = [c for c in df.columns if df[c].dtype == object and c != "date"]
                    out(df[df[col[0]].astype(str).str.startswith(("06", "062", "0620", "B"))].tail(8).to_string())
                elif ds != "trade_sitc_1d":
                    gas = df[df.astype(str).apply(lambda x: x.str.contains("gas|lng", case=False)).any(axis=1)]
                    out(gas.tail(12).to_string()[:2500])
            except Exception as e:  # noqa: BLE001
                out(f"  read failed: {type(e).__name__}: {e}")


def meih3():
    out("\n================ MEIH: POST a natural gas table; MyEnergyStats dashboard ================")
    r = get("https://meih.st.gov.my/statistics", quiet=True)
    if r is None or not r.ok:
        return
    stat = [(t, u) for t, u in anchors(r) if "Eng_Statistic_WAR" in u and re.search(r"categoryId=(2|9)&", u)]
    for t, u in stat[:6]:
        rr = get(u, quiet=True)
        if rr is None or not rr.ok:
            continue
        m = re.search(r'<form method="POST" id="parameterForm"(.*?)</form>', rr.text, re.S)
        acts = re.findall(r"""(?:action|url)\s*[:=]\s*["']([^"']+)["']""", rr.text)
        out(f"\n  [{t}] {u.split('_eventId=')[-1]} form len {len(m.group(1)) if m else None}; actions {acts[:6]}")
        if m:
            names = re.findall(r'name="([^"]+)"', m.group(1))
            out(f"    fields {names}")
            opts = {n: re.findall(r'<option[^>]*value="([^"]*)"', blk)
                    for n, blk in re.findall(r'<select[^>]*name="([^"]+)"[^>]*>(.*?)</select>', m.group(1), re.S)}
            out(f"    last year option {[(k, v[-1] if v else None) for k, v in opts.items()]}")
            for s_ in re.findall(r"<script[^>]*>(.*?)</script>", rr.text, re.S):
                if "parameterForm" in s_ or "submit" in s_:
                    out("    script: " + " ".join(s_.split())[:900])
                    break
            data = {k: (v[-1] if k.lower().endswith("to") else (v[1] if len(v) > 1 else (v[0] if v else "")))
                    for k, v in opts.items()}
            for n in names:
                data.setdefault(n, "")
            for k in data:
                if k.lower().endswith("from") and opts.get(k):
                    data[k] = opts[k][-10] if len(opts[k]) > 10 else opts[k][0]
            pr = S.post(rr.url, data=data, timeout=T, verify=False)
            out(f"    POST {data} -> {pr.status_code} {len(pr.content)}b")
            tables(pr.text, n=2, rows=25)
            for at, au in anchors(pr):
                if re.search(r"excel|xls|csv|export|download", at + au, re.I):
                    out(f"    dl [{at}] {au[:250]}")
    out("\n  MyEnergyStats dashboard (rendered)")
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch()
            pg = b.new_page(user_agent=H["User-Agent"])
            reqs = []
            pg.on("request", lambda q: reqs.append(q.url))
            pg.goto("https://myenergystats.st.gov.my/en/dashboard", timeout=90000, wait_until="networkidle")
            out(f"    title {pg.title()}; frames {[f.url[:150] for f in pg.frames]}")
            for t_, h_ in pg.eval_on_selector_all("a", "els => els.map(e => [e.innerText.trim(), e.href])"):
                if re.search(r"gas|download|\.xlsx|\.csv|\.pdf|balance|database", t_ + h_, re.I):
                    out(f"    [{t_[:70]}] {h_[:200]}")
            out(f"    text: {pg.inner_text('body')[:1500]}")
            out(f"    data-like requests: {[q for q in reqs if re.search(r'api|json|powerbi|tableau|xlsx|csv', q, re.I)][:30]}")
            b.close()
    except Exception as e:  # noqa: BLE001
        out(f"    playwright: {type(e).__name__}: {e}")


def doe3():
    out("\n================ DOE Payload media: wider search ================")
    import urllib.parse
    host = "https://d24qbtp4vooyzi.cloudfront.net"
    found = {}
    for term in ("Gas", "Statistic", "Supply", "Demand", "Energy Plan", "PEP", "Upstream", "Petroleum", "Indigenous",
                 "Production", "Consumption", "Fact", "Highlights", "Energy Sector", "Accomplishment", "Import"):
        q = urllib.parse.urlencode({"limit": 300, "depth": 0, "where[filename][like]": term, "sort": "-updatedAt"})
        r = get(f"{host}/api/media?{q}", quiet=True)
        if r is not None and r.ok and "json" in (r.headers.get("content-type") or ""):
            j = r.json()
            out(f"  '{term}': totalDocs {j.get('totalDocs')}")
            for d in j.get("docs") or []:
                found[d.get("filename")] = ((d.get("updatedAt") or "")[:10], d.get("url") or "")
    keep = {k: v for k, v in found.items()
            if re.search(r"gas|lng|malampaya|supply|demand|statistic|energy plan|situationer|balance|indigenous|"
                         r"upstream|production|consumption|import|highlight|fact", k or "", re.I)
            and not re.search(r"bid|auction|advisory|invitation|award|notice|circular|order|memorandum|tor\b", k or "", re.I)}
    out(f"  {len(found)} files, {len(keep)} kept")
    for fn, (t, u) in sorted(keep.items(), key=lambda x: x[1][0], reverse=True)[:200]:
        out(f"    {t} {fn}")
    for coll in ("pages", "posts", "articles"):
        for fld in ("title", "slug"):
            q = urllib.parse.urlencode({"limit": 20, "depth": 0, f"where[{fld}][like]": "gas"})
            r = get(f"{host}/api/{coll}?{q}", quiet=True)
            ct = (r.headers.get("content-type") or "") if r is not None else ""
            out(f"  {host}/api/{coll} {fld}~gas: {r.status_code if r is not None else None} {ct} "
                f"{(r.text[:400] if r is not None else '')}")


def psa3():
    out("\n================ PSA OpenSTAT trade tables ================")
    r = get("https://openstat.psa.gov.ph/PXWeb/api/v1/en/DB", quiet=True)
    if r is None or not r.ok:
        return
    for db in r.json():
        out(f"  DB {db.get('id')}: {db.get('text')}")
    for q in ("imports", "foreign trade", "commodity imports", "mineral fuels", "fuel"):
        r = get(f"https://openstat.psa.gov.ph/PXWeb/api/v1/en/DB?query={q}", quiet=True)
        if r is not None and r.ok:
            hits = r.json()
            out(f"  query '{q}': {len(hits)} hits")
            for h in hits[:15]:
                out(f"    {h.get('path')}/{h.get('id')}: {h.get('title')[:140]} ({h.get('published')})")
    for u in ("https://psa.gov.ph/statistics/export-import", "https://openstat.psa.gov.ph/"):
        r = get(u)



# ================================================================== ROUND 4
def doe4():
    out("\n================ DOE Payload pages by title; gas pages' files ================")
    import urllib.parse
    host = "https://d24qbtp4vooyzi.cloudfront.net"
    r = get(f"{host}/api/media?limit=1", quiet=True)
    if r is not None and r.ok:
        out(f"  media total: {r.json().get('totalDocs')}")
    pages = {}
    for term in ("Natural Gas", "Statistic", "Situationer", "Energy Plan", "Malampaya", "Supply", "Upstream",
                 "Compendium", "Energy Balance", "LNG", "Gas", "Petroleum", "Downstream", "Data"):
        q = urllib.parse.urlencode({"limit": 100, "depth": 0, "where[title][like]": term})
        r = get(f"{host}/api/pages?{q}", quiet=True)
        if r is None or not r.ok:
            out(f"  pages ~{term}: {r.status_code if r is not None else None}")
            continue
        j = r.json()
        out(f"  pages ~{term}: {j.get('totalDocs')}")
        for d in j.get("docs") or []:
            pages[d.get("path")] = ((d.get("updatedAt") or "")[:10], d.get("title"))
    for pth, (t, ti) in sorted(pages.items(), key=lambda x: x[1][0], reverse=True)[:120]:
        out(f"    {t} {pth} | {ti}")
    gas = [p_ for p_, (t, ti) in pages.items()
           if p_ and re.search(r"gas|statistic|situationer|upstream|compendium|balance|malampaya|lng", (ti or "") + p_, re.I)
           and not re.search(r"bid|auction|procure|contingency|invitation", (ti or "") + p_, re.I)]
    for p_ in gas[:15]:
        rr = get("https://doe.gov.ph" + (p_ if p_.startswith("/") else "/" + p_))
        if rr is not None and rr.ok:
            fl = sorted(set(re.findall(r'https://[^"\'\s<>\\]+?\.(?:pdf|xlsx?|csv)', rr.text)))
            out(f"    files ({len(fl)}): {[urllib.parse.unquote(f.rsplit('/', 1)[-1]) for f in fl][:40]}")
            ctx(plain(rr.text), r"natural gas|LNG|Malampaya", n=2, w=200)


def psa4():
    out("\n================ PSA OpenSTAT Energy (3G) and trade (2L) trees ================")
    def walk(path, depth=0):
        r = get(f"https://openstat.psa.gov.ph/PXWeb/api/v1/en/{path}", quiet=True)
        if r is None or not r.ok:
            out(f"  {path}: {r.status_code if r is not None else None}")
            return
        for it in r.json():
            out(f"  {'  ' * depth}{path}/{it.get('id')} [{it.get('type')}] {it.get('text')[:120]} {it.get('updated', '')}")
            if it.get("type") == "l" and depth < 2:
                walk(f"{path}/{it.get('id')}", depth + 1)
    walk("DB/3G")
    walk("DB/2L")


def ipi4():
    out("\n================ DOSM IPI items: natural gas ================")
    import pandas as pd
    for f, col in (("https://storage.dosm.gov.my/ipi/ipi_5d.csv", "item"), ("https://storage.dosm.gov.my/ipi/ipi_3d.csv", "group")):
        df = pd.read_csv(f, dtype={col: str})
        out(f"  {f}: series {df['series'].unique()}; {col}s {sorted(df[col].unique())[:60]}")
        g = df[df[col].str.startswith("06") & (df["series"] == "abs")]
        for k, d in g.groupby(col):
            out(f"   {k}: {d['date'].min()}..{d['date'].max()}; last {d.tail(4)[['date', 'index']].values.tolist()}")
    pg = get("https://open.dosm.gov.my/data-catalogue/ipi_5d", quiet=True)
    if pg is not None:
        ctx(plain(pg.text), r"natural gas|06200|0620", n=4, w=200)
    for u in ("https://www.dosm.gov.my/portal-main/release-content/index-of-industrial-production-malaysia-august-2026",
              "https://www.dosm.gov.my/site/downloadrelease?id=index-of-industrial-production-malaysia-august-2026&lang=English&admin_view="):
        r = get(u)
        if r is not None and r.ok and r.content[:4] == b"%PDF":
            pdf_text(r.content, r"natural gas", maxp=2, chars=1500)
        elif r is not None and r.ok:
            ctx(plain(r.text), r"natural gas", n=3, w=200)


def bnm4():
    out("\n================ BNM Monthly Highlights & Statistics ================")
    for u in ("https://www.bnm.gov.my/monthly-highlights-statistics", "https://www.bnm.gov.my/statistics",
              "https://www.bnm.gov.my/-/monthly-highlights-and-statistics-in-august-2026"):
        r = get(u)
        if r is not None and r.ok:
            for at, au in anchors(r):
                if re.search(r"\.xlsx?|\.pdf|msb|statistic|highlight", at + au, re.I) and len(at) < 120:
                    out(f"    [{at}] {au[:200]}")
            ctx(plain(r.text), r"natural gas|LNG", n=3, w=150)



# ================================================================== ROUND 5
PX = "https://openstat.psa.gov.ph/PXWeb/api/v1/en/DB/2L/IMT"


def psa5():
    out("\n================ PSA OpenSTAT monthly imports by commodity: LNG ================")
    for tab in ("GKI/0012L4DMKD5.px", "CIF/0012L4DMVD5.px", "GKI/0022L4DMKD4.px", "CIF/0042L4DMVD2.px"):
        r = get(f"{PX}/{tab}", quiet=True)
        if r is None or not r.ok:
            out(f"  {tab}: {r.status_code if r is not None else None}")
            continue
        meta = r.json()
        out(f"\n  {tab}: {meta.get('title')}")
        sel = []
        for v in meta.get("variables", []):
            vals, txt = v.get("values", []), v.get("valueTexts", [])
            out(f"    var {v.get('code')} ({v.get('text')}): {len(vals)} values; first {list(zip(vals, txt))[:4]}")
            hit = [(a, b) for a, b in zip(vals, txt) if re.search(r"271111|liquefied natural|natural gas, liquefied|^2711", a + " " + b, re.I)]
            if hit:
                out(f"      LNG-like: {hit[:10]}")
            if re.search(r"commod", v.get("text", ""), re.I):
                codes = [a for a, _ in hit if "271111" in a] or [a for a, _ in hit][:1]
                sel.append({"code": v["code"], "selection": {"filter": "item", "values": codes}})
            elif re.search(r"country|origin", v.get("text", ""), re.I):
                sel.append({"code": v["code"], "selection": {"filter": "all", "values": ["*"]}})
            else:
                sel.append({"code": v["code"], "selection": {"filter": "all", "values": ["*"]}})
        if not any(s_["selection"]["values"] and s_["selection"]["filter"] == "item" for s_ in sel):
            continue
        q = {"query": sel, "response": {"format": "json-stat2"}}
        pr = S.post(f"{PX}/{tab}", json=q, timeout=T, verify=False)
        out(f"    POST -> {pr.status_code} {len(pr.content)}b")
        if not pr.ok:
            out(f"    {pr.text[:300]}")
            continue
        js = pr.json()
        dims, size, val = js["id"], js["size"], js["value"]
        labels = [list(js["dimension"][d]["category"]["label"].values()) for d in dims]
        import itertools
        rows = [(combo, x) for combo, x in zip(itertools.product(*labels), val) if x not in (None, 0)]
        out(f"    dims {dims} sizes {size}; non-zero cells {len(rows)}")
        for combo, x in rows[:40]:
            out(f"      {combo} = {x}")
    r = get("https://openstat.psa.gov.ph/PXWeb/api/v1/en/DB/3G/ELE/ES/0012F2BPES1.px", quiet=True)
    if r is not None and r.ok:
        for v in r.json().get("variables", []):
            out(f"  TPES var {v.get('code')}: {v.get('valueTexts')[:30]}")


def myes5():
    out("\n================ MyEnergyStats / ST publications (energy balance, handbook) ================")
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(user_agent=H["User-Agent"])
        for u in ("https://myenergystats.st.gov.my/en/", "https://www.st.gov.my/en/web/general/publications",
                  "https://www.st.gov.my/web/general/publication"):
            try:
                pg.goto(u, timeout=60000, wait_until="networkidle")
            except Exception as e:  # noqa: BLE001
                out(f"  {u}: {type(e).__name__}")
                continue
            hrefs = pg.eval_on_selector_all("a", "els => els.map(e => [e.innerText.trim(), e.href])")
            out(f"\n### {u} -> {pg.url}: {len(hrefs)} links")
            for t, h in hrefs:
                if re.search(r"publication|resource|balance|handbook|statistic|\.pdf|\.xlsx|documents/d", t + h, re.I):
                    out(f"    [{t[:70]}] {h[:200]}")
        b.close()



# ================================================================== ROUND 6
def psa6():
    out("\n================ PSA OpenSTAT: query the hidden commodity dimension for LNG ================")
    import itertools
    for tab in ("GKI/0012L4DMKD5.px", "CIF/0012L4DMVD5.px", "GKI/0042L4DMKD2.px"):
        for code in ("27111100", "2711110000", "271111", "27111100000", "2711"):
            q = {"query": [{"code": "Commodity Code", "selection": {"filter": "item", "values": [code]}},
                           {"code": "Country", "selection": {"filter": "all", "values": ["*"]}}],
                 "response": {"format": "json-stat2"}}
            try:
                pr = S.post(f"{PX}/{tab}", json=q, timeout=T, verify=False)
            except Exception as e:  # noqa: BLE001
                out(f"  {tab} {code}: {type(e).__name__}")
                continue
            if not pr.ok:
                out(f"  {tab} {code}: HTTP {pr.status_code} {pr.text[:120]}")
                continue
            js = pr.json()
            dims = js["id"]
            labels = [list(js["dimension"][d]["category"]["label"].values()) for d in dims]
            rows = [(c, x) for c, x in zip(itertools.product(*labels), js["value"]) if x not in (None, 0)]
            out(f"  {tab} {code}: OK dims {dims} sizes {js['size']}; non-zero {len(rows)}; "
                f"commodity label {labels[0][:2]}")
            for c, x in rows[:30]:
                out(f"      {c} = {x}")
            if rows:
                break


if __name__ == "__main__":
    which = sys.argv[1:] or ["psa6"]
    for w in which:
        try:
            globals()[w]()
        except Exception as e:  # noqa: BLE001
            out(f"!! {w} failed: {type(e).__name__}: {e}")
