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
doe2, doe2_render (run with function names as arguments; default = latest round).
Prints status, size, data-like links and text around gas keywords. Nothing is written.
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


if __name__ == "__main__":
    which = sys.argv[1:] or ["meih2", "dosm2", "mets2", "petronas2", "comtrade2", "doe2", "doe2_render"]
    for w in which:
        try:
            globals()[w]()
        except Exception as e:  # noqa: BLE001
            out(f"!! {w} failed: {type(e).__name__}: {e}")
