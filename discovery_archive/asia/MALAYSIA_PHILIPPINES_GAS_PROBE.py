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


if __name__ == "__main__":
    which = sys.argv[1:] or ["my", "ph", "ct"]
    if "ct" in which:
        comtrade()
    if "my" in which:
        malaysia()
    if "ph" in which:
        philippines()
