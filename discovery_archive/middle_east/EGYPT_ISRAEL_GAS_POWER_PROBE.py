"""
Discovery probe: raw gas and power data for EGYPT and ISRAEL (Oct 2026).

Round 1 (requests only): fetch the candidate official sites, print status,
content type, size, <title>, and the links on each page that look like data
(statistics / reports / xlsx / csv / pdf / api), plus any API-looking URLs in
inline scripts. Also queries Israel's CKAN open-data portal (data.gov.il) for
gas and electricity datasets.

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

if __name__ == "__main__":
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
