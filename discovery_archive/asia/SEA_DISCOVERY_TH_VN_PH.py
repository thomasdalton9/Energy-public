"""
South & Southeast Asia dashboard, discovery round 1: Thailand, Vietnam, Philippines.
Web research only so far (the sites are blocked from the Claude sandbox); this checks
from a GitHub Actions runner which candidate raw sources answer and what they hold:

  Thailand    EPPO electricity / natural gas statistics xls (new and old site), EGAT open-data
              CKAN catalogue, RID daily large-reservoir report, thaiwater.net API
  Vietnam     NSMO home page (daily output / mix / market price), EVN reservoir portal
              hochuathuydien.evn.com.vn, statistics office (nso.gov.vn)
  Philippines IEMOP market data pages (WESM csv/zip reports), PAGASA dam levels,
              DOE power statistics, PSA OpenSTAT (PXWeb) trade API

Prints status, content type, size, the page title, a text snippet and every link
that looks like data (xls/xlsx/csv/zip/json/pdf/api). Nothing is written.
"""
import re
import sys
from urllib.parse import urljoin

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36",
     "Accept": "text/html,application/json,application/xhtml+xml,*/*;q=0.8"}
T = (15, 60)
DATA_LINK = re.compile(r"""(?:href|src|url)\s*[=:]\s*["']([^"']+?(?:\.xlsx?|\.csv|\.zip|\.json|\.pdf|/api/[^"']*|"""
                       r"""download[^"']*))["']""", re.I)


def out(*a):
    print(*a, flush=True)


def probe(label, url, show=1500, links=60, **kw):
    out(f"\n{'=' * 90}\n{label}: {url}")
    try:
        r = requests.get(url, headers=H, timeout=T, allow_redirects=True, **kw)
    except Exception as e:
        out(f"  ERROR {type(e).__name__}: {e}")
        return None
    ct = r.headers.get("content-type", "")
    out(f"  -> {r.status_code} {ct} {len(r.content)} bytes final={r.url} last-modified={r.headers.get('last-modified')}")
    if "html" in ct or "json" in ct or "text" in ct or "javascript" in ct:
        t = r.text
        m = re.search(r"<title[^>]*>(.*?)</title>", t, re.S | re.I)
        if m:
            out(f"  title: {m.group(1).strip()[:200]}")
        found = list(dict.fromkeys(urljoin(r.url, x) for x in DATA_LINK.findall(t)))
        if found:
            out(f"  data-like links ({len(found)}):")
            for x in found[:links]:
                out(f"    {x}")
        txt = re.sub(r"<script.*?</script>|<style.*?</style>", " ", t, flags=re.S | re.I)
        txt = re.sub(r"<[^>]+>", " ", txt)
        txt = re.sub(r"\s+", " ", txt)
        out(f"  text: {txt[:show]}")
    return r


def thailand():
    for u in ["https://www.eppo.go.th/index.php/en/en-energystatistics/electricity-statistic",
              "https://www.eppo.go.th/index.php/th/energy-information/static-energy/static-electricity",
              "https://www.eppo.go.th/index.php/en/en-energystatistics/ngv-statistic",
              "https://www.eppo.go.th/index.php/th/energy-information/static-energy/static-gas",
              "https://old.eppo.go.th/index.php/en/en-energystatistics/electricity-statistic",
              "https://www.eppo.go.th/epposite/info/stat/electricity",
              "https://www.eppo.go.th/"]:
        probe("EPPO", u)
    probe("EGAT CKAN package", "https://gdcatalog.egat.co.th/api/3/action/package_list")
    probe("EGAT CKAN search", "https://gdcatalog.egat.co.th/api/3/action/package_search?rows=50")
    probe("RID reservoir", "https://app.rid.go.th/reservoir/")
    probe("RID reservoir api", "https://app.rid.go.th/reservoir/api/dam/public")
    probe("thaiwater dam", "https://api-v3.thaiwater.net/api/v1/thaiwater30/public/dam_load")
    probe("thaiwater dam daily", "https://api-v3.thaiwater.net/api/v1/thaiwater30/public/thailand_main")
    probe("EGAT water", "http://water.egat.co.th/")


def vietnam():
    probe("NSMO", "https://www.nsmo.vn/", show=3000)
    probe("NSMO en", "https://www.nsmo.vn/en/")
    probe("EVN reservoirs", "https://hochuathuydien.evn.com.vn/", show=3000)
    probe("EVN reservoirs page", "https://hochuathuydien.evn.com.vn/PageHoChuaThuyDienEmbedEVN.aspx", show=3000)
    probe("NSO", "https://www.nso.gov.vn/en/")
    probe("EVN", "https://www.evn.com.vn/")


def philippines():
    for u in ["https://www.iemop.ph/market-data/",
              "https://www.iemop.ph/market-data/rtd-market-prices/",
              "https://www.iemop.ph/market-data/rtd-regional-summaries/",
              "https://www.iemop.ph/market-data/dipc-energy-results-raw/",
              "https://www.iemop.ph/market-data/gesq/",
              "https://www.wesm.ph/market-outcomes/market-data"]:
        probe("IEMOP", u, show=2500)
    probe("PAGASA dams", "https://www.pagasa.dost.gov.ph/flood", show=4000)
    probe("PAGASA dams (bagong)", "https://bagong.pagasa.dost.gov.ph/flood", show=4000)
    probe("DOE power stats", "https://doe.gov.ph/energy-statistics/power-statistics")
    probe("DOE natgas", "https://doe.gov.ph/natgas")
    probe("PSA OpenSTAT", "https://openstat.psa.gov.ph/PXWeb/api/v1/en/DB/")


if __name__ == "__main__":
    which = sys.argv[1:] or ["thailand", "vietnam", "philippines"]
    for f in which:
        try:
            globals()[f]()
        except Exception as e:
            out(f"!! {f}: {e!r}")
