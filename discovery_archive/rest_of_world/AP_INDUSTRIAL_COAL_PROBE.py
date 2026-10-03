"""
One-off probe (prints only) for an Asia-Pacific industrial coal-to-LNG switching study: where can coal use by
INDUSTRY SUBSECTOR and by COAL TYPE (coking vs other bituminous / sub-bituminous / lignite) be downloaded?
Round 1 checks: APEC EGEDA annual energy data (energy balance tables per economy), UNdata Energy Statistics
Database (EDATA), India Ministry of Coal statistics, and China NBS energy yearbook pages.
"""
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=60, **kw)
        print(f"{r.status_code} {len(r.content):>10,} B {r.headers.get('content-type', '')[:30]:30} {r.url[:160]}", flush=True)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"FAIL {type(e).__name__}: {str(e)[:120]} {url}", flush=True)


def links(r, pat, n=40):
    if r is None or not r.ok:
        return []
    found = sorted(set(re.findall(r'href="([^"]+)"', r.text, re.I)))
    hit = [u for u in found if re.search(pat, u, re.I)]
    for u in hit[:n]:
        print("   link:", u)
    print(f"   ({len(hit)} of {len(found)} links match)")
    return hit


def section(t):
    print("\n" + "=" * 20, t, "=" * 20, flush=True)


section("APEC EGEDA")
for u in ["https://www.egeda.ewg.apec.org/egeda/database_info/annual_data.html",
          "https://www.egeda.ewg.apec.org/egeda/database_info/index.html",
          "https://www.egeda.ewg.apec.org/",
          "https://www.apec.org/publications/2025/06/apec-energy-statistics-2022",
          "https://aperc.or.jp/publications/reports/outlook.php"]:
    links(get(u), r"\.(xlsx?|zip|csv|pdf)|annual|balance|database|statistic")

section("UNdata EDATA")
r = get("https://data.un.org/Data.aspx?d=EDATA&f=cmID%3aBH")
if r is not None and r.ok:
    print("   commodity options:", re.findall(r'<option[^>]*value="([A-Z0-9]{2,4})"[^>]*>([^<]{3,60})</option>', r.text)[:80])
r = get("https://data.un.org/Explorer.aspx?d=EDATA")
links(r, r"cmID|EDATA", 60)

section("India Ministry of Coal")
for u in ["https://coal.nic.in/en/major-statistics/coal-statistics", "https://coal.nic.in/en/major-statistics/coal-directory",
          "https://coal.nic.in/en/major-statistics/provisional-coal-statistics"]:
    links(get(u), r"\.pdf|statistic|directory")

section("IEA free balance highlights (check access)")
links(get("https://www.iea.org/data-and-statistics/data-product/world-energy-balances-highlights"), r"\.xlsx|download|highlights")
