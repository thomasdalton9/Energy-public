"""
One-off probe (prints only) for an Asia-Pacific industrial coal-to-LNG switching study: where can coal use by
INDUSTRY SUBSECTOR and by COAL TYPE (coking vs thermal) be downloaded?
Round 1: IEA 403; India coal.nic.in pages 404; UNdata EDATA page loads but no commodity codes found;
APEC EGEDA annual page links rev_newbalance_select_form2.html (energy balance form) and ../database/wnewcgi.php.
Round 2: EGEDA forms (actions, select names, options), UNdata EDATA page source around cmID, India coal on coal.gov.in.
"""
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
EG = "https://www.egeda.ewg.apec.org/egeda/"


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=60, **kw)
        print(f"{r.status_code} {len(r.content):>10,} B {r.headers.get('content-type', '')[:30]:30} {r.url[:160]}", flush=True)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"FAIL {type(e).__name__}: {str(e)[:120]} {url}", flush=True)


def forms(r):
    if r is None or not r.ok:
        return
    for f in re.findall(r"<form.*?</form>", r.text, re.S | re.I):
        print("  FORM", re.findall(r'<form[^>]*>', f, re.I)[0][:200])
        for name, body in re.findall(r'<select[^>]*name="([^"]+)"[^>]*>(.*?)</select>', f, re.S | re.I):
            opts = re.findall(r'<option[^>]*value="([^"]*)"[^>]*>\s*([^<]*)', body, re.I)
            print(f"    select {name}: {len(opts)} options e.g. {opts[:25]}")
        for inp in re.findall(r'<input[^>]*>', f, re.I)[:20]:
            print("    input", inp[:160])
    print("  links:", sorted(set(re.findall(r'href="([^"]+)"', r.text)))[:30])


def section(t):
    print("\n" + "=" * 20, t, "=" * 20, flush=True)


section("EGEDA balance form")
forms(get(EG + "database_info/rev_newbalance_select_form2.html"))
section("EGEDA wnewcgi")
forms(get(EG + "database/wnewcgi.php"))
section("EGEDA database_information")
r = get(EG + "database_info/database_information.html")
if r is not None and r.ok:
    print(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))[:2500])

section("UNdata EDATA page source")
r = get("https://data.un.org/Data.aspx?d=EDATA&f=cmID%3aBH")
if r is not None and r.ok:
    t = r.text
    for m in list(re.finditer(r"cmID|DownloadHandler|trID", t))[:12]:
        print("   ...", re.sub(r"\s+", " ", t[max(0, m.start() - 150):m.start() + 200]))
for u in ["https://data.un.org/Handlers/DownloadHandler.ashx?DataFilter=cmID:BH&DataMartId=EDATA&Format=csv",
          "https://data.un.org/ws/rest/dataflow/all/all/latest"]:
    r = get(u)
    if r is not None:
        print("   first bytes:", r.content[:400])

section("India coal (coal.gov.in)")
for u in ["https://coal.gov.in/en/major-statistics/coal-statistics", "https://coal.gov.in/en/major-statistics",
          "https://coal.gov.in/en/public-information/reports/annual-report"]:
    r = get(u)
    if r is not None and r.ok:
        hits = sorted(set(u2 for u2 in re.findall(r'href="([^"]+)"', r.text) if re.search(r"statistic|directory|\.pdf", u2, re.I)))
        print("   ", hits[:40])
