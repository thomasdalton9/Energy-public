"""Probe 3: AESO downloads, NB Power archive, BC Hydro xls/reservoirs, CER CSV heads, StatCan fuel tables."""
import io
import re
import zipfile
import requests
import pandas as pd

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def get(u, **kw):
    try:
        return requests.get(u, headers=H, timeout=(10, 90), **kw)
    except Exception as e:  # noqa: BLE001
        print(f"ERR {u} {type(e).__name__} {str(e)[:100]}", flush=True)


def txt(html):
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", html, flags=re.S | re.I)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))


def links(r, pat=r"."):
    return [(h, re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", l))[:70]) for h, l in
            re.findall(r'href="([^"#]+)"[^>]*>(.*?)</a>', r.text, re.S | re.I) if re.search(pat, h, re.I)]


print("=== AESO pages")
for p in ["historical-generation-data", "hourly-metered-volumes-by-generation-type", "hourly-metered-volumes-by-generating-asset",
          "historical-hourly-aggregated-load-rate-dts-and-generation-rate-sts-mwh-data", "hourly-outage-by-fuel-type", "planning-area-hourly-load-and-generation"]:
    r = get(f"https://www.aeso.ca/market/market-and-system-reporting/data-requests/{p}/")
    if r is None:
        continue
    print("--", p, r.status_code)
    body = txt(r.text)
    i = body.find("Data Requests")
    print(body[i:i + 900])
    for l in links(r, r"\.(csv|xlsx?|zip)|download|assets")[:15]:
        print("   LINK", l)
r = get("https://www.aeso.ca/market/market-and-system-reporting/aeso-application-programming-interface-api/")
print("-- API page", r.status_code, txt(r.text)[400:1500])

print("=== NB Power archive")
s = requests.Session()
s.headers.update(H)
r = s.get("https://tso.nbpower.com/Public/en/system_information_archive.aspx", timeout=60)
body = txt(r.text)
i = body.find("Columns")
print(body[i:i + 2500])
vs = {k: (re.search(rf'id="{k}" value="([^"]*)"', r.text) or [None, ""])[1] for k in ["__VIEWSTATE", "__VIEWSTATEGENERATOR", "__EVENTVALIDATION"]}
print({k: len(v) for k, v in vs.items()})
print(re.findall(r'<select[^>]*name="([^"]+)"', r.text), re.findall(r'<input[^>]*name="([^"]+)"', r.text)[:10])
sel = re.findall(r'<select[^>]*name="([^"]+)"', r.text)
data = dict(vs)
data["__EVENTTARGET"] = "ctl00$cphMainContent$lbGetData"
data["__EVENTARGUMENT"] = ""
if len(sel) >= 2:
    data[sel[0]] = "9"
    data[sel[1]] = "2026"
r2 = s.post("https://tso.nbpower.com/Public/en/system_information_archive.aspx", data=data, timeout=60)
print("NB POST", r2.status_code, r2.headers.get("content-type"), r2.headers.get("content-disposition"), len(r2.content))
print(r2.text[:700])
r3 = get("https://tso.nbpower.com/Public/en/SystemInformation_realtime.asp")
print(txt(r3.text)[:1500])

print("=== BC Hydro xls")
base = "https://www.bchydro.com/content/dam/BCHydro/customer-portal/documents/corporate/suppliers/transmission-system/balancing_authority_load_data/"
for p in ["CurrentHourlyBALoad.xls", "BalancingAuthorityLoad 2026.xls"]:
    r = get(base + p)
    try:
        x = pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None)
        for k, v in x.items():
            print(p, "sheet", k, v.shape)
            print(v.head(8).to_string()[:1200])
            print(v.tail(3).to_string()[:500])
    except Exception as e:  # noqa: BLE001
        print(p, "parse fail", type(e).__name__, str(e)[:200])
for u in ["https://www.bchydro.com/energy-in-bc/operations/transmission-reservoir-data.html",
          "https://www.bchydro.com/energy-in-bc/operations/transmission-reservoir-data/previous-reservoir-elevations.html",
          "https://www.bchydro.com/energy-in-bc/operations/transmission-reservoir-data/reservoir-discharges.html"]:
    r = get(u)
    if r is not None:
        print(u, r.status_code, txt(r.text)[900:1900])
        for l in links(r, r"\.(csv|xlsx?|pdf|json|xml)|reservoir|elevation")[:25]:
            print("   LINK", l)

print("=== Manitoba")
for u in ["https://www.hydro.mb.ca/corporate/operations/", "https://www.hydro.mb.ca/corporate/operations/water-levels/",
          "https://www.hydro.mb.ca/corporate/operations/water_regimes/"]:
    r = get(u)
    if r is not None:
        print(u, r.status_code, txt(r.text)[300:900])
        for l in links(r, r"operat|water|level|export|market|data|reservoir|csv|xls")[:30]:
            print("   LINK", l)

print("=== NS dispatch / sask")
for u in ["https://www.nspower.ca/oasis/system-reports-messages", "https://www.nspower.ca/oasis/dispatch-dashboard", "https://www.nspower.ca/oasis/monthly-reports"]:
    r = get(u)
    if r is not None:
        print(u, r.status_code)
        for l in links(r, r"\.(csv|xlsx?|json)|docs/default|api|report|data|load")[:20]:
            print("   LINK", l)
        b = txt(r.text)
        j = b.find("Access Information")
        print(b[j:j + 700])
r = get("https://www.saskpower.com/about-us/our-company/power-system")
print("SK", r and r.status_code)
for u in ["https://www.saskpower.com/en/about-us/our-company/blog/", "https://www.saskpower.com/our-power-future/infrastructure-projects/power-supply",
          "https://www.saskpower.com/about-us/our-company/power-system/system-data"]:
    r = get(u)
    if r is not None:
        print(u, r.status_code, txt(r.text)[1200:1700])

print("=== CER csv heads")
for u in ["https://www.cer-rec.gc.ca/open/energy/electricity-capacity-dataset.csv",
          "https://www.cer-rec.gc.ca/open/imports-exports/natural-gas-exports-and-imports-monthly.csv",
          "https://www.cer-rec.gc.ca/open/imports-exports/electricity-exports-and-imports-monthly.csv",
          "https://www.cer-rec.gc.ca/open/energy/energyfutures2026/electricity-generation-2026.csv",
          "https://www.cer-rec.gc.ca/open/energy/energyfutures2026/electricity-generation-capacity-2026.csv",
          "https://www.cer-rec.gc.ca/open/energy/energyfutures2026/natural-gas-production-2026.csv",
          "https://www.cer-rec.gc.ca/open/energy/energyfutures2026/benchmark-prices-2026.csv"]:
    r = get(u)
    if r is not None:
        print(u, r.status_code, r.headers.get("content-type"), len(r.content), r.headers.get("last-modified"))
        if r.status_code == 200 and "html" not in r.headers.get("content-type", ""):
            print("   ", r.content[:600].decode("utf-8", "replace").replace("\n", " | "))
r = get("https://open.canada.ca/data/api/3/action/package_search", params={"q": "Energy Future 2026", "rows": 5})
for p in r.json()["result"]["results"]:
    t = (p.get("title_translated") or {}).get("en", p.get("title"))
    if "2026" in t:
        for x in p["resources"]:
            if (x.get("format") or "").upper() == "CSV" and "/open/energy" in x["url"] and "ouvert" not in x["url"]:
                print("  EF2026", x["url"])
for q in ["reservoir storage hydro", "hydroelectric water levels", "Saskatchewan electricity generation", "Alberta electricity generation by fuel"]:
    r = get("https://open.canada.ca/data/api/3/action/package_search", params={"q": q, "rows": 6})
    for p in r.json()["result"]["results"]:
        print(f"  OC[{q}]", ((p.get("title_translated") or {}).get("en") or p.get("title"))[:90])

print("=== StatCan fuel tables")
for pid in (25100084, 25100015, 25100020):
    r = get(f"https://www150.statcan.gc.ca/n1/tbl/csv/{pid}-eng.zip")
    z = zipfile.ZipFile(io.BytesIO(r.content))
    nm = [n for n in z.namelist() if n.endswith(".csv") and "MetaData" not in n][0]
    d = pd.read_csv(z.open(nm), low_memory=False)
    print(pid, d.shape, list(d.columns))
    for c in d.columns:
        if c in ("Geography", "Fuel type", "Type of electricity generation", "Class of electricity producer", "UOM", "SCALAR_FACTOR", "North American Industry Classification System (NAICS)"):
            print("   ", c, sorted(d[c].dropna().unique())[:90])
    print(d.tail(3).to_string()[:600])
