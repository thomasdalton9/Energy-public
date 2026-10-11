"""Probe 2: drill into the endpoints found by round 1 (discovery only)."""
import io
import json
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def get(u, **kw):
    try:
        r = requests.get(u, headers=H, timeout=(10, 60), **kw)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"ERR {u} {type(e).__name__} {str(e)[:100]}", flush=True)


def txt(html):
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", html, flags=re.S | re.I)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))


def links(r, pat=r"."):
    out = []
    for h, label in re.findall(r'href="([^"#]+)"[^>]*>(.*?)</a>', r.text, re.S | re.I):
        if re.search(pat, h + label, re.I):
            out.append((h, re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", label))[:80]))
    return out


print("=== AESO CSD")
r = get("http://ets.aeso.ca/ets_web/ip/Market/Reports/CSDReportServlet")
print(r.status_code, txt(r.text)[:2500])
print("=== AESO WMRQH")
r = get("http://ets.aeso.ca/ets_web/ip/Market/Reports/ActualForecastWMRQHReportServlet")
print(txt(r.text)[:800])
for u in ["http://ets.aeso.ca/ets_web/ip/Market/Reports/HistoricalDailyAveragePoolPriceReportServlet",
          "http://ets.aeso.ca/ets_web/ip/Market/Reports/SMPriceReportServlet",
          "http://ets.aeso.ca/ets_web/ip/Market/Reports/SupplyDemandReportServlet",
          "http://ets.aeso.ca/Market/Reports/Manual/Operations/prices/Hourly_Power_Pool_Price.csv"]:
    r = get(u)
    if r is not None:
        print(u, r.status_code, len(r.content), txt(r.text)[:300])
print("=== AESO data requests links")
r = get("https://www.aeso.ca/market/market-and-system-reporting/data-requests/")
for l in links(r, r"data|generation|hourly|csv|xls|stat|report"):
    print("  ", l)
for u in ["https://www.aeso.ca/market/market-and-system-reporting/", "https://www.aeso.ca/market/market-and-system-reporting/aeso-api/",
          "https://www.aeso.ca/market/market-and-system-reporting/annual-market-statistic-reports/"]:
    r = get(u)
    if r is not None:
        print(u, r.status_code)
        for l in links(r, r"stat|api|generation|csv|xls|report")[:40]:
            print("  ", l)

print("=== HQ")
for ds in ["production-electricite-quebec", "historique-production-electricite-quebec", "demande-electricite-quebec",
           "historique-demande-electricite-quebec", "donnees-hydrometeorologiques", "donnees-hydrometriques",
           "historique-production-consommation-ec-horaire"]:
    base = f"https://donnees.hydroquebec.com/api/explore/v2.1/catalog/datasets/{ds}"
    r = get(base)
    if r is None or r.status_code != 200:
        print(ds, r and r.status_code)
        continue
    j = r.json()
    print(ds, [(f["name"], f["type"]) for f in j.get("fields", [])][:25])
    print("   desc:", txt(j["metas"]["default"].get("description") or "")[:300])
    r2 = get(base + "/records?limit=2&order_by=" + ("date%20desc" if "histor" in ds else "date%20desc"))
    if r2 is not None:
        print("   rec:", r2.status_code, r2.text[:500])
    r3 = get(base + "/records?limit=1&select=min(date),max(date),count(*)") if "histor" in ds else None
    if r3 is not None:
        print("   range:", r3.text[:300])
r = get("https://www.hydroquebec.com/data/documents-donnees/donnees-ouvertes/json/production.json")
print(json.dumps(r.json()["details"][:2]))

print("=== BC Hydro")
base = "https://www.bchydro.com/content/dam/BCHydro/customer-portal/documents/corporate/suppliers/transmission-system/balancing_authority_load_data/"
for p in ["CurrentHourlyBALoad.xls", "CurrentDailyBALoad.xls", "Historical Transmission Data/BalancingAuthorityLoad 2025.xls", "BalancingAuthorityLoad 2026.xls"]:
    r = get(base + p)
    if r is None:
        continue
    print(p, r.status_code, r.headers.get("content-type"), len(r.content), r.headers.get("last-modified"))
    if r.status_code == 200:
        try:
            import pandas as pd
            x = pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None)
            for k, v in x.items():
                print("  sheet", k, v.shape)
                print(v.head(6).to_string()[:900])
                print(v.tail(3).to_string()[:400])
        except Exception as e:  # noqa: BLE001
            print("  parse fail", type(e).__name__, str(e)[:200], r.content[:200])
r = get("https://www.bchydro.com/info/nsi-data/BCHNSI.xml")
print("NSI", r and r.status_code, r and r.text[:400])
r = get("https://www.bchydro.com/energy-in-bc/operations.html")
if r is not None:
    for l in links(r, r"reservoir|water|level|generation")[:20]:
        print("  BCH", l)

print("=== SaskPower")
r = get("https://www.saskpower.com/about-us/our-company/power-system/power-generation")
t = r.text
for m in set(re.findall(r'(https?://[^"\'\s]+(?:json|api|csv|arcgis|xlsx?)[^"\'\s]*)', t, re.I)):
    print("  SP", m[:200])
print(txt(t)[:600])
for u in ["https://www.saskpower.com/about-us/our-company/blog/saskpower-system-data",
          "https://www.saskpower.com/api/system-data", "https://www.saskpower.com/-/media/SaskPower/documents/system-data.json"]:
    pass

print("=== NB Power")
r = get("https://tso.nbpower.com/Public/en/system_information_archive.aspx")
for l in links(r)[:60]:
    print("  NB", l)
print(txt(r.text)[:800])
r = get("https://tso.nbpower.com/Public/en/system_information.aspx")
for u in ["https://tso.nbpower.com/Public/en/SystemInformation_realtime.asp", "https://tso.nbpower.com/Public/en/system_information_realtime.aspx",
          "https://tso.nbpower.com/Public/en/system_information_report.aspx"]:
    r = get(u)
    if r is not None:
        print(u, r.status_code, txt(r.text)[:400])

print("=== NS Power")
for u in ["https://www.nspower.ca/", "https://www.nspower.ca/about-us/electricity", "https://www.nspower.ca/oasis"]:
    r = get(u)
    if r is not None:
        print(u, r.status_code)
        for l in links(r, r"oasis|system|data|load|report|emission|generation")[:25]:
            print("  NS", l)

print("=== Manitoba")
for u in ["https://www.hydro.mb.ca/", "https://www.hydro.mb.ca/corporate/", "https://www.hydro.mb.ca/environment/water_regimes/"]:
    r = get(u)
    if r is not None:
        print(u, r.status_code)
        for l in links(r, r"water|level|reservoir|river|data|export|market")[:25]:
            print("  MB", l)

print("=== CER")
r = get("https://www.cer-rec.gc.ca/open/energy/")
print(txt(r.text)[:800])
for l in links(r)[:40]:
    print("  CER", l)
for q in ["energy futures", "natural gas exports imports", "electricity generation", "pipeline throughput", "reservoir"]:
    r = get("https://open.canada.ca/data/api/3/action/package_search", params={"q": q, "rows": 12, "fq": "organization:cer-rec"})
    if r is not None and r.status_code == 200:
        for p in r.json()["result"]["results"]:
            res = [(x.get("format"), x.get("url")) for x in p.get("resources", []) if (x.get("format") or "").upper() in ("CSV", "XLSX", "ZIP", "XLS")][:3]
            print(f"  OC[{q}] {p['title'][:80] if isinstance(p['title'], str) else p.get('title_translated', {}).get('en','')[:80]} {res}")
r = get("https://open.canada.ca/data/api/3/action/package_search", params={"q": "natural gas exports imports Canada Energy Regulator", "rows": 8})
for p in r.json()["result"]["results"]:
    print("  OC2", (p.get("title_translated") or {}).get("en", p.get("title"))[:90], [(x.get("format"), x.get("url")) for x in p["resources"] if (x.get("format") or "").upper() in ("CSV","XLSX")][:2])

print("=== StatCan")
for pid in [25100084, 25100015, 25100020, 25100021, 25100019, 25100016, 25100017, 25100055, 25100079]:
    try:
        r = requests.post("https://www150.statcan.gc.ca/t1/wds/rest/getCubeMetadata", json=[{"productId": pid}], headers=H, timeout=40)
        o = r.json()[0]["object"]
        print(pid, o["cubeTitleEn"], o["cubeStartDate"], o["cubeEndDate"], o["frequencyCode"],
              [(d["dimensionNameEn"], len(d["member"])) for d in o["dimension"]])
    except Exception as e:  # noqa: BLE001
        print(pid, "ERR", type(e).__name__, str(e)[:100])
