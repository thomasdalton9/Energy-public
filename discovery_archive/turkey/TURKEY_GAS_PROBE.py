"""Probe: where can we get Turkey's monthly natural gas consumption (ideally by sector) from GitHub Actions?
Candidates: EPDK monthly gas sector report page/PDFs, BOTAS site, EPIAS gas transparency API, Eurostat nrg_cb_gasm (geo=TR),
JODI gas. Prints status, size and the first lines/links of each; saves what it fetched to discovery_archive/results/turkey/."""
import json, os, re, sys
from curl_cffi import requests as cr

OUT = "discovery_archive/results/turkey"
os.makedirs(OUT, exist_ok=True)
H = {"Accept": "*/*", "Accept-Language": "tr,en;q=0.8"}
S = cr.Session()


def get(name, url, **kw):
    try:
        r = S.get(url, impersonate="chrome", timeout=40, headers=H, allow_redirects=True, **kw)
    except Exception as e:  # noqa: BLE001
        print(f"[{name}] ERR {type(e).__name__}: {str(e)[:120]}")
        return None
    ct = r.headers.get("content-type", "")
    print(f"[{name}] {r.status_code} {len(r.content)}B {ct[:40]}  {url[:110]}")
    if r.status_code == 200:
        ext = "pdf" if "pdf" in ct else "json" if "json" in ct else "xlsx" if "sheet" in ct or "excel" in ct else "html"
        open(f"{OUT}/{name}.{ext}", "wb").write(r.content[:3_000_000])
    return r


# 1. EPDK monthly gas sector report page -> report links
r = get("epdk_page", "https://www.epdk.gov.tr/Detay/Icerik/3-0-95-1007/dogal-gazaylik-sektor-raporu")
if r is not None and r.status_code == 200:
    links = re.findall(r'href="([^"]+\.(?:pdf|xlsx?|zip)[^"]*)"', r.text, flags=re.I)
    print("  EPDK report links:", len(links))
    for l in links[:12]:
        print("   ", l)
    if links:
        u = links[0] if links[0].startswith("http") else "https://www.epdk.gov.tr" + links[0]
        get("epdk_first_report", u)

# 2. EPIAS natural-gas transparency (several guesses; the electricity platform lives at seffaflik.epias.com.tr)
for i, u in enumerate([
    "https://seffaflik.epias.com.tr/natural-gas/",
    "https://seffaflik.epias.com.tr/natural-gas-service/v1/",
    "https://seffaflik.epias.com.tr/natural-gas-service/swagger-ui/index.html",
    "https://seffaflik.epias.com.tr/natural-gas-service/v3/api-docs",
    "https://seffaflik.epias.com.tr/electricity-service/v3/api-docs",
    "https://gunlukdogalgaz.epias.com.tr/",
    "https://www.epias.com.tr/dogal-gaz-piyasalari/",
]):
    get(f"epias_{i}", u)

# 3. BOTAS
for i, u in enumerate([
    "https://www.botas.gov.tr/",
    "https://www.botas.gov.tr/Sayfa/dogal-gaz-tuketim/",
    "https://www.botas.gov.tr/Sayfa/sektor-raporlari/",
    "https://www.botas.gov.tr/Sayfa/aylik-bulten/",
]):
    get(f"botas_{i}", u)

# 4. Eurostat monthly gas (Turkey is a candidate-country reporter for some datasets)
r = get("eurostat_gasm", "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_cb_gasm?geo=TR&format=JSON&lang=EN&sinceTimePeriod=2023-01")
if r is not None and r.status_code == 200:
    try:
        j = r.json()
        print("  eurostat dims:", {k: list(v["category"]["index"])[:8] for k, v in j["dimension"].items()}, "values:", len(j.get("value", {})))
    except Exception as e:  # noqa: BLE001
        print("  eurostat parse", e, r.text[:200])

# 5. JODI gas (monthly demand/production/imports by country)
get("jodi_gas_csv", "https://www.jodidata.org/_resources/files/downloads/gas-data/jodi_gas_csv_beta.zip")
get("jodi_page", "https://www.jodidata.org/gas/database/data-downloads.aspx")

# 6. IEA / Enerji Atlasi / Ministry (ETKB) energy balance
get("etkb_gas", "https://enerji.gov.tr/bilgi-merkezi-doğal-gaz")
get("etkb_denge", "https://enerjiatlasi.com/dogalgaz/")
