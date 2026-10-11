"""Probe 4."""
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



print("=== AESO xlsx")
for u in ["https://www.aeso.ca/assets/Uploads/Hourly-Metered-Volumes-by-Generation-Type.xlsx"]:
    r = get(u)
    print(u, r.status_code, len(r.content), r.headers.get("last-modified"))
    try:
        x = pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None)
        for k, v in x.items():
            print("  sheet", k, v.shape)
            print(v.head(8).to_string()[:1800])
            print(v.tail(4).to_string()[:900])
    except Exception as e:  # noqa: BLE001
        print("  parse fail", type(e).__name__, str(e)[:200], r.content[:200])
r = requests.get("https://www.aeso.ca/assets/Uploads/Hourly-Metered-Volumes-by-Generating-Asset.csv", headers=H, stream=True, timeout=60)
print("asset csv", r.status_code, r.headers.get("content-length"), r.headers.get("last-modified"))
print(next(r.iter_lines()), next(r.iter_lines()))
r.close()
r = get("https://www.aeso.ca/market/market-and-system-reporting/data-requests/historical-generation-data/")
b = txt(r.text)
i = b.find("Historical Generation Data (CSD)", 3000)
print(b[i:i + 1800])
for l in links(r, r"csd|generation|api|apimanagement|developer")[:20]:
    print("   LINK", l)
for u in ["https://api.aeso.ca/report/v1/csd/summary/current", "https://developer-apim.aeso.ca/", "https://apimgw.aeso.ca/public/v1/csd/summary/current"]:
    r = get(u)
    if r is not None:
        print(u, r.status_code, r.text[:200])

print("=== BC reservoir")
for reg in ["columbia", "peace", "lower-mainland", "vancouver_island"]:
    r = get(f"https://www.bchydro.com/energy-in-bc/operations/transmission-reservoir-data/previous-reservoir-elevations/{reg}.html")
    if r is not None:
        b = txt(r.text)
        i = b.find("Reservoir")
        print(reg, r.status_code, len(r.text))
        for m in re.finditer(r"(Arrow|Williston|Kinbasket|Revelstoke|Upper Campbell)", b):
            print("   ", b[max(0, m.start() - 100):m.start() + 300])
            break
        print("   tables:", len(re.findall("<table", r.text)), [h for h, _ in links(r, r"\.(csv|xls|json)|chart|data")[:10]])
        print("   scripts:", re.findall(r'src="([^"]*(?:chart|reservoir|data)[^"]*)"', r.text)[:10])

print("=== MB water")
for u in ["https://www.hydro.mb.ca/corporate/operations/water-levels/hydrological-data/", "https://www.hydro.mb.ca/corporate/operations/generation/", "https://www.hydro.mb.ca/corporate/operations/transmission/"]:
    r = get(u)
    if r is not None:
        b = txt(r.text)
        i = b.find("Skip to content")
        print(u, r.status_code, len(r.text))
        j = b.find("Hydrological data", 1500)
        print("  ", b[j:j + 1200])
        print("  links", [l for l in links(r, r"\.(csv|xlsx?|json|pdf)|data|gauge|levels")[:15]])

print("=== NS daily / SK")
r = get("https://www.nspower.ca/oasis/system-reports-messages/daily-report")
if r is not None:
    print(r.status_code, [l for l in links(r, r"docs|report|csv|xls|pdf")[:12]])
r = get("https://www.saskpower.com/about-us/our-company/power-system/system-data")
b = txt(r.text)
i = b.find("System Data")
print("SK", len(b), b[i:i + 600])
print(sorted(set(re.findall(r'["\'(]([^"\'()\s]*(?:\.json|api/|\.csv|powerdata|generation)[^"\'()\s]*)', r.text)))[:30])

print("=== CER electricity")
for f in ["electricity-generation-2026", "electricity-capacity-2026", "electricity-interchange-2026"]:
    r = get(f"https://www.cer-rec.gc.ca/open/energy/energyfutures2026/{f}.csv")
    d = pd.read_csv(io.BytesIO(r.content))
    print(f, d.shape, list(d.columns))
    for c in d.columns:
        if c not in ("Value", "Year"):
            print("   ", c, sorted(d[c].dropna().unique())[:40])
    print("   years", d["Year"].min(), d["Year"].max())
d = pd.read_csv(io.BytesIO(get("https://www.cer-rec.gc.ca/open/energy/energyfutures2026/electricity-generation-2026.csv").content))
x = d[(d.Scenario == d.Scenario.iloc[0]) & (d.Year.isin([2022, 2023, 2024, 2025])) & (d.Region == "Alberta")]
print(x.to_string()[:2500])
r = get("https://www.cer-rec.gc.ca/open/energy/energyfutures2026/EF2026-data-dictionary.csv")
print(r.text[:2500])

print("=== StatCan")
r = get("https://www150.statcan.gc.ca/n1/tbl/csv/25100084-eng.zip")
z = zipfile.ZipFile(io.BytesIO(r.content))
d = pd.read_csv(z.open([n for n in z.namelist() if "MetaData" not in n and n.endswith(".csv")][0]), low_memory=False)
print(sorted(d["Fuel type"].unique()))
print(sorted(d.GEO.unique()))
x = d[(d["Fuel type"].str.contains("electricity generated")) & (d.REF_DATE == 2024) & (d.GEO == "Alberta")]
print(x[["North American Industry Classification System (NAICS)", "Fuel type", "UOM", "SCALAR_FACTOR", "VALUE"]].to_string())
r = get("https://www150.statcan.gc.ca/n1/tbl/csv/25100015-eng.zip")
z = zipfile.ZipFile(io.BytesIO(r.content))
d = pd.read_csv(z.open([n for n in z.namelist() if "MetaData" not in n and n.endswith(".csv")][0]), low_memory=False)
print(sorted(d["Type of electricity generation"].unique()), sorted(d["Class of electricity producer"].unique()), d.REF_DATE.max())
x = d[(d.REF_DATE == d.REF_DATE.max())]
print(x[x["Class of electricity producer"].str.startswith("Total")].pivot_table(index="Type of electricity generation", columns="GEO", values="VALUE").iloc[:, :8].to_string())

print("=== HQ levels")
base = "https://donnees.hydroquebec.com/api/explore/v2.1/catalog/datasets/donnees-hydrometeorologiques/records"
r = get(base, params={"limit": 5, "where": "composition_depil_type_mesure like \"niveau\"", "order_by": "date desc"})
print(r.status_code, r.text[:1500])
r = get(base, params={"limit": 0, "group_by": "composition_depil_type_mesure"})
print(r.status_code, r.text[:800])
