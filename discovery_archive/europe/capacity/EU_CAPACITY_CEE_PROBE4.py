"""
Probe 4 (CEE capacity group). Output: results/probe4.txt
 - CH: BFE plant register (ElectricityProductionPlant.csv + PlantDetail.csv) rebuilt to 1 January values by category
 - PL: URE 'Instalacje OZE' quarterly xlsx structure; PSE site crawl for monthly KSE reports
 - HU: MAVIR 'Energy MIX - Power Plants - Installed Capacity Data' page
 - CZ: ERU 'Zpravy o provozu', CEPS transmission-system-data
 - SK: Power BI embed behind installedpower.sepsas.sk
 - RO: ANRE puteri-instalate iframes; Transelectrica
 - BG: ESO pbook, DKER proizvodstvo
"""
import io, json, os, re, zipfile
import requests
import pandas as pd
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36", "Accept": "*/*"}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results"); os.makedirs(OUT, exist_ok=True)
LOG = open(os.path.join(OUT, "probe4.txt"), "w")
def p(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); LOG.write(s + "\n"); LOG.flush()
def get(url, **kw):
    try: return requests.get(url, headers=H, timeout=(10, 120), **kw)
    except Exception as e: p("ERR", url, type(e).__name__, str(e)[:120]); return None
def sq(t, n=400): return re.sub(r"\s+", " ", t)[:n]
def links(name, url, pat, maxn=60, base=None):
    r = get(url)
    if r is None: return None
    p(f"== {name} {url} -> {r.status_code} {len(r.content)}")
    n = 0; seen = set()
    for m in re.finditer(r'href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', r.text, re.S | re.I):
        t = sq(re.sub(r"<[^>]+>", " ", m.group(2)), 90)
        if re.search(pat, m.group(1) + " " + t, re.I) and m.group(1) not in seen:
            seen.add(m.group(1)); p("    ", t, "|", m.group(1)[:170]); n += 1
            if n >= maxn: break
    return r
p("=== CH register")
r = get("https://data.geo.admin.ch/ch.bfe.elektrizitaetsproduktionsanlagen/csv/2056/ch.bfe.elektrizitaetsproduktionsanlagen.zip")
z = zipfile.ZipFile(io.BytesIO(r.content))
pl = pd.read_csv(z.open("ElectricityProductionPlant.csv")); det = pd.read_csv(z.open("PlantDetail.csv"))
pl["BeginningOfOperation"] = pd.to_datetime(pl["BeginningOfOperation"], errors="coerce")
p("rows", len(pl), "no start date", pl["BeginningOfOperation"].isna().sum(), "range", pl["BeginningOfOperation"].min(), pl["BeginningOfOperation"].max())
p("TotalPower unit assumed kW; sum by Sub/Plant category (MW):")
g = pl.groupby(["MainCategory", "SubCategory", "PlantCategory"], dropna=False)["TotalPower"].agg(["count", "sum"]); g["MW"] = g["sum"] / 1000; p(g.to_string())
det["Date"] = pd.to_datetime(det["Date"], errors="coerce")
p("detail rows", len(det), "date range", det["Date"].min(), det["Date"].max(), "Power sum (kW?)", det["Power"].sum())
p("Initial vs Total power: plants where Total != Initial", int((pl["TotalPower"] != pl["InitialPower"]).sum()))
# 1 January cumulative by subcategory: plants started before date, TotalPower
for sub in ["subcat_1", "subcat_2", "subcat_3", "subcat_4", "subcat_6", "subcat_7", "subcat_8", "subcat_9", "subcat_10"]:
    s = pl[pl.SubCategory == sub]
    row = []
    for y in range(2020, 2027):
        d = pd.Timestamp(f"{y}-01-01"); row.append(round(s[s.BeginningOfOperation < d]["TotalPower"].sum() / 1000))
    p(sub, row)
# PV via detail: sum of Power (kW) of increments dated before Jan 1
pv = det[det.PlantCategory.isin(["plantcat_8", "plantcat_9", "plantcat_10"])]
p("PV by detail Jan-1 (MW):", [round(pv[pv.Date < pd.Timestamp(f'{y}-01-01')]["Power"].sum() / 1000) for y in range(2020, 2027)])
pvp = pl[pl.SubCategory == "subcat_2"]
p("PV plants InitialPower Jan-1:", [round(pvp[pvp.BeginningOfOperation < pd.Timestamp(f'{y}-01-01')]["InitialPower"].sum() / 1000) for y in range(2020, 2027)])
p("PV monthly additions 2025 by plant start (MW):", (pvp[pvp.BeginningOfOperation.dt.year == 2025].groupby(pvp.BeginningOfOperation.dt.month)["InitialPower"].sum() / 1000).round(0).to_dict())
p("Detail monthly additions 2025 (MW):", (pv[pv.Date.dt.year == 2025].groupby(pv.Date.dt.month)["Power"].sum() / 1000).round(0).to_dict())
p("=== PL")
r = links("URE OZE page", "https://www.ure.gov.pl/pl/oze/potencjal-krajowy-oze/8108,Instalacje-odnawialnych-zrodel-energii-stan-na-30-czerwca-2026-r.html", r"xlsx")
x = get("https://www.ure.gov.pl/download/9/16383/mapaozena30-06-2026.xlsx")
if x is not None and x.status_code == 200:
    p("xlsx", len(x.content), x.headers.get("last-modified"))
    xl = pd.ExcelFile(io.BytesIO(x.content)); p(xl.sheet_names)
    for s in xl.sheet_names[:6]:
        d = xl.parse(s, header=None); p("--", s, d.shape); p(d.head(12).to_string()[:3000])
for u in ["https://www.pse.pl/dane-systemowe/funkcjonowanie-kse", "https://www.pse.pl/dane-systemowe/funkcjonowanie-kse/raporty-roczne-z-funkcjonowania-kse-za-rok",
          "https://www.pse.pl/dane-systemowe/funkcjonowanie-kse/raporty-miesieczne-z-funkcjonowania-kse", "https://www.pse.pl/web/pse-eng/data/kse-operation"]:
    links("PSE " + u, u, r"raport|report|moc|capacity|xls|pdf|zainst|miesi|month", 40)
links("ARE", "https://www.are.waw.pl/", r"statyst|moc|capacity|raport|xls", 30)
p("=== HU")
links("MAVIR mix", "https://mavir.hu/web/mavir-en/energy-mix-power-plants-installed-capacity-data", r"xls|csv|document|\.zip|data|mix|capacity|instal", 60)
r = get("https://mavir.hu/web/mavir-en/energy-mix-power-plants-installed-capacity-data")
if r is not None:
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", r.text, flags=re.S); t = sq(re.sub(r"<[^>]+>", " ", t), 6000); i = t.find("Energy MIX"); p(t[i:i+2500])
p("=== CZ")
links("ERU zpravy-o-provozu", "https://www.eru.gov.cz/zpravy-o-provozu", r"zpr|provoz|xls|pdf|měsíč|mesic|elektr", 50)
links("CEPS tsdata", "https://www.ceps.cz/en/transmission-system-data", r"data|instal|generation|capacity", 50)
links("CEPS all-data full", "https://www.ceps.cz/en/all-data", r"/en/", 120)
p("=== SK Power BI")
r = get("https://installedpower.sepsas.sk/")
if r is not None:
    for m in re.finditer(r"resourceKey|resolvedClusterUri|embedType|modelsAndExploration|ReportEmbed", r.text): pass
    keys = set(re.findall(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", r.text)); p("uuids in page", keys)
    for m in re.finditer(r"(resourceKey|tenantId|clusterUri|ClusterUri)\W{0,10}[:=]\W{0,10}['\"]([^'\"]{3,120})", r.text): p("   ", m.group(1), m.group(2))
    i = r.text.find("resourceKey"); p(sq(r.text[i-300:i+600], 900))
p("=== RO")
r = get("https://anre.ro/puteri-instalate")
if r is not None:
    for m in re.finditer(r"<iframe[^>]+>", r.text): p("iframe", m.group(0)[:300])
    for m in re.finditer(r'(?:powerbi|datawrapper|tableau|arcgis|grafana)[^"\' <]{0,200}', r.text, re.I): p("embed", m.group(0)[:250])
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", r.text, flags=re.S); t = sq(re.sub(r"<[^>]+>", " ", t), 20000); i = t.find("Puteri Instalate", 200); p(t[i:i+1800])
    for m in re.finditer(r'href=["\']([^"\']+\.(?:xlsx?|csv|pdf))["\']', r.text): p("file", m.group(1))
links("ANRE rapoarte", "https://anre.ro/despre/rapoarte/", r"raport|monitor|capacit|putere", 30)
p("=== BG")
links("ESO pbook", "https://www.eso.bg/doc/?pbook", r"xls|pdf|doc|книжка|pbook|\d{4}", 40)
links("DKER proizvodstvo", "https://www.dker.bg/bg/elektroenergetika/proizvodstvo.html", r"xls|pdf|doc|upload|мощност|производ", 40)
links("DKER EN", "https://www.dker.bg/en/home", r"report|statistic|energy|electric|xls|pdf", 40)
