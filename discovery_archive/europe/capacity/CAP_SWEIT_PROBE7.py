"""Probe 7: France technology values + SDES national rows, Spain alternatives (datos.gob.es, REE pages), Terna download-centre JS, PT older PDFs.
Probe 6 findings (12 Oct 2026): register non-aggregated rows are date-accurate (solar non-aggregated date-based within about 2% of the snapshot's
non-aggregated part) but <36 kW aggregates (solar 6.15 GW now, 5.7 GW at end-2025) carry one early date each; undated old hydro (5.97 GW) and
thermal (5.05 GW) rows exist; closed plants vanish from the current register. SDES quarterly xlsx (media/9714 solar, media/9715 wind) are
regional/departmental quarterly tables since 2005 (nombre/puissance per quarter-end). DGEG archive: monthly PDFs dgeg-arr-YYYY-MM from 2024-01, year-end PDFs
before. REData potencia-instalada 500 for every variant. Terna download centre datasets: CapacityRenewableSources, CapacityThermal, GenerationPlants."""
import json, re, subprocess, sys
from html import unescape
import pandas as pd
pd.set_option("display.width", 250); pd.set_option("display.max_columns", 30); pd.set_option("display.max_rows", 200)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0"

def get(u, n=800, show=True, hdr=(), out=None):
    cmd = ["curl", "-sL", "-m", "120", "-A", UA] + [x for h in hdr for x in ("-H", h)]
    if out:
        cmd += ["-o", out]
    cmd += ["-w", "\n#HTTP %{http_code} %{content_type} %{size_download}", u]
    o = subprocess.run(cmd, capture_output=True, text=True, timeout=150)
    t = o.stdout
    head = t[t.rfind("#HTTP"):].strip()
    body = t[:t.rfind("#HTTP")]
    if show:
        print("\n=====", u[:230], "|", head, flush=True)
        print(body[:n].replace("\n", " "), flush=True)
    return body

B = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/"
R = "registre-national-installation-production-stockage-electricite-agrege"
print("#### A. France technologies")
p = "/tmp/cur.csv"
get(f"{B}{R}/exports/csv?delimiter=%3B&use_labels=false", show=False, out=p)
d = pd.read_csv(p, sep=";", low_memory=False)
d["mw"] = pd.to_numeric(d["puismaxinstallee"], errors="coerce") / 1000
d["agg"] = d["nominstallation"].astype(str).str.startswith("Agr")
print(d.columns.tolist())
g = d.groupby(["filiere", "technologie", "codecombustible", "typestockage"], dropna=False).agg(n=("mw", "size"), mw=("mw", "sum"), agg=("agg", "sum")).round(1)
print(g.to_string())
print(d.groupby(["filiere", "regime"], dropna=False).mw.agg(["size", "sum"]).round(1).to_string())
print(d[d.filiere.isin(["Thermique non renouvelable"])].groupby(["combustible", "codecombustible"], dropna=False).mw.agg(["size", "sum"]).round(1).to_string())
print("date samples:", d["datemiseenservice"].dropna().head(3).tolist(), d["dateraccordement"].dropna().head(3).tolist())

print("\n#### A2. SDES national rows")
for mid, shs in ((9714, ["solaire PV"]), (9715, ["éolien_terrestre", "éolien_offshore"])):
    get(f"https://www.statistiques.developpement-durable.gouv.fr/media/{mid}/download?inline", show=False, out=f"/tmp/sdes{mid}.xlsx")
    x = pd.ExcelFile(f"/tmp/sdes{mid}.xlsx")
    for sh in shs:
        df = x.parse(sh, header=None)
        print(mid, sh, df.shape)
        print(df.iloc[:, :3].to_string()[:3000])
        hdr = df.iloc[2, 3:].tolist(); sub = df.iloc[3, 3:].tolist()
        print("last header cells:", hdr[-6:], sub[-6:])
        for r in range(len(df)):
            nm = str(df.iloc[r, 0])
            if re.search(r"france|total|ensemble|métropol", nm, re.I):
                print("ROW", r, nm, df.iloc[r, -8:].tolist())

print("\n#### B. Spain alternatives")
for u in ["https://datos.gob.es/apidata/catalog/dataset?q=potencia+instalada+electrica&_pageSize=20",
          "https://datos.gob.es/apidata/catalog/dataset?q=RAIPEE&_pageSize=20",
          "https://datos.gob.es/apidata/catalog/dataset?q=registro+administrativo+instalaciones+produccion+energia+electrica&_pageSize=20"]:
    b = get(u, show=False)
    try:
        j = json.loads(b)
        for it in j["result"]["items"]:
            print("  ", it.get("_about"), "|", str(it.get("title"))[:120])
    except Exception as e:
        print("ERR", u, b[:200])
h = get("https://www.sistemaelectrico-ree.es/en/spanish-electricity-system/installed-capacity", show=False)
t = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", h, flags=re.S | re.I)
t = " ".join(unescape(re.sub(r"<[^>]+>", " ", t)).split())
i = t.find("Installed capacity", 2500)
print(t[i:i + 2500])
print("tables:", len(re.findall(r"<table", h)), "csv/xls refs:", sorted(set(re.findall(r'href="([^"]+\.(?:csv|xlsx?|json))"', h)))[:10])
print("drupal settings snippet:", re.findall(r'"(?:chart|data)[A-Za-z]*(?:Url|Path|Endpoint)"\s*:\s*"([^"]+)"', h)[:10])
print("fetch urls:", sorted(set(re.findall(r'(?:/[a-z_-]+)*/(?:api|datos|chart|data)/[A-Za-z0-9_/.?=&-]{3,100}', h)))[:20])

print("\n#### C. Terna JS")
h = get("https://dati.terna.it/en/download-center", show=False)
scripts = re.findall(r'<script[^>]+src="([^"]+)"', h)
print([s for s in scripts if "media" in s or "download" in s.lower()][:40])
inline = re.findall(r"<script(?![^>]+src)[^>]*>(.*?)</script>", h, re.S)
print("inline scripts", len(inline), [len(x) for x in inline])
for k, sc in enumerate(inline):
    if "urlSearchParams" in sc or "selectedDataset" in sc:
        for m in re.finditer(r"urlSearchParams(?:NoPaging)?\s*\(|selectedDataset\b", sc):
            pass
        i = sc.find("urlSearchParamsNoPaging")
        print("inline", k, len(sc), "idx", i)
        print(sc[max(0, i - 200):i + 3500])
        break
for s in scripts:
    if s.startswith("/") and re.search(r"download|dati|center", s, re.I):
        b = get("https://dati.terna.it" + s, show=False)
        i = b.find("urlSearchParamsNoPaging")
        print("SCRIPT", s, len(b), i)
        if i >= 0:
            print(b[max(0, i - 200):i + 3500])

print("\n#### D. PT older PDFs")
try:
    import pymupdf as fitz
except Exception:
    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "pymupdf"]); import pymupdf as fitz
for name, url in [("2024-01", "https://www.dgeg.gov.pt/media/tmfl1tla/renovaveis-202401.pdf"),
                  ("2024-02", "https://www.dgeg.gov.pt/media/mdmivv3u/dgeg-arr-2024-02.pdf"),
                  ("2023-12", "https://www.dgeg.gov.pt/media/amioxiec/dgeg-arr-2023-12.pdf"),
                  ("2021-08", "https://www.dgeg.gov.pt/media/1s2i0vpx/dgeg-arr-2021-08.pdf"),
                  ("2025-12", "https://www.dgeg.gov.pt/media/qybp4lbt/dgeg-arr-2025-12.pdf")]:
    get(url, show=False, out=f"/tmp/{name}.pdf")
    try:
        doc = fitz.open(f"/tmp/{name}.pdf")
    except Exception as e:
        print(name, "ERR", e); continue
    for pg in doc:
        t = pg.get_text()
        if re.search(r"Potência instalada\s*\nMW", t):
            lines = [x.strip() for x in t.split("\n")]
            print("\n---", name, "page", pg.number + 1)
            print(re.sub(r"\n(?=[\d >\-])", " | ", t)[:1700])
            break
    else:
        print(name, "no Potência instalada MW table; pages", len(doc))
