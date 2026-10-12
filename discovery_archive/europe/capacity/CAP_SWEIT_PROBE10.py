"""Probe 10: Terna download-centre API (CapacityRenewableSources / CapacityThermal / GenerationPlants) and REData widget-name guesses.
Probe 9 findings: DGEG PDFs Jan 2024 onward are monthly ('Potencia Instalada (MW)', last column = report month; year-end columns before); Dec PDFs
2020-2023 are year-end; ree.es/en/datos/generation/installed-capacity-generation is a script-rendered page (no tables, no apidatos string);
Terna request = /api/sitecore/dati/downloadcenter/records with filterDataset, filterSource, filterType (Lorda/Netta) ... params."""
import json, re, subprocess, sys
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
        print("\n=====", u[:300], "|", head, flush=True)
        print(body[:n].replace("\n", " "), flush=True)
    return body

print("#### A. Terna JS rest of params and fetch code")
b = get("https://dati.terna.it/-/media/themes/terna/dati/dati-theme/scripts/pre-optimized-min.js", show=False)
i = b.find('this.addParam("filterSubcategory"')
print(b[i:i + 5200])
for m in list(re.finditer(r"downloadcenter/records|/api/sitecore/dati", b))[:6]:
    print("\n--- API ctx", m.start(), b[max(0, m.start() - 700):m.start() + 500])

print("\n#### B. Terna API tries")
API = "https://dati.terna.it/api/sitecore/dati/downloadcenter/records"
for ds, extra in [("CapacityRenewableSources", ""), ("CapacityThermal", ""), ("GenerationPlants", "")]:
    for q in [f"filterDataset={ds}&pageSize=8&pageIndex=0",
              f"filterDataset={ds}&filterYear=2025&pageSize=8&pageIndex=0",
              f"filterDataset={ds}&year=2025&month=12&pageSize=8&pageIndex=0"]:
        get(f"{API}?{q}", 1800, hdr=("Accept: application/json", "X-Requested-With: XMLHttpRequest"))

print("\n#### C. REData widget-name guesses")
H = ("Accept: application/json",)
for w in ["generacion/potencia-instalada-generacion", "generacion/installed-capacity-generation", "generacion/potencia-instalada-renovable",
          "generacion/evolucion-potencia-instalada", "generacion/potencia-instalada-ccaa", "generacion/estructura-potencia"]:
    get(f"https://apidatos.ree.es/es/datos/{w}?start_date=2025-01-01T00:00&end_date=2025-12-31T23:59&time_trunc=month", 300, hdr=H)
h = get("https://www.ree.es/en/datos/generation/installed-capacity-generation", show=False)
js = re.findall(r'<script[^>]+src="([^"]+)"', h)
print("scripts", len(js), [s for s in js if "ree" in s or s.startswith("/")][:20])
for s in js:
    if s.startswith("/") and ("custom" in s or "datos" in s or "chart" in s or "ree" in s):
        bb = get("https://www.ree.es" + s, show=False)
        hits = [bb[max(0, m.start() - 150):m.start() + 250].replace("\n", " ") for m in re.finditer(r"apidatos|potencia-instalada|/datos/", bb)][:3]
        print("JS", s, len(bb), hits)
