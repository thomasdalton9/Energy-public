"""Probe 8: Spain REE system-report data endpoint (/en/datos/generation/installed-capacity on sistemaelectrico-ree.es), Terna download-centre
request parameters (pre-optimized-min.js), DGEG older PDF layouts.
Probe 7 findings: SDES quarterly Total France rows exist (solar 33,964.8 MW and onshore wind 24,387 MW, offshore 2,316 MW at 30 Jun 2026);
ODRE register splits Hydraulique 'Pompage turbinage' (pumped, 5.06 GW), wind onshore/offshore ('Terrestre', 'En mer pose/flottant'), thermal fuels
(GAZ, CHARB, FIOUL), batteries ('Stockage non hydraulique'); sistemaelectrico-ree.es pages call /en/datos/generation/installed-capacity?..."""
import json, re, subprocess, sys
from html import unescape
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
        print("\n=====", u[:260], "|", head, flush=True)
        print(body[:n].replace("\n", " "), flush=True)
    return body

print("#### A. Spain REE system-report data endpoint")
h = get("https://www.sistemaelectrico-ree.es/en/spanish-electricity-system/installed-capacity", show=False)
for m in list(re.finditer(r"/en/datos/generation/installed-capacity", h))[:4]:
    print("CTX:", h[max(0, m.start() - 300):m.start() + 700].replace("\n", " "))
base = "https://www.sistemaelectrico-ree.es"
H = ("Accept: application/json",)
for q in ["start_date=2025-01-01T00:00&end_date=2025-12-31T23:59&time_trunc=month",
          "start_date=2021-01-01T00:00&end_date=2025-12-31T23:59&time_trunc=month",
          "start_date=2021-01-01T00:00&end_date=2025-12-31T23:59&time_trunc=year",
          "start_date=2025-01-01T00:00&end_date=2025-12-31T23:59&time_trunc=month&all_ccaa=allCcaa",
          "end_date=2025-12-31T23:59&time_trunc=month&start_date=2025-01-01T00:00&geo_trunc=electric_system&geo_limit=peninsular&geo_ids=8741"]:
    for prefix in ("/en/datos/generation/installed-capacity?", "/es/datos/generacion/potencia-instalada?"):
        get(base + prefix + q, 1500, hdr=H)

print("\n#### B. Terna JS")
b = get("https://dati.terna.it/-/media/themes/terna/dati/dati-theme/scripts/pre-optimized-min.js", show=False)
print(len(b))
for kw in ["urlSearchParams(){", "CapacityRenewableSources"]:
    ms = [m.start() for m in re.finditer(re.escape(kw), b)]
    print(kw, ms[:12])
i = b.find("urlSearchParams(){")
print(b[max(0, i - 100):i + 5000])
for i in [m.start() for m in re.finditer("CapacityRenewableSources", b)][:6]:
    print("\n--- CRS ctx", i, b[max(0, i - 300):i + 700])

print("\n#### C. DGEG older PDFs")
import pymupdf as fitz
for name, url in [("2024-01", "https://www.dgeg.gov.pt/media/tmfl1tla/renovaveis-202401.pdf"),
                  ("2024-02", "https://www.dgeg.gov.pt/media/mdmivv3u/dgeg-arr-2024-02.pdf"),
                  ("2023-12", "https://www.dgeg.gov.pt/media/amioxiec/dgeg-arr-2023-12.pdf"),
                  ("2021-12", "https://www.dgeg.gov.pt/media/5fjpfuk1/dgeg-arr-2021-12_v2.pdf")]:
    get(url, show=False, out=f"/tmp/{name}.pdf")
    try:
        doc = fitz.open(f"/tmp/{name}.pdf")
    except Exception as e:
        print(name, "ERR", e); continue
    hit = False
    for pg in doc:
        t = pg.get_text()
        if re.search(r"pot[eê]ncia\s+instalada", t, re.I) and re.search(r"Hídrica", t) and re.search(r"Eólica", t) and re.search(r"MW", t):
            print("\n---", name, "page", pg.number + 1, "of", len(doc))
            print(re.sub(r"\n(?=[\d >\-])", " | ", t)[:1500])
            hit = True
            break
    if not hit:
        print(name, "no table page; pages", len(doc))
