"""Probe 9: ree.es data pages (Spain), Terna request parameter code, DGEG older PDFs (pymupdf installed first).
Probe 8 findings: the REE system-report pages link to https://www.ree.es/en/datos/generation/installed-capacity?end_date=..&start_date=..&systemElectric=nacional&time_trunc=year
(sistemaelectrico-ree.es/en/datos/... is 404); Terna pre-optimized-min.js builds the query with addParam('filterSource', ...) etc."""
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

print("#### A. ree.es data pages")
for q in ["end_date=2025-12-31T23%3A59&start_date=2025-01-01T00%3A00&systemElectric=nacional&time_trunc=month",
          "end_date=2025-12-31T23%3A59&start_date=2021-01-01T00%3A00&systemElectric=nacional&time_trunc=month"]:
    h = get("https://www.ree.es/en/datos/generation/installed-capacity?" + q, show=False)
    print(len(h), q)
    t = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", h, flags=re.S | re.I)
    t = " ".join(unescape(re.sub(r"<[^>]+>", " ", t)).split())
    print("TEXT:", t[:1500])
    print("apidatos refs:", sorted(set(re.findall(r'[^"\'\s]*apidatos[^"\'\s<>]*', h)))[:10])
    print("json-ish:", [h[m.start()-100:m.start()+400].replace("\n", " ") for m in re.finditer(r"potencia-instalada|installed-capacity\.json|drupalSettings", h)][:3])
    print("xls/csv links:", sorted(set(re.findall(r'href="([^"]+(?:csv|xls|xlsx|json)[^"]*)"', h)))[:10])
    print("tables:", len(re.findall(r"<table", h)))
    break

print("\n#### B. Terna request code")
b = get("https://dati.terna.it/-/media/themes/terna/dati/dati-theme/scripts/pre-optimized-min.js", show=False)
i = b.find('addParam("filterSource"')
print(b[max(0, i - 4500):i + 2500])
j = b.find("downloadcenter")
print("downloadcenter idx", [m.start() for m in re.finditer("downloadcenter|DownloadCenter", b)][:10])

print("\n#### C. DGEG older PDFs")
subprocess.run([sys.executable, "-m", "pip", "install", "-q", "pymupdf"])
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
