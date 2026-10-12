"""Probe 12: Italy - Terna download-centre request URL and Terna monthly report.
Probe 11 findings: Spain REData 'generacion/potencia-instalada-generacion' returns monthly national capacity by technology (16 technologies plus
'Potencia instalada total', MW) for a range of up to ~13 months per call (a 6-year range answers HTTP 400 'datos no disponibles'); Terna records
endpoint: GET with f=xlsx answers 500 and POST without the right URL 404."""
import json, re, subprocess, sys
from html import unescape
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0"

def get(u, n=800, show=True, hdr=(), out=None, post=None):
    cmd = ["curl", "-sL", "-m", "120", "-A", UA] + [x for h in hdr for x in ("-H", h)]
    if post is not None:
        cmd += ["-X", "POST", "--data", post]
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

def anchors(html, pat, lim=60):
    k = 0
    for m in re.finditer(r'<a [^>]*href="([^"]+)"[^>]*>(.*?)</a>', html, re.S | re.I):
        h, tx = m.group(1), re.sub(r"<[^>]+>", " ", m.group(2))
        tx = " ".join(unescape(tx).split())
        if re.search(pat, h + " " + tx, re.I):
            print("   ", h[:200], "|", tx[:100]); k += 1
            if k >= lim: break

print("#### A. Terna JS: request URL")
b = get("https://dati.terna.it/-/media/themes/terna/dati/dati-theme/scripts/pre-optimized-min.js", show=False)
for i in [m.start() for m in re.finditer("downloadcenter|DownloadCenter", b)]:
    print("--- ctx", i, b[max(0, i - 500):i + 500].replace("\n", " "))
i = b.find("let w=u,y=null")
print("--- u definition", [b[max(0, m.start() - 300):m.start() + 300] for m in re.finditer(r"[,;{]u=[^,;]{0,160}(?:records|api|url)", b)][:4])
for m in list(re.finditer(r"records\?|/records|\"records\"|records`", b))[:6]:
    print("--- records ctx", m.start(), b[max(0, m.start() - 400):m.start() + 400].replace("\n", " "))
print("--- 'urls' objects:", [b[m.start():m.start() + 400] for m in re.finditer(r"urls:\{", b)][:3])

print("\n#### B. Terna monthly report page")
for u in ["https://www.terna.it/it/sistema-elettrico/statistiche/rapporto-mensile", "https://www.terna.it/en/electric-system/statistics/monthly-report"]:
    h = get(u, show=False)
    print(u, len(h))
    anchors(h, r"AdactoBackend|download|\.pdf|\.xls|rapporto|monthly|2026|2025", 60)
    print("data-attrs:", sorted(set(re.findall(r'data-(?:file|href|url|download|src)[a-z-]*="([^"]{5,200})"', h)))[:30])
    print("ids:", sorted(set(re.findall(r"image=(\d+)", h)))[:60])

print("\n#### C. dati.terna.it statistical data")
for u in ["https://dati.terna.it/en/generation/statistical-data", "https://dati.terna.it/it/generazione/dati-statistici",
          "https://dati.terna.it/en/generation/statistical-data/installed-capacity"]:
    h = get(u, show=False)
    t = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", h, flags=re.S | re.I)
    t = " ".join(unescape(re.sub(r"<[^>]+>", " ", t)).split())
    print(u, len(h), t[900:2500])
    anchors(h, r"installed|capacity|potenza|download|statistic", 25)
