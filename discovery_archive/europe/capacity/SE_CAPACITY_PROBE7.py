"""
Probe 7 (SE Europe capacity). Output: results/se_probe7.txt
EPDK monthly report list structure: HTML around some titles, what each linked download is (magic bytes, filename), and the
installed-capacity table text in the older (2016-2018) PDF layout.
"""
import os, re, sys, io, html
import requests
from urllib.parse import urljoin
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36", "Accept-Language": "tr,en;q=0.8"}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results"); os.makedirs(OUT, exist_ok=True)
LOG = open(os.path.join(OUT, "se_probe7.txt"), "w")
def p(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); LOG.write(s + "\n"); LOG.flush()
import fitz
U = "https://www.epdk.gov.tr/Detay/Icerik/3-0-23-3/elektrikaylik-sektor-raporlar"
t = requests.get(U, headers=H, timeout=60).text
p("html size", len(t))
for title in ("2026 Yılı Elektrik Piyasası Ocak Ayı", "2025 Yılı Elektrik Piyasası Ocak Ayı", "2021 Yılı Elektrik Piyasası Ocak Ayı", "2019 Yılı Elektrik Piyasası Ocak Ayı"):
    i = t.find(title)
    p("=====", title, i)
    if i >= 0: p(re.sub(r"\s+", " ", t[i - 200:i + 1800]))
# every distinct a-tag href containing Download or pdf/xls for 2025 Jan region
j = t.find("2025 Yılı Elektrik Piyasası Ocak Ayı")
seg = t[j:j + 2500]
links = re.findall(r'href="([^"]+)"', seg)
p("links near 2025 Ocak:", links[:12])
def peek(u):
    r = requests.get(urljoin("https://www.epdk.gov.tr/", html.unescape(u)), headers=H, timeout=90)
    p("   ", u[:90], r.status_code, len(r.content), r.headers.get("content-disposition"), r.headers.get("content-type"), r.content[:8])
    return r
for u in links[:6]:
    if "Download" in u or re.search(r"\.(pdf|xls|doc)", u, re.I):
        try: peek(u)
        except Exception as e: p("ERR", type(e).__name__, str(e)[:80])
# older layout
named = []
for m in re.finditer(r'href="([^"]*DownloadDocument\?id=[^"]+)"', t):
    pre = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", t[max(0, m.start() - 700):m.start()])))
    mm = re.findall(r"(\d{4}) Yılı Elektrik Piyasası (\w+) Ayı Sektör Raporu", pre)
    named.append((mm[-1] if mm else None, m.group(1)))
want = [("2018", "Ocak"), ("2017", "Aralık"), ("2016", "Haziran"), ("2023", "Haziran"), ("2021", "Haziran")]
for w in want:
    for nm, u in named:
        if nm == w:
            try:
                r = requests.get(urljoin("https://www.epdk.gov.tr/", html.unescape(u)), headers=H, timeout=90)
            except Exception as e:
                p("ERR", w, e); continue
            p("-----", w, r.status_code, len(r.content), r.headers.get("content-disposition"), r.content[:6])
            if r.content[:4] == b"%PDF":
                d = fitz.open(stream=r.content, filetype="pdf"); shown = 0
                for k in range(len(d)):
                    tx = d[k].get_text()
                    if re.search(r"(?i)kurulu\s*g[üu]ç", tx) and re.search(r"(?i)(do[ğg]al\s*gaz)", tx) and re.search(r"(?i)(linyit|hidrolik)", tx):
                        p(f"  [page {k+1}]", re.sub(r"[ \t]+", " ", tx)[:2500].replace("\n", " | ")); shown += 1
                        if shown >= 3: break
            break
