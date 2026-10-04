"""Cambodia probe 2: EAC English annual reports (monthly tables?), salient-feature back-editions, EDC annual reports, ODC datasets."""
import io
import re
import requests
import pdfplumber
from bs4 import BeautifulSoup
from urllib.parse import urljoin, quote

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
MONTH = re.compile(r"\b(January|February|Jan|Feb|Mar|Apr|monthly|month|peak|import|purchas)", re.I)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=(20, 150), **kw)
        print(f"GET {url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}B LM={r.headers.get('last-modified')}", flush=True)
        return r
    except Exception as e:
        print(f"GET {url} -> ERR {e}", flush=True)
        return None


def links(url, pat=None):
    r = get(url)
    if not r or r.status_code != 200:
        return []
    s = BeautifulSoup(r.text, "html.parser")
    out = []
    for a in s.find_all("a", href=True):
        h = urljoin(url, a["href"])
        t = " ".join(a.get_text().split())[:90]
        if (pat is None or re.search(pat, h + " " + t, re.I)) and h not in [o[0] for o in out]:
            out.append((h, t))
            print("   LINK", h, "|", t)
    return [h for h, _ in out]


def pdf_scan(url, pages=(), keyword_pages=True, chars=3000, maxhits=12):
    r = get(url)
    if not r or r.status_code != 200 or not r.content[:5].startswith(b"%PDF"):
        return
    with pdfplumber.open(io.BytesIO(r.content)) as p:
        n = len(p.pages)
        print(f"  PDF pages={n}")
        hits = 0
        for i in range(n):
            t = p.pages[i].extract_text() or ""
            first = " | ".join(t.splitlines()[:3])[:200]
            print(f"  p{i+1}: {first}")
            show = i in pages
            if keyword_pages and hits < maxhits and len(re.findall(r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b", t)) >= 6:
                show = True
                hits += 1
            if show:
                print(f"  ===== page {i+1} =====\n" + t[:chars])


print("######## EAC English annual reports")
ls = links("https://eac.gov.kh/site/annualreport?lang=en", r"\.pdf")
en = [h for h in ls if "/uploads/" in h and h.endswith(".pdf")]
print("EN pdfs:", en[:6])
pdf_scan(en[0] if en else "https://eac.gov.kh/uploads/annual_report/english/Annual-Report-2024-en.pdf", chars=4000)

print("######## Salient feature back editions")
for y in range(2010, 2024):
    for lang in ("english", "khmer"):
        sfx = "en" if lang == "english" else "kh"
        u = f"https://eac.gov.kh/uploads/salient_feature/{lang}/salient_feature_{y}_{sfx}.pdf"
        try:
            r = requests.head(u, headers=H, timeout=(20, 60), allow_redirects=True)
            print("HEAD", u, r.status_code, r.headers.get("content-type"), r.headers.get("content-length"), r.headers.get("last-modified"))
        except Exception as e:
            print("HEAD", u, "ERR", e)
r = get("https://eac.gov.kh/uploads/salient_feature/english/salient_feature_2023_en.pdf")
if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
    with pdfplumber.open(io.BytesIO(r.content)) as p:
        print((p.pages[1].extract_text() or "")[:4000])

print("######## EDC annual reports (English)")
ls = links("https://edc.com.kh/annually_page/change/english/annuallyReport", r"\.pdf")
pdfs = [h for h in ls if h.lower().endswith(".pdf")]
for h in pdfs[:2]:
    pdf_scan(h, chars=3500, maxhits=8)

print("######## ODC datasets")
for d in ["electric-power-generation-by-type-and-location", "electric-power-installation-generation-and-consumption", "electricity-generation-plants"]:
    r = get(f"https://data.opendevelopmentcambodia.net/api/3/action/package_show?id={d}")
    if r is not None and r.status_code == 200:
        try:
            j = r.json()["result"]
            print("  ", j.get("title"), j.get("metadata_modified"))
            for res in j.get("resources", []):
                print("    RES", res.get("format"), res.get("url"), res.get("last_modified"))
        except Exception as e:
            print("  json err", e, r.text[:300])
