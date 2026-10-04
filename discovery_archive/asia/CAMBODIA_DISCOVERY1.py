"""Cambodia electricity source probe 1: EAC salient features / annual reports, EDC reports, MME, ODC."""
import io
import re
import requests
import pdfplumber
from bs4 import BeautifulSoup
from urllib.parse import urljoin

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=(20, 120), **kw)
        print(f"GET {url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}B LM={r.headers.get('last-modified')}", flush=True)
        return r
    except Exception as e:
        print(f"GET {url} -> ERR {e}", flush=True)
        return None


def links(url, pat=None, limit=200):
    r = get(url)
    if not r or r.status_code != 200 or "html" not in (r.headers.get("content-type") or ""):
        return []
    s = BeautifulSoup(r.text, "html.parser")
    print("  title:", s.title.string if s.title else None)
    out = []
    for a in s.find_all("a", href=True):
        h = urljoin(url, a["href"])
        t = " ".join(a.get_text().split())[:90]
        if pat is None or re.search(pat, h + " " + t, re.I):
            out.append((h, t))
    seen = set()
    for h, t in out[:limit]:
        if h in seen:
            continue
        seen.add(h)
        print("   LINK", h, "|", t)
    return [h for h, _ in out]


def pdf(url, pages=None, chars=3500):
    r = get(url)
    if not r or r.status_code != 200 or not r.content[:5].startswith(b"%PDF"):
        return
    with pdfplumber.open(io.BytesIO(r.content)) as p:
        n = len(p.pages)
        print(f"  PDF pages={n}")
        for i in (pages if pages is not None else range(n)):
            if i >= n:
                break
            t = p.pages[i].extract_text() or ""
            print(f"  --- page {i+1} ---\n" + t[:chars])


print("######## EAC")
for u in ["https://eac.gov.kh/", "https://eac.gov.kh/site/annualreport", "https://www.eac.gov.kh/site/annualreport?lang=en",
          "https://eac.gov.kh/site/salientfeature", "https://eac.gov.kh/site/salient_feature"]:
    links(u, r"pdf|report|salient|statistic|annual|data|download")
for y in [2025, 2024]:
    pdf(f"https://eac.gov.kh/uploads/salient_feature/english/salient_feature_{y}_en.pdf", chars=5000)
get("https://eac.gov.kh/uploads/salient_feature/khmer/salient_feature_2025_kh.pdf")

print("######## EDC")
for u in ["https://edc.com.kh/", "https://edc.com.kh/annually_page/annuallyReport", "https://edc.com.kh/generalReport_page/generalReport"]:
    ls = links(u, r"pdf|report|upload|statistic|annual|monthly")
pdfs = [h for h in ls if h.lower().endswith(".pdf")]
print("EDC general report pdfs:", pdfs[:30])
for h in pdfs[:2]:
    pdf(h, pages=range(4))

print("######## MME")
for u in ["https://mme.gov.kh/", "https://www.mme.gov.kh/en/", "https://data.opendevelopmentcambodia.net/dataset/?q=electricity+generation",
          "https://aeds.aseanenergy.org/", "https://www.aseanenergy.org/"]:
    links(u, r"pdf|report|statistic|electric|energy|data", limit=80)
