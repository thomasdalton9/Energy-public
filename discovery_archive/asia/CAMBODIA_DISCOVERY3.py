"""Cambodia probe 3: layout of EAC annual-report Annex 2(a)/(b)/(c) (generation + imports by type/country) across
editions, chapter 4.1 text, peak-demand mentions; identify EDC's hashed annual-report PDFs."""
import io
import re
import requests
import pdfplumber

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def fetch(url):
    try:
        r = requests.get(url, headers=H, timeout=(20, 200))
        print(f"GET {url} -> {r.status_code} {len(r.content)}B LM={r.headers.get('last-modified')}", flush=True)
        return r.content if r.status_code == 200 and r.content[:4] == b"%PDF" else None
    except Exception as e:
        print("ERR", url, e)


for y in [2024, 2023, 2019, 2014, 2010]:
    b = fetch(f"https://eac.gov.kh/uploads/annual_report/english/Annual-Report-{y}-en.pdf")
    if not b:
        continue
    with pdfplumber.open(io.BytesIO(b)) as p:
        print(f"#### {y}: {len(p.pages)} pages")
        shown = 0
        for i, pg in enumerate(p.pages):
            t = pg.extract_text() or ""
            for ln in t.splitlines():
                if re.search(r"peak|maximum demand", ln, re.I):
                    print(f"   [p{i+1} peak] {ln[:200]}")
            if re.search(r"Annex 2\s*\(?[abc]\)?|Generation and Import|4\.1\s+Generation|Energy Available|Summary Information on Generation", t) and shown < 8:
                shown += 1
                print(f"  ===== {y} page {i+1} =====\n{t[:5000]}")
                if re.search(r"Annex 2", t):
                    for tb in pg.extract_tables()[:3]:
                        print("  TABLE:")
                        for row in tb[:60]:
                            print("   ", row)

print("######## EDC hashed reports")
for h in ["def65040b39aa223f54ff3d577b336e3", "a4969d005440f67a4e57ebf240c48d85", "976446fa42bb6ef28accc18451d43e9e",
          "3750351feb4f2a77fa4199d913920486", "ca13382b6fdbc63e265e3699fb8739b4", "249dd26cfb9304ad8ed0c8c3cb8a4989",
          "32ecb115dc1cbc1135e0ea108367e584", "c106dc23b073a675eef20265f90a51ad", "3bbde440527426a39bdb2cee7cf3c25c",
          "e99a8260ade17f47c36e254329b97af4"]:
    b = fetch(f"https://admin.edc.com.kh/images/annuallyreport/{h}.pdf")
    if not b:
        continue
    with pdfplumber.open(io.BytesIO(b)) as p:
        txt = [(pg.extract_text() or "") for pg in p.pages[:60]]
        print(f"  pages={len(p.pages)} first: {' | '.join(' '.join(t.split())[:150] for t in txt[:4])}")
        for i, t in enumerate(txt):
            if len(re.findall(r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b", t)) >= 8 or re.search(r"purchas|import", t, re.I):
                print(f"  --- p{i+1} ---\n{t[:2500]}")
b = fetch("https://admin.edc.com.kh/images/annuallyreport/Annual Report 2017_en.pdf")
if b:
    with pdfplumber.open(io.BytesIO(b)) as p:
        for i, pg in enumerate(p.pages):
            t = pg.extract_text() or ""
            if len(re.findall(r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b", t)) >= 8 or re.search(r"energy purchas|import", t, re.I):
                print(f"  --- EDC2017 p{i+1} ---\n{t[:2500]}")
