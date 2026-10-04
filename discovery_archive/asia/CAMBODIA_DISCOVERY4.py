"""Cambodia probe 4: print EAC annual-report Annex 2(a)/(b)/(c) pages (generation + imports by type / country) for
several editions, and the sources table of each salient-features edition."""
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


for y in [2024, 2021, 2017, 2013, 2008, 2005]:
    b = fetch(f"https://eac.gov.kh/uploads/annual_report/english/Annual-Report-{y}-en.pdf")
    if not b:
        continue
    with pdfplumber.open(io.BytesIO(b)) as p:
        for i, pg in enumerate(p.pages):
            t = pg.extract_text() or ""
            head = " ".join(t.splitlines()[:4])
            if re.search(r"Annex\s*2|Annex\s*II", head, re.I) or re.search(r"Summary Information on (Generation|Capacity)", head, re.I):
                print(f"  ===== {y} page {i+1} =====\n{t[:6000]}")
                for tb in pg.extract_tables()[:2]:
                    print("  TABLE:")
                    for row in tb[:50]:
                        print("   ", row)

for y in [2023, 2022, 2021, 2020, 2019, 2018, 2017]:
    b = fetch(f"https://eac.gov.kh/uploads/salient_feature/english/salient_feature_{y}_en.pdf")
    if not b:
        continue
    with pdfplumber.open(io.BytesIO(b)) as p:
        for i, pg in enumerate(p.pages[:4]):
            t = pg.extract_text() or ""
            if re.search(r"Import|Hydro", t):
                print(f"  ===== SF{y} page {i+1} =====\n{t[:4500]}")
