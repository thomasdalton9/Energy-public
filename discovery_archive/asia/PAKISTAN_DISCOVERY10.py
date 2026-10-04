"""
Pakistan power discovery, round 10: K-Electric after the uniform FCA (from 2025?). 
  a) every K-Electric-related file on NEPRA's news / admission listings and the KE tariff page from 2024 on
  b) text of KE's 'PROVISIONAL REQUEST FOR MONTHLY FUEL COST VARIATION' (Nov 2024): own generation by plant / fuel,
     purchases by source incl. CPPA-G?
  c) the two newest KE FCA decisions (2025): full summary tables
"""
import io
import re
from urllib.parse import urljoin, unquote

import pdfplumber
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
s = requests.Session()
s.headers.update(H)
KEY = re.compile(r"K-?Electric|\bKE\b|KEL\b|KESC", re.I)


def get(u):
    try:
        return s.get(u, timeout=(15, 180), verify=False)
    except Exception as e:  # noqa: BLE001
        print(f"  ERR {u}: {e}", flush=True)
        return None


def dump(u, pat, maxpages=400, maxchars=3000):
    r = get(u)
    if r is None or r.status_code != 200:
        print(f"  {getattr(r, 'status_code', None)} {u}")
        return
    print(f"\n######## {unquote(u)[-120:]} ({len(r.content)} b)", flush=True)
    with pdfplumber.open(io.BytesIO(r.content)) as p:
        n = 0
        print(f"  {len(p.pages)} pages")
        for i, pg in enumerate(p.pages[:maxpages]):
            t = pg.extract_text() or ""
            if re.search(pat, t, re.I):
                n += 1
                if n <= 8:
                    print(f"  --- page {i + 1}:")
                    print("  " + t[:maxchars].replace("\n", "\n  "))


def main():
    files = {}
    for page in ("https://nepra.org.pk/news.php", "https://nepra.org.pk/tariff/Distribution%20K-Electric.php",
                 "https://nepra.org.pk/tariff/Petitions.php", "https://nepra.org.pk/"):
        r = get(page)
        if r is None:
            continue
        for h in re.findall(r'href\s*=\s*["\']([^"\']+\.(?:pdf|xlsx?))["\']', r.text, re.I):
            n = unquote(h)
            if KEY.search(n) and re.search(r"/(2024|2025|2026)/", n):
                files[urljoin(page, h)] = page
    for u in sorted(files):
        print("  " + unquote(u).split(".pk/")[-1])
    dump("https://nepra.org.pk/Admission%20Notices/2024/12%20Dec/PROVISIONAL%20REQUEST%20FOR%20MONTHLY%20FUEL%20COST%20VARIATION%20FOR%20NOVEMBER%202024.PDF",
         r"CPPA|NTDC|National Grid|BQPS|sent ?out|Tapal|Gul Ahmed|Lucky|SNPC|Summary", maxpages=40)
    for u in sorted(files):
        if re.search(r"/2025/|/2026/", u) and re.search(r"FCA|fuel", unquote(u), re.I) and "TRF-362" in unquote(u):
            dump(u, r"GWh", maxchars=2500)


if __name__ == "__main__":
    main()
