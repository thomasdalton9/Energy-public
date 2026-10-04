"""
Pakistan power discovery, round 12: K-Electric, the formats asia/PAKISTAN_NEPRA.py does not parse yet.
  a) every K-Electric tariff file 2021-2026 (KE tariff page) and every KE FCA/FPA admission notice (news.php)
  b) the GWh tables of the unparsed KE decisions (provision monthly FCA 2025, JUL-MAR 2023-24, Nov 2024 - Mar 2025,
     Aug 2021) and of two KE admission-notice filings (Aug 2022 FPA, Dec 2022 FCA data)
  c) NEPRA State of Industry Report 2025: K-Electric own generation / purchases lines
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
KEY = re.compile(r"K-?Electric|\bKE\b|\bKEL\b", re.I)


def get(u):
    try:
        return s.get(u, timeout=(15, 240), verify=False)
    except Exception as e:  # noqa: BLE001
        print(f"  ERR {u}: {e}", flush=True)
        return None


def text(u, pages=None):
    r = get(u)
    if r is None or r.status_code != 200:
        print(f"  {getattr(r, 'status_code', None)} {unquote(u)}")
        return []
    with pdfplumber.open(io.BytesIO(r.content)) as p:
        print(f"  {len(p.pages)} pages, {len(r.content)} b")
        return [pg.extract_text() or "" for pg in (p.pages[:pages] if pages else p.pages)]


def show(u, pat=r"GW", ctx=14, pages=None, maxhits=12):
    print(f"\n######## {unquote(u).split('.pk/')[-1]}", flush=True)
    pp = text(u, pages)
    hits = 0
    for i, t in enumerate(pp):
        lines = t.splitlines()
        idx = [j for j, l in enumerate(lines) if re.search(pat, l, re.I)]
        if not idx:
            continue
        hits += 1
        if hits > maxhits:
            break
        a, b = max(0, idx[0] - ctx), min(len(lines), idx[-1] + 4)
        print(f"  --- page {i + 1} lines {a}-{b}:")
        print("  " + "\n  ".join(lines[a:b][:70]))


def main():
    ke = {}
    for page in ("https://nepra.org.pk/tariff/Distribution%20K-Electric.php", "https://nepra.org.pk/news.php"):
        r = get(page)
        if r is None:
            continue
        for h in re.findall(r'href\s*=\s*["\']([^"\']+\.pdf)["\']', r.text, re.I):
            n = unquote(h)
            if re.search(r"/(?:KESC|K-Electric)/20(2[1-6])/", n) or (KEY.search(n.rsplit("/", 1)[-1]) and
                                                                    "admission" in n.lower() and
                                                                    re.search(r"FCA|FPA|fuel", n, re.I)):
                ke[urljoin(page, h)] = 1
    for u in sorted(ke):
        print("  " + unquote(u).split(".pk/")[-1])
    T = "https://nepra.org.pk/tariff/Tariff/K-Electric/"
    for u in [T + "2025/TRF-362%20K-ELECTRIC%20PROVISION%20MONTHLY%20FCA%2009-05-2025%205863-67.PDF",
              T + "2025/TRF-362%20K-ELECTRIC%20PROVISION%20MONTHLY%20FCA%20APRIL%202025%2009-07-2025%2010505-09.PDF",
              T + "2025/TRF-362%20KE%20FCA%20DEC-2024%2006-02-2025%203403-07.pdf",
              T + "2025/TRF-362%20K-Electric%20FCA%20January%202025%2003-04-2025%204912-16.PDF",
              T + "2024/TRF-362%20KE%20FCA%20JUL-MAR%202023-24%2006-06-2024%208448-52.PDF",
              "https://nepra.org.pk/tariff/Tariff/KESC/2021/TRF-362%20KE%20FCA%20Aug%202021%2027-10-2021%2039785-89.PDF"]:
        show(u)
    # other 2025/2026 KE tariff files not yet seen
    for u in sorted(ke):
        n = unquote(u)
        if re.search(r"/K-Electric/(2025|2026)/", n) and re.search(r"FCA|fuel|provision", n, re.I) and \
                not re.search(r"5863-67|10505-09|3403-07|4912-16|PAR-14|TRF-14", n):
            show(u, maxhits=3)
    A = "https://www.nepra.org.pk/Admission%20Notices/"
    show(A + "2022/09%20Sep/K-Electric%20FPA%20for%20the%20month%20of%20August%202022.pdf", r"GW|CPPA|NTDC|sent", pages=20)
    show(A + "2023/1%20Jan/KE%20FCA%20Data%20for%20the%20month%20of%20December%202022.PDF", r"GW|CPPA|NTDC|sent", pages=20)
    show("https://www.nepra.org.pk/publications/State%20of%20Industry%20Reports/State%20of%20Industry%20Report%202025.pdf",
         r"K-?Electric.{0,80}(GWh|generat|purchas)|(own generation|purchased from).{0,60}K-?Electric|KE.{0,10}own", ctx=6,
         maxhits=25)


if __name__ == "__main__":
    main()
