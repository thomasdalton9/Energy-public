"""
Pakistan power discovery, round 13: K-Electric tables in full.
  a) NEPRA KE decisions Oct 2024 - Apr 2025 ('Company Wide Mix' sent-out by fuel / source) and Jul 2023 - Mar 2024
  b) KE admission-notice filings 2022 (Aug 2022 FPA, Dec 2022 FCA data): pages 3-12 (sent-out, purchases by source)
  c) NEPRA State of Industry Report 2025 (330 MB): K-Electric lines (pypdfium2 text, matching pages only)
"""
import io
import re
import tempfile

import pdfplumber
import pypdfium2 as pdfium
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
s = requests.Session()
s.headers.update(H)
T = "https://nepra.org.pk/tariff/Tariff/K-Electric/"
A = "https://www.nepra.org.pk/Admission%20Notices/"


def get(u):
    try:
        return s.get(u, timeout=(15, 600), verify=False)
    except Exception as e:  # noqa: BLE001
        print(f"  ERR {u}: {e}", flush=True)
        return None


def pages(u, pat, first=None, last=None, maxc=3500):
    print(f"\n######## {requests.utils.unquote(u).split('.pk/')[-1]}", flush=True)
    r = get(u)
    if r is None or r.status_code != 200:
        print(f"  status {getattr(r, 'status_code', None)}")
        return
    with pdfplumber.open(io.BytesIO(r.content)) as p:
        n = len(p.pages)
        print(f"  {n} pages")
        for i in range(first or 0, min(last or n, n)):
            t = p.pages[i].extract_text() or ""
            if pat is None or re.search(pat, t, re.I):
                print(f"  --- page {i + 1} ({len(t)} chars)")
                print("  " + t[:maxc].replace("\n", "\n  "))
            else:
                print(f"  (page {i + 1}: {len(t)} chars)")


def soir(u):
    print(f"\n######## {u}", flush=True)
    with tempfile.NamedTemporaryFile(suffix=".pdf") as f:
        with s.get(u, timeout=(15, 900), verify=False, stream=True) as r:
            for ch in r.iter_content(1 << 20):
                f.write(ch)
        f.flush()
        doc = pdfium.PdfDocument(f.name)
        print(f"  {len(doc)} pages", flush=True)
        for i in range(len(doc)):
            t = doc[i].get_textpage().get_text_range()
            if re.search(r"K-?Electric|\bKE\b", t) and re.search(r"GWh|own generation|purchas", t, re.I):
                lines = t.splitlines()
                hit = [j for j, l in enumerate(lines) if re.search(r"K-?Electric|\bKE\b", l) and
                       re.search(r"GWh|generat|purchas|sent|\d{3,}", l, re.I)]
                if hit:
                    print(f"  --- page {i + 1}")
                    for j in hit[:12]:
                        print("   " + " | ".join(x.strip() for x in lines[max(0, j - 2):j + 4])[:600])


def main():
    for f in ["2025/TRF-362%20KE%20FCA%20NOV-2024%2012-02-2025%202328-32.pdf",
              "2025/TRF-362%20KE%20FCA%20DEC-2024%2006-02-2025%203403-07.pdf",
              "2025/TRF-362%20K-Electric%20FCA%20January%202025%2028-03-2025%204866-70.PDF",
              "2025/TRF-362%20K-ELECTRIC%20PROVISION%20MONTHLY%20FCA%2009-05-2025%205863-67.PDF",
              "2025/TRF-362%20KE%20FCA%20MAR-2025%2005-06-2025%207890-94.pdf",
              "2025/TRF-362%20K-ELECTRIC%20PROVISION%20MONTHLY%20FCA%20APRIL%202025%2009-07-2025%2010505-09.PDF"]:
        pages(T + f, r"Company\s*Wide|Sent\s*out|GWh")
    pages(T + "2024/TRF-362%20KE%20FCA%20JUL-MAR%202023-24%2006-06-2024%208448-52.PDF", None, 1, 10, 3000)
    pages(A + "2022/09%20Sep/K-Electric%20FPA%20for%20the%20month%20of%20August%202022.pdf", None, 2, 12, 2500)
    pages(A + "2023/1%20Jan/KE%20FCA%20Data%20for%20the%20month%20of%20December%202022.PDF", None, 2, 12, 2500)
    soir("https://www.nepra.org.pk/publications/State%20of%20Industry%20Reports/State%20of%20Industry%20Report%202025.pdf")


if __name__ == "__main__":
    main()
