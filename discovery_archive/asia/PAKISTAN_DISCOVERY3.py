"""
Pakistan power discovery, round 3.
  a) cppa.gov.pk XWDISCOs Energy Purchase Data / Fuel Adjustment pages: the content block (file links, Excel?)
  b) monitoring.cppa.gov.pk front page
  c) OCR (tesseract) of the "Summary" fuel table in the scanned CPPA-G energy purchase data PDFs (2021-2026 Mar),
     whose text layer is missing or garbage
"""
import io
import re
import shutil
import subprocess

import pdfplumber
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36"}
T = (15, 90)
s = requests.Session()
s.headers.update(H)
AN = "https://nepra.org.pk/Admission%20Notices/"
PDFS = [
    "2026/04%20Apr/XWDISCOS%20ENERGY%20PURCHASE%20DATA%20FOR%20THE%20MONTH%20OF%20MARCH%202026.PDF",
    "2025/10%20Oct/XWDISCOs%20ENERGY%20PURCHASE%20DATA%20FOR%20MONTH%20OF%20SEPTEMBER%202025.PDF",
    "2024/02%20Feb/XWDISCOS%20ENERGY%20PURCHASE%20DATA%20FOR%20THE%20MONTH%20OF%20January%202024.PDF",
    "2023/03%20Mar/FCA%20Data%20of%20XWDISCOs%20for%20the%20month%20of%20February%202023.PDF",
    "2022/04%20Apr/FCA%20Data%20of%20XWDISCOs%20for%20the%20month%20of%20March%202022.pdf",
]


def out(*a):
    print(*a, flush=True)


def get(u):
    try:
        return s.get(u, timeout=T, verify=False)
    except Exception as e:  # noqa: BLE001
        out(f"  ERR {u}: {e}")
        return None


def cppa():
    for u in ("https://cppa.gov.pk/xwdiscos-energy-purchase-data", "https://cppa.gov.pk/fuel-adjustment-notifications",
              "http://monitoring.cppa.gov.pk/"):
        r = get(u)
        if r is None:
            continue
        t = r.text
        k = t.rfind("</nav>")
        f = t.find("<footer")
        body = t[k if k > 0 else 0:f if f > k else len(t)]
        out(f"\n==== {u} {r.status_code} total {len(t)}b, content {len(body)}b")
        out(re.sub(r"\n\s*\n+", "\n", body)[:9000])
        hrefs = sorted(set(re.findall(r'(?:href|src|data-[a-z-]+|action|onclick)\s*=\s*["\']([^"\']+)["\']', t)))
        out(f"  all attrs ({len(hrefs)}): {[h for h in hrefs if 'cdn-cgi' not in h][:200]}")


def ocr_pdf(u):
    out(f"\n######## OCR {u}")
    r = get(AN + u)
    if r is None or r.status_code != 200:
        out(f"  status {getattr(r, 'status_code', None)}")
        return
    import pytesseract
    with pdfplumber.open(io.BytesIO(r.content)) as p:
        out(f"  {len(p.pages)} pages")
        for i, pg in enumerate(p.pages):
            if i < 3:
                continue
            img = pg.to_image(resolution=300).original
            t = pytesseract.image_to_string(img, config="--psm 6")
            hit = re.search(r"Summary|Coal.?Local|Grand Total", t, re.I)
            out(f"  --- page {i + 1} {img.size} ocr {len(t)} chars hit={bool(hit)}")
            if re.search(r"Summary", t, re.I) or re.search(r"Coal.?Local", t, re.I):
                out(t[:6000])
            else:
                out(t[:400])


def main():
    cppa()
    out(f"tesseract: {shutil.which('tesseract')}")
    if shutil.which("tesseract"):
        out(subprocess.run(["tesseract", "--version"], capture_output=True, text=True).stdout[:200])
        for u in PDFS:
            try:
                ocr_pdf(u)
            except Exception as e:  # noqa: BLE001
                out(f"  OCR error {e}")


if __name__ == "__main__":
    main()
