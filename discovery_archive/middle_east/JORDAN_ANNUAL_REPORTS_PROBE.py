"""
Probe (Oct 2026): dump the text of every NEPCO annual report (English 2003-2024, Arabic 2024/2025) and every MEMR
annual report (English 2015-2024) so the production parser middle_east/JORDAN_NEPCO_ANNUAL.py can be written against
the real layouts. Output: probe_out/<name>.json = {"url", "last_modified", "bytes", "pages": [text per page]} plus a
few PDFs kept whole for local testing. Workflow (manual): discovery_archive/workflows/jordan_annual_reports_probe.yml.

NEPCO blocks a runner after ~65 requests (PDFs included), so this sends one GET for the list page and one per PDF.
"""
import io
import json
import os
import re
import sys
import time

import requests

requests.packages.urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "en-US,en;q=0.9"}
OUT = "probe_out"
KEEP = {"nepco_2024_en", "nepco_2025_ar", "nepco_2024_ar", "nepco_2010_en", "nepco_2005_en", "nepco_2016_en"}


def get(url, timeout=180):
    try:
        r = requests.get(url, headers=H, timeout=(15, timeout), verify=False)
        print("GET", url, r.status_code, len(r.content), r.headers.get("content-type"), r.headers.get("last-modified"), flush=True)
        return r if r.status_code == 200 else None
    except Exception as e:
        print("GET", url, "ERR", type(e).__name__, str(e)[:120], flush=True)
        return None


def dump(name, url):
    import pdfplumber
    r = get(url)
    if r is None or not r.content.startswith(b"%PDF"):
        return
    pages = []
    t0 = time.time()
    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
        for pg in pdf.pages:
            try:
                pages.append(pg.extract_text() or "")
            except Exception as e:
                pages.append(f"<<ERR {e}>>")
    json.dump({"url": url, "last_modified": r.headers.get("last-modified"), "bytes": len(r.content), "pages": pages},
              open(f"{OUT}/{name}.json", "w"), ensure_ascii=False)
    if name in KEEP or name.startswith("memr"):
        open(f"{OUT}/{name}.pdf", "wb").write(r.content)
    print(f"  {name}: {len(pages)} pages, {time.time() - t0:.0f}s", flush=True)


def main():
    os.makedirs(OUT, exist_ok=True)
    which = sys.argv[1] if len(sys.argv) > 1 else "all"
    if which in ("all", "memr"):
        r = get("https://www.memr.gov.jo/En/List/Annual_Reports", timeout=60)
        links = sorted(set(re.findall(r'href="([^"]+\.pdf)"', r.text, re.I))) if r is not None else []
        print("MEMR links", links)
        for y in range(2013, 2026):
            cand = [l for l in links if str(y) in l and "annual" in l.lower()]
            url = cand[0] if cand else f"https://www.memr.gov.jo/ebv4.0/root_storage/en/eb_list_page/annual_report_{y}.pdf"
            if url.startswith("/"):
                url = "https://www.memr.gov.jo" + url
            dump(f"memr_{y}_en", url)
    if which in ("all", "nepco"):
        r = get("https://www.nepco.com.jo/en/AnnualReports.aspx", timeout=60)
        links = sorted(set(re.findall(r'href="([^"]+\.pdf)"', r.text, re.I))) if r is not None else []
        print("NEPCO links", links)
        for y in range(2003, 2025):
            cand = [l for l in links if re.search(rf"{y}_en\.pdf$", l, re.I)]
            url = cand[0] if cand else (f"/store/docs/web/{y}_en.pdf" if y < 2019 else f"/store/DOCS/web/{y}_EN.pdf")
            url = url if url.startswith("http") else "https://www.nepco.com.jo/" + url.lstrip("./").lstrip("/")
            dump(f"nepco_{y}_en", url)
        dump("nepco_2024_ar", "https://www.nepco.com.jo/store/DOCS/web/2024_AR.pdf")
        dump("nepco_2025_ar", "https://www.nepco.com.jo/store/DOCS/web/2025_AR.pdf")


if __name__ == "__main__":
    main()
