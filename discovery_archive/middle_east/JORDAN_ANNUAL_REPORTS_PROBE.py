"""
Probe (Oct 2026): dump the text of every NEPCO annual report (English 2003-2024, Arabic 2024/2025) and every MEMR
annual report (English 2015-2024) so the production parser middle_east/JORDAN_NEPCO_ANNUAL.py can be written against
the real layouts. Output: probe_out/<name>.json = {"url", "last_modified", "bytes", "pages": [text per page]} plus a
few PDFs kept whole for local testing. Workflow (manual): discovery_archive/workflows/jordan_annual_reports_probe.yml.

NEPCO blocks a runner after ~65 requests (PDFs included), so this sends one GET for the list page and one per PDF.
The workflow committed the text to discovery_archive/results/jordan_annual/report_text.json.gz (+ the MEMR list pages);
GitHub artifact downloads are blocked from the sandbox.

RESULTS (12 Oct 2026):
 * NEPCO English 2003-2024 all download (Last-Modified 2004-2025); 2007, 2009 and 2016 are IMAGE-ONLY (no text layer);
   2016 needed a retry (connection cut mid-download). The AnnualReports.aspx page links no PDFs (script-built), the URL
   pattern works. Arabic 2024 (29 May 2025) and 2025 (12 Aug 2026) download; 2025_EN is 404.
 * Table titles move (table numbers change almost every few years): matched by title text. Purchased by type of Fuel
   (2018 reports on, 4-5 years), Production by Type of Fuel (2008-2017, transposed: 'Total 2017 a b c d e' rows; the
   2017 report prints the numbers on the line BEFORE the year), Fuel Consumption (by type of fuel) for Electricity
   Generation (T.T.O.E), Peak load for Interconnected System (dates from the 2020 reports; the 2019 report labels the
   columns Summer/Winter, later ones Morning/Evening), older 'System Peak Loads', Purchases / Sales tables (imports,
   exports by counterparty). The 2019 report's 'by type of Generation' table prints its labels one row out of line
   (Wind = combined cycle), so wind/solar/hydro are kept only where they add up to renewables. The 2010 sales header
   is printed in reverse year order, the data are chronological.
 * Arabic 2025: the text layer is usable - numbers in visual order (newest year first, % column first), labels
   reversed; every table checked against the English reports on the shared years agrees (fuel 25/25, purchases 40/40,
   peak 29/30, sales 18/20: the 2024 Jericho / Iraq-transmission export figures are swapped relative to the English
   2024 report).
 * MEMR: list page https://www.memr.gov.jo/En/List/Annual_Reports links the English files 2015-2024 (2020 =
   en_memr_annual_report_2020_-_23.2.2022.pdf, 2021 = annual_report_2021_en.pdf, 2023 = annual_report_2023-4.pdf;
   unknown paths answer the home page with HTTP 200). Gas for power is in the prose: 2016 LNG 122,676 'Mcf' (= MMcf) /
   335 per day, 2017 126,785 MMscf / 347, 2019 345 MMcf/d, 2020 337, 2021 340, 2022 340 (and 344 in another
   section), 2023 115,000 / 315, 2024 119,456 / 327; Risha 2016 11.2, 2017 9.8, 2019 19, 2022 14.7 MMcf/d. 2015 and
   2018 give no gas-for-power figure (2013-2015 give Egyptian pipeline imports in mcm, 2016-2018 LNG imports).
Production script: middle_east/JORDAN_NEPCO_ANNUAL.py.
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


def get(url, timeout=180, tries=1):
    for t in range(tries):
        r = _get(url, timeout)
        if r is not None:
            return r
    return None


def _get(url, timeout):
    try:
        r = requests.get(url, headers=H, timeout=(15, timeout), verify=False)
        print("GET", url, r.status_code, len(r.content), r.headers.get("content-type"), r.headers.get("last-modified"), flush=True)
        return r if r.status_code == 200 else None
    except Exception as e:
        print("GET", url, "ERR", type(e).__name__, str(e)[:120], flush=True)
        return None


def dump(name, url):
    import pdfplumber
    r = get(url, tries=3)
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
        links = []
        for k, lu in enumerate(["https://www.memr.gov.jo/En/List/Annual_Reports",
                                "https://www.memr.gov.jo/Ar/List/التقارير_السنوية"]):
            r = get(lu, timeout=60)
            if r is None:
                continue
            open(f"{OUT}/memr_list_{k}.html", "w").write(r.text)
            links += [l for l in re.findall(r'[^"\'\s<>()]+\.pdf', r.text, re.I) if "/en/" in l.lower()]
        links = sorted(set(links))
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


def pack():
    import gzip
    allt = {}
    old = "discovery_archive/results/jordan_annual/report_text.json.gz"
    if os.path.exists(old):
        allt = json.load(gzip.open(old, "rt"))
    for fn in sorted(os.listdir(OUT)):
        if fn.endswith(".json"):
            allt[fn[:-5]] = json.load(open(f"{OUT}/{fn}"))
    os.makedirs("discovery_archive/results/jordan_annual", exist_ok=True)
    with gzip.open("discovery_archive/results/jordan_annual/report_text.json.gz", "wt") as f:
        json.dump(allt, f, ensure_ascii=False)
    for fn in os.listdir(OUT):
        if fn.endswith(".html"):
            open(f"discovery_archive/results/jordan_annual/{fn}", "w").write(open(f"{OUT}/{fn}").read())


if __name__ == "__main__":
    main()
    pack()
