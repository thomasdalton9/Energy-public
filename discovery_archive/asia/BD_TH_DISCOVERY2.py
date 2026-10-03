"""
Bangladesh + Thailand discovery, round 2 (after BD_TH_DISCOVERY.py): open the candidate files.

Bangladesh
  - BPDB / NLDC daily report as HTML: https://misc.bpdb.gov.bd/daily-generation?date=DD-MM-YYYY - table structure,
    and how far back dates work (2019..2026)
  - BPDB power-generation-unit (unit list / capacity) and daily-max-generation pages
  - Petrobangla daily gas report (reports_type 6922d2b181fc96cef9e99f16): list size / pagination, first PDFs' text
  - Petrobangla monthly MIS report: first PDF's text
Thailand
  - EPPO gas tables T03_01_01 / T03_02_01..03 (+ -1 / -2 variants), electricity T05_01_01-1, T05_02_01, T05_03_01
  - EGAT statistics pages: embedded chart data / data links (daily peak or generation)
"""
import io
import json
import re

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*", "Accept-Language": "en-US,en;q=0.9"}
T = (20, 90)


def out(*a):
    print(*a, flush=True)


def get(u, **kw):
    try:
        r = requests.get(u, headers=H, timeout=T, verify=False, **kw)
        out(f"GET {r.url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} ERROR {type(e).__name__}: {str(e)[:150]}")
        return None


def pdf_text(content, pages=2, n=2500):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        out(f"    {len(pdf.pages)} pages")
        for i, p in enumerate(pdf.pages[:pages]):
            out(f"    --- page {i + 1}: " + (p.extract_text() or "")[:n].replace("\n", " | "))
            tbl = p.extract_tables()
            if tbl:
                out(f"    tables on page {i + 1}: {len(tbl)}; first rows: {tbl[0][:6]}")


def html_tables(r, maxrows=12):
    try:
        tables = pd.read_html(io.StringIO(r.text))
    except Exception as e:  # noqa: BLE001
        out(f"    read_html: {e!r}")
        return []
    out(f"    {len(tables)} tables")
    for i, t in enumerate(tables[:8]):
        out(f"    table {i}: shape {t.shape}; columns {list(t.columns)[:14]}")
        out(t.head(maxrows).to_string()[:2500])
    return tables


def bangladesh():
    out("\n==================== BPDB daily report (HTML)")
    r = get("https://misc.bpdb.gov.bd/daily-generation?date=30-09-2026")
    if r is not None and r.status_code == 200:
        html_tables(r, 15)
        for m in re.finditer(r".{0,120}(MMCFD|Energy Generated|Unserved|Production Cost|Total Gas).{0,160}", r.text):
            out("    ctx: " + re.sub(r"<[^>]+>|\s+", " ", m.group(0))[:280])
        forms = re.findall(r"<form[^>]*>.*?</form>", r.text, re.S)
        flat = [re.sub(r"\s+", " ", f)[:300] for f in forms[:3]]
        out(f"    forms: {flat}")
    for d in ("01-01-2019", "01-01-2020", "01-01-2021", "01-06-2022", "01-01-2023", "01-01-2024", "15-03-2025"):
        r = get(f"https://misc.bpdb.gov.bd/daily-generation?date={d}")
        if r is not None and r.status_code == 200:
            t = re.sub(r"<[^>]+>|\s+", " ", r.text)
            m = re.search(r"Energy Generated[^|]{0,60}", t)
            n_rows = len(re.findall(r"<tr", r.text))
            pdf = re.findall(r'storage/daily_entry/[^"]+', r.text)[:1]
            found = m.group(0) if m else "no Energy Generated text"
            out(f"    {d}: {n_rows} <tr>; {found}; pdf {pdf}")
    out("\n==================== BPDB power-generation-unit / daily-max-generation")
    for u in ("https://misc.bpdb.gov.bd/power-generation-unit", "https://misc.bpdb.gov.bd/daily-max-generation"):
        r = get(u)
        if r is not None and r.status_code == 200:
            html_tables(r, 20)
            out("    links: " + str(sorted(set(re.findall(r'href="([^"]+)"', r.text)))[:40]))

    out("\n==================== Petrobangla daily gas report")
    base = "https://petrobangla.org.bd/pages/reports"
    flt = json.dumps({"reports_type": "6922d2b181fc96cef9e99f16"})
    r = get(base, params={"filters": flt})
    pdfs = []
    if r is not None and r.status_code == 200:
        pdfs = re.findall(r'https://objectstorage[^"\']+\.pdf', r.text)
        out(f"    {len(pdfs)} pdf links on page 1")
        # titles / dates near each link
        for m in re.finditer(r'.{0,300}objectstorage[^"\']+\.pdf', r.text, re.S):
            out("    row: " + re.sub(r"<[^>]+>|\s+", " ", m.group(0))[-330:-150])
        for p in re.findall(r'[?&](page|p|offset|limit)=\d+', r.text)[:10]:
            out(f"    pagination param seen: {p}")
        api = sorted(set(re.findall(r'["\'](/api/[^"\']+|https?://[^"\']*api[^"\']*)["\']', r.text)))[:20]
        out(f"    api-looking urls: {api}")
        for q in ({"filters": flt, "page": 2}, {"filters": flt, "page": 10}):
            r2 = get(base, params=q)
            if r2 is not None and r2.status_code == 200:
                p2 = re.findall(r'https://objectstorage[^"\']+\.pdf', r2.text)
                out(f"    page {q['page']}: {len(p2)} pdfs; first {p2[:2]}")
    for u in pdfs[:3]:
        p = get(u)
        if p is not None and p.status_code == 200 and p.content[:4] == b"%PDF":
            pdf_text(p.content, 2, 3000)
    out("\n==================== Petrobangla monthly MIS report")
    r = get("https://petrobangla.org.bd/pages/monthly-reports")
    if r is not None and r.status_code == 200:
        mis = re.findall(r'https://objectstorage[^"\']+\.pdf', r.text)
        for u in mis[:1]:
            p = get(u)
            if p is not None and p.status_code == 200 and p.content[:4] == b"%PDF":
                pdf_text(p.content, 4, 2000)


def thailand():
    out("\n==================== EPPO tables")
    for t in ("T03_01_01", "T03_01_01-1", "T03_01_01-2", "T03_02_01", "T03_02_01-1", "T03_02_02", "T03_02_02-1",
              "T03_02_03", "T03_02_03-1", "T05_01_01-1", "T05_02_01", "T05_02_01-2", "T05_03_01"):
        r = get(f"https://www.eppo.go.th/wp-content/uploads/2026/04/{t}.xls")
        if r is None or r.status_code != 200:
            continue
        try:
            x = pd.ExcelFile(io.BytesIO(r.content))
            d = x.parse(x.sheet_names[0], header=None)
            out(f"  {t}: sheets {x.sheet_names[:5]}; shape {d.shape}")
            out(d.head(14).to_string()[:2200])
            out("  ...tail:\n" + d.dropna(how="all").tail(4).to_string()[:900])
        except Exception as e:  # noqa: BLE001
            out(f"  {t}: {e!r}")
    out("\n==================== EGAT statistics pages")
    for u in ("https://www.egat.co.th/home/statistics/", "https://www.egat.co.th/home/statistics-all-latest/",
              "https://www.egat.co.th/home/statistics-all-egat/", "https://www.egat.co.th/home/fuel/"):
        r = get(u)
        if r is None or r.status_code != 200:
            continue
        files = sorted(set(re.findall(r'https?://[^"\'\s]+\.(?:xlsx?|csv|json)', r.text)))
        out(f"    data files: {files[:20]}")
        for m in re.finditer(r'.{0,80}(series|datasets|"data"\s*:|highcharts|chart\.js|wpdatatable|tablepress|iframe).{0,200}',
                             r.text, re.I):
            out("    ctx: " + re.sub(r"\s+", " ", m.group(0))[:300])
        text = re.sub(r"<script.*?</script>|<style.*?</style>", " ", r.text, flags=re.S)
        text = re.sub(r"<[^>]+>|\s+", " ", text)
        for m in re.finditer(r".{0,100}(พีค|Peak|เมกะวัตต์|MW|ล้านหน่วย|GWh).{0,120}", text):
            out("    txt: " + m.group(0)[:240])


def main():
    bangladesh()
    thailand()


if __name__ == "__main__":
    main()
