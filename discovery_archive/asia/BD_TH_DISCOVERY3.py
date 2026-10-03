"""
Bangladesh discovery, round 3 (after BD_TH_DISCOVERY2.py): structure needed to write the pulls.
  - BPDB / NLDC daily report HTML (misc.bpdb.gov.bd/daily-generation?date=DD-MM-YYYY): every table's header and
    first / summary rows, for a 2026 and a 2021 date; the summary block (energy generated, unserved, gas supplied)
  - BPDB power-generation-unit page tables (unit list / capacity)
  - Petrobangla daily gas report: full text of the latest PDF (all sections: production, IOC, RLNG, distribution),
    and the listing depth (pages 50 / 100 / 200 / 400: dates shown)
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


def bpdb(date):
    r = get(f"https://misc.bpdb.gov.bd/daily-generation?date={date}")
    if r is None or r.status_code != 200:
        return
    tables = pd.read_html(io.StringIO(r.text))
    out(f"  {date}: {len(tables)} tables")
    for i, t in enumerate(tables):
        out(f"  --- table {i}: shape {t.shape}")
        out("  columns: " + str([str(c)[:60] for c in t.columns][:20]))
        with pd.option_context("display.width", 250, "display.max_columns", 20):
            out(t.head(8).to_string()[:2500])
            if len(t) > 8:
                out("  ... tail:")
                out(t.tail(6).to_string()[:2000])
    text = re.sub(r"<script.*?</script>|<style.*?</style>", " ", r.text, flags=re.S)
    text = re.sub(r"<[^>]+>|\s+", " ", text)
    i = text.find("Daily Generation")
    out("  text excerpt: " + text[max(0, i - 200): i + 1500])


def petrobangla():
    base = "https://petrobangla.org.bd/pages/reports"
    flt = json.dumps({"reports_type": "6922d2b181fc96cef9e99f16"})
    r = get(base, params={"filters": flt})
    pdfs = re.findall(r'https://objectstorage[^"\']+\.pdf', r.text) if r is not None else []
    if pdfs:
        p = get(pdfs[0])
        import pdfplumber
        with pdfplumber.open(io.BytesIO(p.content)) as pdf:
            for pg in pdf.pages:
                out("  FULL TEXT:\n" + (pg.extract_text() or ""))
                for j, tb in enumerate(pg.extract_tables()):
                    out(f"  table {j}: {len(tb)} rows; " + str(tb[:3])[:600])
    for page in (50, 100, 150, 200, 300, 400):
        r2 = get(base, params={"filters": flt, "page": page})
        if r2 is None or r2.status_code != 200:
            continue
        rows = [re.sub(r"<[^>]+>|\s+", " ", m.group(0))[-200:-120]
                for m in re.finditer(r'.{0,300}objectstorage[^"\']+\.pdf', r2.text, re.S)]
        out(f"  page {page}: {len(rows)} reports; first {rows[:1]} last {rows[-1:]}")


def main():
    out("==================== BPDB daily report tables")
    bpdb("30-09-2026")
    bpdb("15-06-2021")
    out("\n==================== BPDB power-generation-unit")
    r = get("https://misc.bpdb.gov.bd/power-generation-unit")
    if r is not None and r.status_code == 200:
        for i, t in enumerate(pd.read_html(io.StringIO(r.text))):
            out(f"  table {i}: shape {t.shape}; columns {list(t.columns)[:12]}")
            out(t.head(10).to_string()[:2000])
    out("\n==================== Petrobangla daily gas report")
    petrobangla()


if __name__ == "__main__":
    main()
