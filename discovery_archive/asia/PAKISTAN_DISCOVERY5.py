"""
Pakistan power discovery, round 5.
  a) cppa.gov.pk/downloads/xwdiscos-energy-purchase-data/<id> for other ids: older fiscal years?
  b) full Summary block of the FY 2024-25 workbook (one sheet) and its sheet list; the FY 2025-26 Q4 workbook
  c) text of the CPPA-hosted monthly PDFs (Mar 2026, Jul 2024, Aug 2026): is the Summary in the text layer?
  d) NEPRA FCA decision PDFs 2021-2024: any text-layer generation-by-fuel table (Hydel / RLNG / Nuclear lines)?
"""
import io
import re

import pandas as pd
import pdfplumber
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36"}
T = (15, 120)
s = requests.Session()
s.headers.update(H)
C = "https://cppa.gov.pk/storage/uploads/downloads/"
NT = "https://nepra.org.pk/tariff/Tariff/Ex-WAPDA%20DISCOS/"
DEC = ["2021/TRF-100%20MFPA%20XWDISCOs%20FCA%20Mar%202021%2007-05-2021%2024819-35.PDF",
       "2022/TRF-100%20MFPA%20FCA%20Jun%202022%2012-08-2022%2015173-88.pdf",
       "2023/TRF-100%20MFPA%20FCA%20XWDISCOS%20Jan-2023%2010-03-2023%204991-5006.PDF",
       "2023/TRF-100%20XWDISCOS%20FCA%20AUG-2023%2005-10-2023%2033676-92.PDF",
       "2024/TRF-100%20XWDISCOs%20FPA%20Mar-2024%2008-05-2024%206732-47.PDF",
       "2024/TRF-100%20MFPA%20FCA%20June%202024%20XWDISCOs%2008-08-2024%2012471-86.PDF"]
FUEL = re.compile(r"hydel|hydro|rlng|nuclear|coal|bagasse|wind|solar|rfo|hsd|iran|total", re.I)


def out(*a):
    print(*a, flush=True)


def get(u):
    try:
        return s.get(u, timeout=T, verify=False)
    except Exception as e:  # noqa: BLE001
        out(f"  ERR {u}: {e}")
        return None


def ids():
    for n in list(range(1, 150)) + list(range(192, 232)):
        u = f"https://cppa.gov.pk/downloads/xwdiscos-energy-purchase-data/{n}"
        q = get(u)
        if q is None or q.status_code != 200:
            continue
        t = q.text
        k = t.find('<div class="col-md-9">')
        c = t[k:]
        h3 = [re.sub(r"\s+", " ", x).strip() for x in re.findall(r"<h3[^>]*>(.*?)</h3>", c, re.S)]
        files = re.findall(r'href="(/storage/uploads/downloads/[^"]+)"[^>]*>(.*?)</a>', c, re.S)
        if any(re.search(r"energy|purchase|fca|fuel|epp|xwdisco", h + " ".join(f[1] for f in files), re.I) for h in h3 or [""]):
            names = [re.sub(r"\s+", " ", f[1]).strip() for f in files][:30]
            out(f"  id {n}: {h3[:2]} files {len(files)}: {names}")


def xlsx(name, key):
    r = get(C + key)
    x = pd.ExcelFile(io.BytesIO(r.content))
    out(f"\n==== {name}: sheets {x.sheet_names}")
    for sh in x.sheet_names[:1] + x.sheet_names[-1:]:
        df = x.parse(sh, header=None)
        m = df.astype(str).apply(lambda c: c.str.strip().str.fullmatch("Summary", case=False)).any(axis=1)
        if not m.any():
            out(f"  {sh}: no Summary")
            continue
        i0 = m[m].index[0]
        blk = df.iloc[i0 - 2:i0 + 20].dropna(how="all", axis=1)
        out(f"  == {sh} rows {i0 - 2}..{i0 + 20}")
        with pd.option_context("display.width", 400, "display.max_columns", 30, "display.max_colwidth", 22):
            out(blk.to_string())
        # fuel-type column per plant row?
        out(f"  header row: {[str(v)[:25] for v in df.iloc[4].tolist()]}")


def pdftext(name, key, nepra=False):
    r = get(key if nepra else C + key)
    if r is None or r.status_code != 200:
        out(f"  {name}: {getattr(r, 'status_code', None)}")
        return
    with pdfplumber.open(io.BytesIO(r.content)) as p:
        out(f"\n==== {name}: {len(p.pages)} pages")
        for i, pg in enumerate(p.pages):
            t = pg.extract_text() or ""
            lines = [l for l in t.splitlines() if FUEL.search(l) and re.search(r"\d", l)]
            out(f"  -- page {i + 1}: {len(t)} chars; fuel lines {len(lines)}; head {t[:120]!r}")
            if "Summary" in t:
                out("     SUMMARY: " + t[t.find("Summary"):][:2500].replace("\n", "\n     "))
            elif nepra:
                for l in lines[:25]:
                    out("     " + l[:200])


def main():
    xlsx("FY2024-25 xlsx", "f57JDYLMoYz07rq9RzrtvYzpg7Np6p3PRia5oU8P.xlsx")
    xlsx("FY2025-26 Q4 xlsx", "xslLsxQ6QJQ0Lx41ColRs4jzS19fuAtIJtIOdwke.xlsx")
    pdftext("CPPA Mar 2026", "bTVIsVBE7KKw32GTVGnoJIHtMD4sznJoDmjCV7Av.pdf")
    pdftext("CPPA Jul 2024", "CTpdnVaoCwGzmxKFRyfJI6pHwss5AUpvjBmTtH1U.pdf")
    pdftext("CPPA Aug 2026", "bIIjrVNvZQfaKwU7PIsCctEaOomiKE4uZeoaqKOi.pdf")
    for d in DEC:
        pdftext("NEPRA " + d, NT + d, nepra=True)
    ids()


if __name__ == "__main__":
    main()
