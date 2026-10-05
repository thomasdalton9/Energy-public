"""One-off multi-source probe, run 3 (deep dump; run 1 found ERCOT/EIA/CCAF/PUCT reachable): sourced ERCOT large-load figures (GW by category) for the Load breakout tab of
americas/TEXAS_DEMAND_REGRESSION.py. Crawls ERCOT/PUCT/EIA/CBECI/LBNL/Comptroller pages, lists data/report links, downloads the
promising ones (pdf/xlsx/csv) and prints sentences / rows with MW/GW near large-load keywords. Output is the log only."""
import io, re, sys, json
import requests

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36", "Accept": "*/*"}
KEY = re.compile(r"large.?load|crypto|bitcoin|data.?cent|LLIS|controllable load|\bCLR\b|\bLFL\b|flexible load|hydrogen|batch zero|interconnection (status|queue)|co-?located|behind.the.meter", re.I)
NUM = re.compile(r"\b\d[\d,.]*\s?(GW|MW|gigawatt|megawatt|TWh|EH/s)", re.I)
LINKKEY = re.compile(r"large.?load|llis|crypto|bitcoin|data.?cent|load.?forecast|ltlf|batch|flexible|controllable|clr|cdr|sb.?6|58317|interconnection", re.I)
seen_files = set()


def get(u, t=45):
    try:
        r = requests.get(u, headers=H, timeout=t, allow_redirects=True)
        return r
    except Exception as e:  # noqa
        print(f"   FAILED {u}: {type(e).__name__}: {str(e)[:120]}")
        return None


def abs_url(base, l):
    from urllib.parse import urljoin
    return urljoin(base, l)


def show_text(t, label, n=14):
    t = re.sub(r"\s+", " ", t)
    hits = [s.strip() for s in re.split(r"(?<=[.;])\s", t) if KEY.search(s) and NUM.search(s)]
    for s in hits[:n]:
        print(f"   * [{label}] {s[:380]}")
    return len(hits)


def read_file(u, r):
    ct = r.headers.get("content-type", "")
    b = r.content
    print(f"   FILE {u[:150]} {len(b)} bytes {ct}")
    try:
        if u.lower().split("?")[0].endswith(".pdf") or b[:4] == b"%PDF":
            from pypdf import PdfReader
            rd = PdfReader(io.BytesIO(b))
            print(f"     pdf pages {len(rd.pages)}")
            for i, p in enumerate(rd.pages[:60]):
                tx = p.extract_text() or ""
                if KEY.search(tx) and NUM.search(tx):
                    print(f"     --- page {i+1}")
                    for ln in tx.split("\n"):
                        if KEY.search(ln) or NUM.search(ln):
                            print("       |", ln.strip()[:200])
        elif b[:2] == b"PK":
            from openpyxl import load_workbook
            wb = load_workbook(io.BytesIO(b), read_only=True, data_only=True)
            for ws in wb.worksheets[:12]:
                print(f"     sheet {ws.title}")
                for i, row in enumerate(ws.iter_rows(values_only=True)):
                    if i > 25: break
                    v = [str(x)[:25] for x in row if x is not None]
                    if v: print("       |", " ; ".join(v)[:220])
        else:
            tx = b.decode("utf8", "ignore")
            print("     head:", tx[:600].replace("\n", " | "))
    except Exception as e:  # noqa
        print(f"     parse failed {type(e).__name__}: {e}")


def crawl(u, depth_files=6, follow_pages=0):
    print(f"\n== {u}")
    r = get(u)
    if r is None: return
    print(f"   status {r.status_code} {len(r.content)} bytes {r.headers.get('content-type')}")
    if r.status_code != 200: return
    ct = r.headers.get("content-type", "")
    if "html" not in ct and "json" not in ct and "text" not in ct:
        read_file(u, r); return
    show_text(re.sub(r"<[^>]+>", " ", r.text), "page")
    links = sorted(set(re.findall(r'href=["\']([^"\']+)["\']', r.text)))
    cand = [abs_url(u, l) for l in links if (LINKKEY.search(l) or re.search(r"\.(xlsx?|csv|pdf|zip)(\?|$)", l, re.I))]
    docs = [c for c in cand if re.search(r"\.(xlsx?|csv|pdf)(\?|$)", c, re.I)]
    pages = [c for c in cand if c not in docs]
    for c in (docs + pages)[:40]:
        print("   link:", c[:200])
    n = 0
    for c in docs:
        if c in seen_files or n >= depth_files: continue
        seen_files.add(c); n += 1
        rr = get(c, 90)
        if rr is not None and rr.status_code == 200: read_file(c, rr)
        elif rr is not None: print("   file status", rr.status_code, c[:120])
    for c in pages[:follow_pages]:
        crawl(c, 3, 0)


ERC = "https://www.ercot.com"
PAT = re.compile(r"overview|tac-?report|lli|large|status|batch|ltlf|long.?term|load.?forecast|cdr|adjust|llwg|lfl|crypto|data.?cent|rfi|mora|reserves", re.I)


def listing(u, maxn=80):
    r = get(u)
    print(f"\n== LIST {u}")
    if r is None or r.status_code != 200:
        print("   status", None if r is None else r.status_code); return []
    t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
    show_text(t, "page", 30)
    out = []
    for l in sorted(set(re.findall(r'href=["\']([^"\']+)["\']', r.text))):
        a = abs_url(u, l)
        if "/files/docs/" in a and PAT.search(a):
            out.append(a)
    for a in out[-maxn:]:
        print("   doc:", a[:200])
    return out


def pdf_dump(u, maxpages=80, ctx=1):
    r = get(u, 120)
    if r is None or r.status_code != 200:
        print("   status", None if r is None else r.status_code, u); return
    from pypdf import PdfReader
    rd = PdfReader(io.BytesIO(r.content))
    print(f"\n== PDF {u} ({len(rd.pages)} pages)")
    for i, p in enumerate(rd.pages[:maxpages]):
        tx = p.extract_text() or ""
        if re.search(r"large.?load|crypto|data.?cent|flexible load|controllable|approved to energize|observed", tx, re.I) and NUM.search(tx):
            print(f"  --- p{i+1}")
            for ln in tx.split("\n"):
                if ln.strip(): print("    |", ln.strip()[:220])


def xlsx_dump(u, kw=KEY, maxrows=400):
    r = get(u, 120)
    if r is None or r.status_code != 200:
        print("   status", None if r is None else r.status_code, u); return
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(r.content), data_only=True)
    print(f"\n== XLSX {u}: sheets {wb.sheetnames}")
    for ws in wb.worksheets[:15]:
        print("  sheet", ws.title, ws.max_row, "x", ws.max_column)
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i > maxrows: break
            v = [str(x)[:40] for x in row if x is not None]
            if v and (i < 6 or kw.search(" ".join(v))):
                print("    |", " ; ".join(v)[:260])


def dump_xlsx_sheet(u, sheet_pat, maxrows=130, width=34):
    r = get(u, 120)
    if r is None or r.status_code != 200:
        print("   status", None if r is None else r.status_code, u); return
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(r.content), data_only=True)
    print(f"\n== XLSX {u}: sheets {wb.sheetnames}")
    for ws in wb.worksheets:
        if not re.search(sheet_pat, ws.title, re.I): continue
        print("  SHEET", ws.title)
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i > maxrows: break
            v = [(str(round(x, 1)) if isinstance(x, float) else str(x).replace("\n", " ")[:300]) for x in row if x is not None]
            if v: print(f"   {i+1:3d}|", " ; ".join(v)[:700])


def pdf_pages(u, pages=None, pat=None):
    r = get(u, 120)
    if r is None or r.status_code != 200:
        print("   status", None if r is None else r.status_code, u); return
    from pypdf import PdfReader
    rd = PdfReader(io.BytesIO(r.content))
    print(f"\n== PDF {u} ({len(rd.pages)} pages)")
    for i, p in enumerate(rd.pages):
        tx = p.extract_text() or ""
        if (pages and (i + 1) in pages) or (pat and re.search(pat, tx, re.I)):
            print(f"  --- p{i+1}")
            for ln in tx.split("\n"):
                if ln.strip(): print("    |", ln.strip()[:240])


E = "https://www.ercot.com/files/docs/"
# A. CDR Dec 2025: large load table by type + load resources / CLR
dump_xlsx_sheet(E + "2025/12/19/CapacityDemandandReservesReport_December2025.xlsx", r"Findings|Demand|Load-Resource|Large|Scenario", 125)
# B. latest MORA: crypto demand response and load resources rows
dump_xlsx_sheet(E + "2026/10/02/MORA_December2026.xlsx", r".", 70)
# C. newest operational overview (Aug 2026) and the March 2026 TAC report: queue pages in full
pdf_pages(E + "2026/09/16/ERCOT-Monthly-Operational-Overview-August-2026.pdf", pat=r"large load|queue|crypto|data cent|load resource|controllable")
pdf_pages(E + "2026/03/12/March-TAC-Report.pdf", pages={2, 3, 4, 5, 8, 9, 10})
# D. 2025 LTLF report (docx): large-load categories in the forecast adjustments
r = get(E + "2025/04/08/2025_LTLF_Report.docx", 120)
if r is not None and r.status_code == 200:
    import zipfile
    z = zipfile.ZipFile(io.BytesIO(r.content))
    x = z.read("word/document.xml").decode("utf8", "ignore")
    paras = [re.sub(r"<[^>]+>", "", p) for p in re.split(r"</w:p>", x)]
    print("\n== LTLF docx paragraphs", len(paras))
    for t in paras:
        t = t.strip()
        if t and re.search(r"large load|crypto|data cent|hydrogen|industrial|officer|adjust|TSP", t, re.I) and re.search(r"\d", t):
            print("   *", t[:500])
# E. CBECI: Texas / US numbers inside the JS bundles
for u in ["https://ccaf.io/cbnsi/js/us-states.js", "https://ccaf.io/cbnsi/js/countries.js"]:
    r = get(u)
    print("\n== CBECI js", u, None if r is None else (r.status_code, len(r.text)))
    if r is not None and r.status_code == 200:
        for m in list(re.finditer(r"Texas", r.text))[:3]:
            print("   ", r.text[max(0, m.start() - 100): m.end() + 200].replace("\n", " "))
r = get("https://ccaf.io/cbnsi/cbeci/mining_map")
if r is not None:
    for js in sorted(set(re.findall(r'src=["\']([^"\']*_nuxt[^"\']+\.js)["\']', r.text))):
        j = get(abs_url("https://ccaf.io/cbnsi/cbeci/mining_map", js), 60)
        if j is None or j.status_code != 200: continue
        for m in list(re.finditer(r"(storage\.googleapis|firestore|firebaseio|\.csv|mining_map[a-z_/]*(data|json))", j.text))[:4]:
            print("   jsref", js[-14:], j.text[max(0, m.start() - 80): m.end() + 120].replace("\n", " "))
# F. PUCT, LBNL, GridStatus (run 2 crashed before these)
for u in ["https://interchange.puc.texas.gov/search/filings/?ControlNumber=58317&ItemMatch=Equal&UtilityType=A&ItemNumber=1",
          "https://www.puc.texas.gov/"]:
    r = get(u)
    if r is None: continue
    print(f"\n== PUCT {u} {r.status_code}")
    t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
    show_text(t, "puct", 12)
    for l in sorted(set(re.findall(r'href=["\']([^"\']+)["\']', r.text))):
        if re.search(r"sb.?6|large.?load|58317|data.?cent", l, re.I): print("   link:", abs_url(u, l)[:200])
r = get("https://www.gridstatus.io/insights/ercot-large-load", 60)
if r is not None and r.status_code == 200:
    print("\n== gridstatus"); print(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))[:2500])
pdf_pages("https://eta-publications.lbl.gov/sites/default/files/2024-12/lbnl-2024-united-states-data-center-energy-usage-report.pdf", pat=r"Texas|ERCOT")
print("\nDONE")
