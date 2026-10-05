"""One-off multi-source probe, run 2 (targeted; run 1 found ERCOT/EIA/CCAF/PUCT reachable): sourced ERCOT large-load figures (GW by category) for the Load breakout tab of
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


# 1. ERCOT listings: where the monthly large-load status lives
docs = []
for u in [ERC + "/committees/inactive/lfltf", ERC + "/committees/tac", ERC + "/committees/board", ERC + "/committees/llwg", ERC + "/committees/inactive/llwg",
          ERC + "/gridinfo/load/forecast", ERC + "/gridinfo/resource", ERC + "/services/rq/large-load-integration", ERC + "/about/legal/data-center-impact-rfi",
          ERC + "/committees/rpg", ERC + "/committees/rpg/", ERC + "/services/rq/rpg", ERC + "/news/mediakit", ERC + "/news/mediakit/fact-sheets", ERC + "/about/fastfacts"]:
    docs += listing(u)
docs = sorted(set(docs))
print("\nTOTAL candidate docs", len(docs))

# 2. the documents the web search surfaced + the newest status/overview files found above
named = [ERC + "/files/docs/2026/03/12/March-TAC-Report.pdf", ERC + "/files/docs/2026/03/05/February-TAC-Report.pdf",
         ERC + "/files/docs/2026/07/17/ERCOT-Monthly-Operational-Overview-June-2026.pdf",
         ERC + "/files/docs/2026/03/18/ERCOT-Monthly-Operational-Overview-February-2026.pdf",
         ERC + "/files/docs/2024/09/05/LLI%20Queue%20Status%20Update%20-%202024-9-6.pdf"]
extra = [d for d in docs if re.search(r"TAC-Report|Operational-Overview|Status-Update|LLI|Large-Load|Long-Term|LTLF|Reserves|MORA", d, re.I) and d.lower().endswith(".pdf")]
extra = sorted(extra, key=lambda d: re.search(r"/(20\d\d/\d\d/\d\d)/", d).group(1) if re.search(r"/(20\d\d/\d\d/\d\d)/", d) else "")[-14:]
for u in named + extra:
    pdf_dump(u)

# 3. xlsx files from run 1 (fixed parser)
for u in [ERC + "/files/docs/2026/06/18/Batch-Zero-Load-Information-Form-06172026.xlsx", ERC + "/files/docs/2026/06/26/Batch_Zero_Readiness_FAQs_V8.1.xlsx",
          ERC + "/files/docs/2026/09/09/BZ-Verification-RFI-Exhibit-List.xlsx", ERC + "/files/docs/2025/12/19/CapacityDemandandReservesReport_December2025.xlsx",
          ERC + "/files/docs/2025/10/06/ERCOT-Adjusted-Load-Forecast-Winter-2025-2026-for-RS-Magnitude-2025.10.07-.xlsx",
          ERC + "/files/docs/2025/04/08/ERCOT-Peak-Demand-Scenarios.xlsx", ERC + "/files/docs/2025/04/08/2025-ERCOT-Monthly-Peak-Demand-and-Energy-Forecast.xlsx"]:
    xlsx_dump(u)

# 4. Cambridge mining map: extract anything about United States / Texas from the page and its scripts
for u in ["https://ccaf.io/cbnsi/cbeci/mining_map", "https://ccaf.io/cbnsi/cbeci/mining_map/methodology"]:
    r = get(u)
    if r is None: continue
    print(f"\n== CBECI {u}")
    t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
    for m in re.finditer(r"(United States|Texas|hashrate share|download|csv)", t):
        print("   ~", t[max(0, m.start() - 120): m.end() + 200])
        break
    for l in sorted(set(re.findall(r'(?:href|src)=["\']([^"\']+)["\']', r.text))):
        if re.search(r"csv|json|xlsx|data|download|api|\.js", l, re.I): print("   asset:", abs_url(u, l)[:200])
    for m in list(re.finditer(r"\"?(United States|Texas)\"?[^{}]{0,200}\d", r.text))[:6]:
        print("   data~", r.text[m.start(): m.end() + 120][:300].replace("\n", " "))

# 5. EIA: crypto / data-centre articles, EIA-860M latest, Texas
listing("https://www.eia.gov/todayinenergy/index.php?tg=cryptocurrency", 5)
r = get("https://www.eia.gov/todayinenergy/detail.php?id=61364")
if r is not None:
    t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
    for m in re.finditer(r"Texas|ERCOT", t):
        print("   EIA61364:", t[max(0, m.start() - 150): m.end() + 250])
r = get("https://www.eia.gov/electricity/data/eia860m/")
if r is not None:
    ls = [abs_url("https://www.eia.gov/electricity/data/eia860m/", l) for l in re.findall(r'href=["\']([^"\']+\.xlsx)["\']', r.text)]
    print("EIA860M latest:", ls[-3:], "of", len(ls))
    if ls:
        rr = get(ls[-1], 120)
        if rr is not None and rr.status_code == 200:
            import pandas as pd
            x = pd.ExcelFile(io.BytesIO(rr.content))
            print("   sheets", x.sheet_names)
            df = x.parse(x.sheet_names[0], header=2)
            print("   cols", list(df.columns)[:30])
            for c in df.columns:
                if "Technology" in str(c) or "Energy Source" in str(c) or "Sector" in str(c):
                    print("   ", c, df[c].astype(str).value_counts().head(12).to_dict())
for q in ("ERCOT data center", "Texas data centers electricity demand", "Texas natural gas data centers"):
    r = get(f"https://www.eia.gov/search/?q={q.replace(' ', '+')}")
    print("EIA search", q, None if r is None else r.status_code)

# 6. PUCT SB6 / large-load filings
for u in ["https://interchange.puc.texas.gov/search/filings/?ControlNumber=58317&ItemMatch=Equal&UtilityType=A&ItemNumber=1",
          "https://www.puc.texas.gov/industry/electric/rules/sb6/", "https://www.puc.texas.gov/"]:
    r = get(u)
    if r is None: continue
    print(f"\n== PUCT {u} {r.status_code}")
    t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
    show_text(t, "puct", 12)
    for l in sorted(set(re.findall(r'href=["\']([^"\']+)["\']', r.text))):
        if re.search(r"sb.?6|large.?load|58317|data.?cent", l, re.I): print("   link:", abs_url(u, l)[:200])

# 7. LBNL / third-party
for u in ["https://eta-publications.lbl.gov/sites/default/files/2024-12/lbnl-2024-united-states-data-center-energy-usage-report.pdf",
          "https://www.gridstatus.io/insights/ercot-large-load"]:
    r = get(u, 90)
    print("\n==", u, None if r is None else (r.status_code, len(r.content)))
    if r is not None and r.status_code == 200 and u.endswith(".pdf"): pdf_dump(u, 60)
    elif r is not None and r.status_code == 200:
        show_text(re.sub(r"<[^>]+>", " ", r.text), "gridstatus", 12)
print("\nDONE")
