"""One-off multi-source probe, run 4 (observed large-load history; run 1 found ERCOT/EIA/CCAF/PUCT reachable): sourced ERCOT large-load figures (GW by category) for the Load breakout tab of
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


def dump_xlsx_sheet(u, sheet_pat, r0, r1):
    r = get(u, 120)
    if r is None or r.status_code != 200:
        print("   status", None if r is None else r.status_code, u); return
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(r.content), data_only=True)
    for ws in wb.worksheets:
        if not re.search(sheet_pat, ws.title, re.I): continue
        print("  SHEET", ws.title)
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i < r0: continue
            if i > r1: break
            v = [(str(round(x, 1)) if isinstance(x, float) else str(x).replace("\n", " ")[:300]) for x in row if x is not None]
            if v: print(f"   {i+1:3d}|", " ; ".join(v)[:500])


E = "https://www.ercot.com/files/docs/"
MON = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
OBS = re.compile(r"Of the ([\d,]+) MW that have received Approval to Energize, ERCOT has observed a non-?\s*simultaneous monthly\s*peak consumption of ([\d,]+) MW in (\w+ \d{4})", re.S)
SIM = re.compile(r"simultaneous monthly peak consumption of ([\d,]+) MW in (\w+ \d{4})")
import datetime
found = {}
for yr, mi in [(2024, m) for m in range(9, 13)] + [(2025, m) for m in range(1, 13)] + [(2026, 1)]:
    name = f"{MON[mi-1]}-{yr}"
    ny, nm = (yr + 1, 1) if mi == 12 else (yr, mi + 1)
    ok = False
    for d in range(12, 24):
        for nmv in (f"ERCOT-Monthly-Operational-Overview-{name}.pdf", f"ERCOT-Monthly-Operational-Overview-Final-{name}.pdf"):
            u = f"{E}{ny}/{nm:02d}/{d:02d}/{nmv}"
            try:
                h = requests.head(u, headers=H, timeout=20, allow_redirects=True)
            except Exception:
                continue
            if h.status_code == 200:
                ok = True
                r = get(u, 120)
                from pypdf import PdfReader
                rd = PdfReader(io.BytesIO(r.content))
                txt = " ".join((p.extract_text() or "") for p in rd.pages[:20])
                t2 = re.sub(r"\s+", " ", txt)
                for m in OBS.finditer(t2):
                    print(f"OBS|{name}|{u}|approved {m.group(1)}|nonsim {m.group(2)}|{m.group(3)}")
                for m in SIM.finditer(t2):
                    print(f"SIM|{name}|simultaneous {m.group(1)}|{m.group(2)}")
                for m in re.finditer(r"(Observed Energized[^.]{0,200}\d[\d,]* MW[^.]{0,100})", t2):
                    print(f"QTXT|{name}|{m.group(1)[:300]}")
                break
        if ok: break
    if not ok: print(f"MISSING|{name}")

# MORA crypto demand response rows (latest)
dump_xlsx_sheet(E + "2026/10/02/MORA_December2026.xlsx", r"Monthly Outlook|Capacity by", 66, 130)
dump_xlsx_sheet(E + "2026/08/07/MORA_October2026.xlsx", r"Monthly Outlook", 66, 100)

# CBECI Firebase / Firestore REST (the mining-map data is loaded client-side)
for u in ["https://firestore.googleapis.com/v1/projects/ccaf-afea/databases/(default)/documents/countryProperties",
          "https://firestore.googleapis.com/v1/projects/ccaf-afea/databases/(default)/documents/mining_map",
          "https://firestore.googleapis.com/v1/projects/ccaf-afea/databases/(default)/documents/miningMapData",
          "https://ccaf.io/cbnsi/cbeci/api/countries", "https://ccaf.io/cbnsi/api/cbeci/mining_map/countries"]:
    r = get(u, 40)
    print("\n== CBECI api", u[:110], None if r is None else (r.status_code, r.text[:300].replace("\n", " ")))
r = get("https://ccaf.io/cbnsi/cbeci/mining_map")
if r is not None:
    t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
    for m in list(re.finditer(r"United States[^.]{0,200}%", t))[:5]:
        print("   CBECI text:", t[m.start(): m.end() + 100][:400])
print("\nDONE")
