"""One-off probe 2 (5 Oct 2026): ERCOT 'Large Load Project Distribution by Type' charts in every TAC Large Load Interconnection Status Update and Operational
Overview; saves the chart images to discovery_archive/results/ercot_bytype/ for pixel/visual reading, OCRs the labels; prior CDR editions; LTLF 2025 docx;
utility 10-Q queue sentences. Log + images only."""
import io, re, os, sys, time, zipfile
import requests, fitz
from PIL import Image
import pytesseract
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36", "Accept": "*/*"}
E = "https://www.ercot.com/files/docs/"
OUTD = "discovery_archive/results/ercot_bytype"
os.makedirs(OUTD, exist_ok=True)
MON = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

def get(u, t=120):
    try:
        return requests.get(u, headers=H, timeout=t)
    except Exception as e:
        print("  FAILED", u, type(e).__name__); return None

def find(names, months, days=range(1, 29)):
    """HEAD-probe files/docs/Y/M/D/name for each (year, month) in months; return list of (label, url)."""
    out = []
    for (y, m) in months:
        hit = None
        for (ny, nm) in [(y, m), (y + (m == 12), m % 12 + 1)]:
            for d in days:
                for nmv in names:
                    u = f"{E}{ny}/{nm:02d}/{d:02d}/{nmv.format(M=MON[m-1], Y=y, m3=MON[m-1][:3])}"
                    try:
                        h = requests.head(u, headers=H, timeout=15, allow_redirects=True)
                    except Exception:
                        continue
                    if h.status_code == 200:
                        hit = u; break
                if hit: break
            if hit: break
        print(("FOUND " if hit else "MISSING ") + f"{y}-{m:02d} {hit or ''}")
        if hit: out.append((f"{y}-{m:02d}", hit))
    return out

def ocr(im):
    try: return pytesseract.image_to_string(im, config="--psm 6")
    except Exception as e: return f"OCR-ERR {e}"

TITLE = re.compile(r"distribution by type|by type|queue|status update|observations|energize", re.I)
def dig(label, u):
    r = get(u)
    if r is None or r.status_code != 200: print("  no pdf", u); return
    doc = fitz.open(stream=r.content, filetype="pdf")
    print(f"\n==== {label} {u} pages={len(doc)}")
    for i, pg in enumerate(doc):
        tx = pg.get_text()
        head = re.sub(r"\s+", " ", tx)[:160]
        if not re.search(r"large load", head, re.I) and not re.search(r"Distribution by Type|Interconnection Queue", tx, re.I): continue
        imgs = [x for x in pg.get_images(full=True)]
        print(f"--- {label} page {i+1}: {head!r} imgs={len(imgs)}")
        if re.search(r"by type", tx, re.I):
            for k, x in enumerate(imgs):
                p = fitz.Pixmap(doc, x[0])
                if p.n > 4: p = fitz.Pixmap(fitz.csRGB, p)
                if p.width < 400: continue
                fn = f"{OUTD}/{label}_p{i+1}_img{k}.png"
                p.save(fn)
                big = Image.open(fn).convert("RGB")
                big = big.resize((big.width * 2, big.height * 2))
                print(f"   saved {fn} {p.width}x{p.height}; OCR:", re.sub(r"[ \t]+", " ", ocr(big)).replace("\n", " | ")[:900])
            pix = pg.get_pixmap(dpi=110); pix.save(f"{OUTD}/{label}_p{i+1}_page.png")
        if re.search(r"Interconnection Queue", tx, re.I):
            # status-by-year table lives in the image; OCR it large
            for k, x in enumerate(imgs):
                p = fitz.Pixmap(doc, x[0])
                if p.n > 4: p = fitz.Pixmap(fitz.csRGB, p)
                if p.width < 800: continue
                big = Image.frombytes("RGB", (p.width, p.height), p.samples)
                big = big.resize((big.width * 2, big.height * 2))
                print(f"   QUEUE-TABLE {label} img{k}:", re.sub(r"[ \t]+", " ", ocr(big)).replace("\n", " | ")[:1600])

tac = find(["{M}-TAC-Report.pdf", "{M}-TAC-Report-Final.pdf", "TAC-Report-{M}-{Y}.pdf", "{M}-{Y}-TAC-Report.pdf"],
           [(2025, m) for m in range(1, 13)] + [(2026, m) for m in range(1, 10)])
for lab, u in tac: dig("TAC" + lab, u)
oo = find(["ERCOT-Monthly-Operational-Overview-{M}-{Y}.pdf"], [(2025, 1), (2025, 6), (2026, 1), (2026, 6)], range(12, 24))
for lab, u in oo: dig("OO" + lab, u)

# prior CDR editions
from openpyxl import load_workbook
for (y, m) in [(2023, 12), (2024, 5), (2024, 12), (2025, 5), (2026, 5), (2026, 6)]:
    cd = find(["CapacityDemandandReservesReport_{M}{Y}.xlsx", "CapacityDemandandReservesReport_{M}{Y}_Revised.xlsx", "CDR_{M}{Y}.xlsx", "CapacityDemandandReservesReport_{m3}{Y}.xlsx"], [(y, m)], range(1, 32))
    for lab, u in cd:
        r = get(u)
        if r is None or r.status_code != 200: continue
        try: wb = load_workbook(io.BytesIO(r.content), data_only=True)
        except Exception as e: print("parse fail", e); continue
        print("\n==== CDR", u, wb.sheetnames)
        for ws in wb.worksheets:
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                v = [(str(round(x, 1)) if isinstance(x, float) else str(x).replace("\n", " ")[:100]) for x in row if x is not None]
                s = " ; ".join(v)
                if re.search(r"data cent|crypto|hydrogen|oil & gas|oil and gas|large load|officer|contract|industrial", s, re.I) and re.search(r"\d{3}", s) and i < 200 and not re.search(r"curtail|scaled|less 25", s, re.I):
                    print(f"  [{ws.title}] r{i+1}: {s[:330]}")

# LTLF 2025 docx
r = get(E + "2025/04/08/2025_LTLF_Report.docx")
if r is not None and r.status_code == 200:
    z = zipfile.ZipFile(io.BytesIO(r.content)); x = z.read("word/document.xml").decode("utf8", "ignore")
    paras = [re.sub(r"<[^>]+>", "", p) for p in re.findall(r"<w:p[ >].*?</w:p>", x, re.S)]
    print("\n==== LTLF docx paragraphs with type keywords")
    for p in paras:
        if re.search(r"data cent|crypto|hydrogen|oil and gas|oil & gas|industrial", p, re.I) and re.search(r"\d", p): print("  *", p[:420])
print("DONE")
