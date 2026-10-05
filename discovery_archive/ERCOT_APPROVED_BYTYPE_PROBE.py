"""One-off probe (5 Oct 2026): slide 'Loads Approved to Energize - By Zone & Project Type' in every ERCOT Monthly Operational Overview (Apr 2025 - Aug 2026), TAC reports Feb/Mar 2026 and the
March 2026 'Updated' pptx. Saves the slide images to discovery_archive/results/ercot_bytype/ and prints text + OCR. Log + images."""
import io, re, os, sys, zipfile, requests, pymupdf
from PIL import Image
import pytesseract
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
E = "https://www.ercot.com/files/docs/"
OUTD = "discovery_archive/results/ercot_bytype"; os.makedirs(OUTD, exist_ok=True)
OVN = "ERCOT-Monthly-Operational-Overview-"
FILES = [("2025-04", "2025/05/15/" + OVN + "April-2025.pdf"), ("2025-06", "2025/07/17/" + OVN + "June-2025.pdf"), ("2025-08", "2025/09/22/" + OVN + "August-2025.pdf"),
         ("2025-10", "2025/11/17/" + OVN + "October-2025.pdf"), ("2025-12", "2026/01/15/" + OVN + "December-2025.pdf"), ("2026-02", "2026/03/18/" + OVN + "February-2026.pdf"),
         ("2026-04", "2026/05/19/" + OVN + "April-2026.pdf"), ("2026-06", "2026/07/17/" + OVN + "June-2026.pdf"), ("2026-07", "2026/08/17/" + OVN + "July-2026.pdf"),
         ("2026-08", "2026/09/16/" + OVN + "August-2026.pdf"), ("TAC2026-02", "2026/03/05/February-TAC-Report.pdf"), ("TAC2026-03", "2026/03/12/March-TAC-Report.pdf")]
TITLE = re.compile(r"project type|by zone|approved to energize\s*[–-]\s*by|energized by type|by type", re.I)
def ocr(im):
    try: return pytesseract.image_to_string(im, config="--psm 6")
    except Exception as e: return f"OCR-ERR {e}"
def pdf(label, u):
    r = requests.get(u, headers=H, timeout=180)
    if r.status_code != 200: print("MISSING", label, r.status_code); return
    doc = pymupdf.open(stream=r.content, filetype="pdf")
    print(f"\n==== {label} {u} pages={len(doc)}")
    for i, pg in enumerate(doc):
        tx = re.sub(r"\s+", " ", pg.get_text())
        if not TITLE.search(tx[:200]): continue
        print(f"--- {label} p{i+1}: {tx[:300]!r}")
        for k, x in enumerate(pg.get_images(full=True)):
            p = pymupdf.Pixmap(doc, x[0])
            if p.n > 4: p = pymupdf.Pixmap(pymupdf.csRGB, p)
            if p.width < 500: continue
            fn = f"{OUTD}/{label}_p{i+1}_img{k}.png"; p.save(fn)
            big = Image.open(fn).convert("RGB"); big = big.resize((big.width * 2, big.height * 2))
            print(f"   saved {fn} {p.width}x{p.height}; OCR:", re.sub(r"[ \t]+", " ", ocr(big)).replace("\n", " | ")[:1200])
        if len(pg.get_images()) == 0:
            pg.get_pixmap(dpi=120).save(f"{OUTD}/{label}_p{i+1}_page.png")
for lab, f in FILES: pdf(lab, E + f)
# pptx slides 5 and 12
r = requests.get(E + "2026/03/27/March-TAC-Report-Updated_03262026.pptx", headers=H, timeout=180)
from pptx import Presentation
prs = Presentation(io.BytesIO(r.content))
for i, sl in enumerate(prs.slides, 1):
    title = " ".join(sh.text_frame.text for sh in sl.shapes if sh.has_text_frame)[:200]
    print(f"pptx slide {i}: {re.sub(chr(10), ' ', title)!r}")
    if TITLE.search(title):
        for k, sh in enumerate(sl.shapes):
            if sh.shape_type == 13:
                fn = f"{OUTD}/PPTX_s{i}_pic{k}.{sh.image.ext}"; open(fn, "wb").write(sh.image.blob)
                im = Image.open(fn).convert("RGB")
                if im.width < 500: continue
                print(f"   saved {fn} {im.width}x{im.height} OCR:", re.sub(r"[ \t]+", " ", ocr(im.resize((im.width * 2, im.height * 2)))).replace("\n", " | ")[:1500])
print("DONE")
