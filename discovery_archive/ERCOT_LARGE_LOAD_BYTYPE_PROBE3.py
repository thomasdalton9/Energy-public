"""One-off probe 3 (5 Oct 2026): older/other ERCOT 'LLI Queue Status Update' decks (by-type pie/bar) by brute-force file names; RTP / RPG / Board pages with
large-load tables by type. Saves 'by type' chart images and prints OCR."""
import io, re, os, datetime, requests, fitz
from concurrent.futures import ThreadPoolExecutor
from PIL import Image
import pytesseract
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
E = "https://www.ercot.com/files/docs/"
OUTD = "discovery_archive/results/ercot_bytype"
os.makedirs(OUTD, exist_ok=True)
MON = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

def head(u):
    try:
        return u if requests.head(u, headers=H, timeout=15, allow_redirects=True).status_code == 200 else None
    except Exception:
        return None
cands = []
d0 = datetime.date(2024, 5, 1)
while d0 <= datetime.date(2026, 9, 30):
    for off in (0, 1, 2, 3):
        f = d0 - datetime.timedelta(days=off)
        base = f"{E}{f.year}/{f.month:02d}/{f.day:02d}/"
        y, m, d = d0.year, d0.month, d0.day
        for n in (f"LLI%20Queue%20Status%20Update%20-%20{y}-{m}-{d}.pdf", f"LLI%20Queue%20Status%20Update%20-%20{y}-{m:02d}-{d:02d}.pdf",
                  f"Large%20Load%20Interconnection%20Status%20Update%20-%20{y}-{m}-{d}.pdf", f"LLI-Queue-Status-Update-{y}-{m}-{d}.pdf"):
            cands.append(base + n)
    d0 += datetime.timedelta(days=1)
print("candidates", len(cands))
with ThreadPoolExecutor(24) as ex:
    found = sorted(set(u for u in ex.map(head, cands) if u))
print("FOUND", len(found))
for u in found: print("  ", u)

def ocr(im):
    try: return pytesseract.image_to_string(im, config="--psm 6")
    except Exception as e: return f"OCR-ERR {e}"
def dig(label, u, kw=r"by type|type"):
    r = requests.get(u, headers=H, timeout=120)
    doc = fitz.open(stream=r.content, filetype="pdf")
    print(f"\n==== {label} {u} pages={len(doc)}")
    for i, pg in enumerate(doc):
        tx = pg.get_text()
        head_ = re.sub(r"\s+", " ", tx)[:150]
        if re.search(r"distribution by type|by type", tx, re.I) or re.search(r"data cent|crypto", tx, re.I):
            print(f"--- p{i+1}: {head_!r}")
            if len(tx) > 60: print("   TEXT:", re.sub(r"\s+", " ", tx)[:1500])
            for k, x in enumerate(pg.get_images(full=True)):
                p = fitz.Pixmap(doc, x[0])
                if p.n > 4: p = fitz.Pixmap(fitz.csRGB, p)
                if p.width < 500: continue
                fn = f"{OUTD}/{label}_p{i+1}_img{k}.png"; p.save(fn)
                big = Image.open(fn).convert("RGB"); big = big.resize((big.width * 2, big.height * 2))
                print(f"   saved {fn}; OCR:", re.sub(r"[ \t]+", " ", ocr(big)).replace("\n", " | ")[:700])
for n, u in enumerate(found):
    m = re.search(r"(\d{4}-\d{1,2}-\d{1,2})", u)
    dig("LLI" + (m.group(1) if m else str(n)), u)

# RTP / RPG / board pages
def links(u, pat):
    try: r = requests.get(u, headers=H, timeout=60)
    except Exception as e: print("fail", u); return []
    print("\n== LIST", u, r.status_code)
    if r.status_code != 200: return []
    from urllib.parse import urljoin
    out = sorted(set(urljoin(u, l) for l in re.findall(r'href=["\']([^"\']+)["\']', r.text) if re.search(pat, l, re.I)))
    for l in out: print("  ", l)
    return out
rtp = []
for p in ["https://www.ercot.com/gridinfo/transmission", "https://www.ercot.com/committees/rpg", "https://www.ercot.com/gridinfo/transmission/regional-transmission-plan",
          "https://www.ercot.com/gridinfo/load", "https://www.ercot.com/committees/board", "https://www.ercot.com/gridinfo/resource"]:
    rtp += links(p, r"regional.?transmission|RTP|large.?load|llis|load.?forecast")
for u in [x for x in rtp if x.lower().endswith(".pdf")][:6]:
    try:
        r = requests.get(u, headers=H, timeout=120); doc = fitz.open(stream=r.content, filetype="pdf")
    except Exception: continue
    print("\n==== RTP-ish", u, len(doc))
    for i, pg in enumerate(doc):
        tx = re.sub(r"\s+", " ", pg.get_text())
        if re.search(r"data cent", tx, re.I) and re.search(r"crypto", tx, re.I) and re.search(r"\d{3}", tx):
            print(f"--- p{i+1}: {tx[:1200]}")
print("DONE")
