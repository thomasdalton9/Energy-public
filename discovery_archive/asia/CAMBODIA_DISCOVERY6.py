"""Cambodia probe 6: why the 2019 / 2020 (and 2011 / 2016) EAC annual reports gave no Annex 2 values - find the pages
with pypdfium2 (fast text) and print pdfplumber text + tables there."""
import io
import re
import time
import pdfplumber
import pypdfium2 as pdfium
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
for y in [2019, 2020, 2011, 2016]:
    url = f"https://eac.gov.kh/uploads/annual_report/english/Annual-Report-{y}-en.pdf"
    b = requests.get(url, headers=H, timeout=(20, 240)).content
    t0 = time.time()
    doc = pdfium.PdfDocument(b)
    texts = [doc[i].get_textpage().get_text_range() for i in range(len(doc))]
    print(f"#### {y}: {len(doc)} pages, pdfium text in {time.time() - t0:.1f}s; chars/page "
          f"{sum(len(t) for t in texts) // max(1, len(texts))}", flush=True)
    hits = [i for i, t in enumerate(texts) if re.search(r"Annex\s*2|Generation and Import|Generation Type", t, re.I)]
    print("  pages with Annex 2 / Generation and Import:", [h + 1 for h in hits])
    for i in hits[:8]:
        print(f"  --- pdfium p{i + 1} head: {' | '.join(texts[i].splitlines()[:6])[:300]}")
    with pdfplumber.open(io.BytesIO(b)) as p:
        for i in hits[:6]:
            if i < 10:
                continue
            pg = p.pages[i]
            t = pg.extract_text() or ""
            print(f"  ===== {y} p{i + 1} (pdfplumber) =====\n{t[:3000]}")
            for tb in pg.extract_tables()[:3]:
                print("  TABLE:")
                for row in tb[:30]:
                    print("   ", row)
