"""Cambodia probe 7: the 2019 / 2020 EAC annual reports print Annex 2 rotated (pdfplumber reads it reversed) - dump
pypdfium2's text of those pages (and, for comparison, 2021's) to build a text fallback."""
import re
import pypdfium2 as pdfium
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
for y in [2019, 2020, 2021]:
    b = requests.get(f"https://eac.gov.kh/uploads/annual_report/english/Annual-Report-{y}-en.pdf", headers=H,
                     timeout=(20, 240)).content
    doc = pdfium.PdfDocument(b)
    for i in range(len(doc)):
        t = doc[i].get_textpage().get_text_range()
        if re.search(r"Annex\s*2\s*\(?[abc]\)?", " ".join(t.split()[:40]), re.I):
            print(f"===== {y} p{i + 1} rotation={doc[i].get_rotation()} =====")
            print(repr(t[:2500]))
