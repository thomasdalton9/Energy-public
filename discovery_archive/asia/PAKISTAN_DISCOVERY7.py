"""
Pakistan power discovery, round 7: why some NEPRA FCA decisions fail in asia/PAKISTAN_NEPRA.py. For each, print the
'Source Wise Generation' block of the text layer (or say there is none), and a tesseract OCR of the Annex pages.
"""
import io
import os
import re
import shutil
import sys

import pdfplumber

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "asia"))
import PAKISTAN_NEPRA as P  # noqa: E402

B = "https://nepra.org.pk/tariff/Tariff/Ex-WAPDA%20DISCOS/"
FILES = ["2022/TRF-100%20MFPA%20FCA%20Jun%202022%2012-08-2022%2015173-88.pdf",
         "2021/TRF-100%20MFPA%20XWDISCOs%20Jan%202021%2009-03-2021%2013188-04.PDF",
         "2021/TRF-100%20MFPA%20XWDISCOs%20Feb%202021%2007-04-2021%2018911-27.PDF",
         "2021/TRF-100%20MFPA%20XWDISCOs%20%20FCA%20Jun%2006-08-2021%20.PDF",
         "2023/TRF-100%20MFPA%20FCA%20EX-WAPDA%20&%20KE%20AUG%20&%20SEP%202022%2009-03-2023%204933-43.PDF",
         "2023/TRF-100%20MFPA%20MAR-2023%20XWDISCOS%20FCA%2025-05-2023%2013105-20.PDF",
         "2024/TRF-100%20MFPA%20FCA%20June%202024%20XWDISCOs%2008-08-2024%2012471-86.PDF",
         "2023/TRF-100%20XWDISCOS%20MFPA%20DEC-2023%2002-02-2024%201695-1710.PDF"]


def main():
    for f in FILES:
        u = B + f
        print(f"\n######## {f}", flush=True)
        try:
            b = P.get(u).content
        except Exception as e:  # noqa: BLE001
            print(f"  {e}")
            continue
        with pdfplumber.open(io.BytesIO(b)) as p:
            pages = [pg.extract_text() or "" for pg in p.pages]
            text = "\n".join(pages)
            hits = [m.start() for m in re.finditer(r"source\s*-?\s*wise\s*generation", text, re.I)]
            print(f"  {len(p.pages)} pages, {len(hits)} 'source wise generation' hits; annex pages: "
                  f"{[i + 1 for i, t in enumerate(pages) if re.search(r'annex', t, re.I)]}")
            for h in hits:
                print("  ---- block:\n    " + text[h:h + 1300].replace("\n", "\n    "))
            try:
                print("  parsed:", {k.strftime("%Y-%m"): (v[0], v[1]) for k, v in P.read_decision(u).items()})
            except Exception as e:  # noqa: BLE001
                print(f"  parse error {e}")
            if shutil.which("tesseract"):
                import pytesseract
                for i in range(len(p.pages) - 1, max(len(p.pages) - 6, 0), -1):
                    img = p.pages[i].to_image(resolution=300).original
                    t = pytesseract.image_to_string(img, config="--psm 6")
                    if re.search(r"source\s*-?\s*wise", t, re.I):
                        print(f"  ---- OCR page {i + 1}:\n    " + t[:2500].replace("\n", "\n    "))


if __name__ == "__main__":
    main()
