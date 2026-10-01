"""
Discovery round 9 (after CAPACITY_PBUE_DISCOVERY8.py).

Round 8 found:
  Peru: COES monthly bulletin Excel (Aug 2026) has sheets 'Portada', 'Contenido', '1'..'41' - no cell matched
        'POTENCIA EFECTIVA/INSTALADA'; Jan 2021 name differs (0 bytes).
  Ecuador: no BNEE file before Apr 2024 under the guessed names. ARCONEL annual statistics are PDFs:
        id 1274 Estadistica-2025.pdf (47 MB), 439 Estadistica2024_abr.pdf, 26 (2023), 273 (2022), 275 (2021).

This round: the bulletin's table of contents and the Jan 2021 bulletin's real name; pages of ARCONEL's 2025
statistics PDF that hold capacity by type of plant (multi-year tables).
"""
import io
import re
import sys
from urllib.parse import quote

import pandas as pd

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from CAPACITY_PBUE_DISCOVERY import get  # noqa: E402
from CAPACITY_PBUE_DISCOVERY7 import coes_items  # noqa: E402

pd.set_option("display.width", 300)
pd.set_option("display.max_columns", 30)
COES = "https://www.coes.org.pe/Portal/browser/download?url="


def peru():
    print("\n================ PERU")
    r = get(COES + quote("Publicaciones/Boletines/2026/08_AGOSTO/BOLETÍN_AGOSTO 2026.xlsx"))
    xl = pd.ExcelFile(io.BytesIO(r.content))
    c = pd.read_excel(xl, sheet_name="Contenido", header=None).dropna(how="all")
    print(c.to_string(max_colwidth=90)[:6000])
    for s in xl.sheet_names[2:12]:
        d = pd.read_excel(xl, sheet_name=s, header=None).dropna(how="all").dropna(axis=1, how="all")
        print(f"     --- sheet {s} {d.shape}")
        print(d.head(12).to_string(max_colwidth=30)[:2500])
    coes_items("Publicaciones/Boletines/2021/01_ENERO/")
    coes_items("Publicaciones/Boletines/2023/06_JUNIO/")


def ecuador():
    print("\n================ ECUADOR")
    import pdfplumber
    r = get("https://arconel.gob.ec/wp-content/uploads/downloads/2026/04/Estadistica-2025.pdf")
    if r is None or r.status_code != 200:
        return
    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
        print(f"     pages: {len(pdf.pages)}")
        shown = 0
        for i, page in enumerate(pdf.pages):
            t = page.extract_text() or ""
            u = t.upper()
            if ("POTENCIA" in u and ("TURBOGAS" in u or "MCI" in u or "TIPO DE CENTRAL" in u)) and shown < 12:
                print(f"\n     ===== page {i + 1}")
                print(t[:3000])
                shown += 1


if __name__ == "__main__":
    for w in sys.argv[1:] or ["peru", "ecuador"]:
        try:
            globals()[w]()
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            print(f"!! {w} failed: {type(e).__name__}: {e}")
        sys.stdout.flush()
