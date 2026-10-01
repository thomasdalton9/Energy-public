"""
Discovery round 10: layout of COES annual statistics chapter 2 for 2021 and 2022.

PERU_COES_CAPACITY.py's test run found the 2023-2025 unit lists fine (13.1 / 13.8 / 14.4 GW effective) but for
2021 and 2022 the largest 'POTENCIA EFECTIVA' + 'TIPO DE GENERACION' block gives the same hydro / wind / oil
totals in both years (2.68 GW hydro) - a different table. This dumps every sheet's title rows, every header row
that mentions POTENCIA, and the size / MW sum of each candidate block, so the right table can be picked.
"""
import io
import re
import sys
import unicodedata
from urllib.parse import quote

import pandas as pd

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from CAPACITY_PBUE_DISCOVERY import get  # noqa: E402

pd.set_option("display.width", 300)
pd.set_option("display.max_columns", 30)
FILES = {
    2021: "Publicaciones/Estadisticas Anuales/2021/Excel/Capítulo 02_Estado Actual de la Infraestructura del SEIN.xlsx",
    2022: "Publicaciones/Estadisticas Anuales/2022/Excel/Capítulo 02_ESTADO DE LA INFRAESTRUCTURA DEL SEIN.xlsx",
}


def key(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper()
    return " ".join(s.split())


for year, path in FILES.items():
    print(f"\n================ {year}")
    r = get("https://www.coes.org.pe/Portal/browser/download?url=" + quote(path))
    xl = pd.ExcelFile(io.BytesIO(r.content))
    print("sheets:", xl.sheet_names)
    for sheet in xl.sheet_names:
        raw = pd.read_excel(xl, sheet_name=sheet, header=None)
        titles = [str(v)[:120] for v in raw.iloc[:8].values.ravel() if isinstance(v, str)]
        print(f"\n--- {sheet!r} {raw.shape}; title cells: {titles[:6]}")
        for i in range(len(raw)):
            row = [key(v) for v in raw.iloc[i] if isinstance(v, str)]
            if any("POTENCIA" in c for c in row) and any(("TIPO" in c or "RECURSO" in c or "CENTRAL" in c
                                                          or "TECNOLOG" in c) for c in row):
                print(f"   header row {i}: {[str(v)[:30] for v in raw.iloc[i] if pd.notna(v)]}")
                print(raw.iloc[i + 1:i + 6, :12].to_string(max_colwidth=22))
                num = raw.iloc[i + 1:].apply(lambda c: pd.to_numeric(c, errors="coerce"))
                print("   column sums below header:", {j: round(v, 1) for j, v in num.sum().items() if v})
        if re.search(r"TOTAL|RESUMEN", " ".join(key(t) for t in titles)):
            print(raw.dropna(how="all").iloc[:40, :12].to_string(max_colwidth=24))
