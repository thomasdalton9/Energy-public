"""
Discovery round 6 (after CAPACITY_PBUE_DISCOVERY5.py).

Round 5 found:
  Peru: POST /Portal/browser/vistadatos {baseDirectory, url, indicador, initialLink, orderFolder} returns the
        folder HTML (19.8 kB for Potencia Efectiva, 14.2 kB for Boletines) - item links not in the patterns tried.
        Chapter 2 'BD' sheet: column 'POTENCIA EFECTIVA (MW)' has some non-numeric cells (needs to_numeric).
  Bolivia: consdiesel_2025.xlsx lists one diesel plant, MOXOS.
  Ecuador: ARCONEL BNEE_<mes>_<anio>.xls (one sheet) has 'Potencia Nominal en Generacion' MW, Total and S.N.I.,
        Renovable / No renovable by type; only the latest month's .xls is linked; WP media has a PNG per month
        (abril 2024 .. junio 2026). Balance multianual 2015-2024 has energy only in its first rows.

This round: COES folder HTML in readable form; ARCONEL BNEE full layout + how older months are served.
"""
import io
import re
import sys
from urllib.parse import quote

import pandas as pd

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from CAPACITY_PBUE_DISCOVERY import get, links  # noqa: E402

pd.set_option("display.width", 300)
pd.set_option("display.max_columns", 40)
COES = "https://www.coes.org.pe/Portal/"


def coes_html(base, rel=None):
    r = get(COES + "browser/vistadatos", method="POST",
            data={"baseDirectory": base, "url": rel or base, "indicador": "", "initialLink": "", "orderFolder": ""})
    if r is None or r.status_code != 200:
        return ""
    t = re.sub(r"<script.*?</script>", "", r.text, flags=re.S)
    rows = re.findall(r"<tr.*?</tr>", t, flags=re.S)
    print(f"     {rel or base}: {len(rows)} rows")
    for row in rows[:5]:
        print("       RAW", re.sub(r"\s+", " ", row)[:700])
    for row in rows[:80]:
        cells = [re.sub(r"<[^>]+>|\s+", " ", c).strip() for c in re.findall(r"<td.*?</td>", row, flags=re.S)]
        attrs = re.findall(r"(?:onclick|href|data-\w+)=[\"']([^\"']+)[\"']", row)
        print("       ", " | ".join(c for c in cells if c)[:200], " || ", " ; ".join(attrs)[:300])
    return t


def peru():
    print("\n================ PERU")
    coes_html("Operación/Estudios/Potencia Efectiva/")
    coes_html("Publicaciones/Boletines/")
    coes_html("Publicaciones/Estadisticas Anuales/2021/")
    url = COES + "browser/download?url=" + quote("Publicaciones/Estadisticas Anuales/2025/Excel/"
                                                 "Capítulo 02_ESTADO DE LA INFRAESTRUCTURA DEL SEIN.xlsx")
    r = get(url)
    if r is not None and r.status_code == 200:
        bd = pd.read_excel(io.BytesIO(r.content), sheet_name="BD", header=3)
        mw = bd.columns[7]
        bad = bd[pd.to_numeric(bd[mw], errors="coerce").isna() & bd[mw].notna()]
        print("     non-numeric MW cells:", len(bad))
        print(bad.iloc[:10, :9].to_string())
        bd[mw] = pd.to_numeric(bd[mw], errors="coerce")
        print(bd.groupby(["TIPO DE GENERACIÓN", "TIPO DE RECURSO ENERGÉTICO"])[mw].agg(["count", "sum"]).to_string())
        print(bd.groupby("TECNOLOGÍA")[mw].agg(["count", "sum"]).to_string())
        sgi = pd.read_excel(io.BytesIO(r.content), sheet_name="SGI", header=None)
        print(sgi.iloc[:45, :12].to_string(max_colwidth=26))


def ecuador():
    print("\n================ ECUADOR")
    r = get("https://arconel.gob.ec/wp-content/uploads/downloads/2026/09/BNEE_junio_2026_revACH.xls")
    if r is not None and r.status_code == 200:
        d = pd.read_excel(io.BytesIO(r.content), header=None)
        print(d.iloc[:, :9].to_string(max_colwidth=40))
    for u in ["https://arconel.gob.ec/wp-json/", "https://arconel.gob.ec/wp-json/wp/v2/dlm_download?search=BNEE&per_page=50",
              "https://arconel.gob.ec/wp-json/download-monitor/v1/downloads?search=BNEE",
              "https://arconel.gob.ec/wp-content/uploads/downloads/2026/07/",
              "https://arconel.gob.ec/wp-json/wp/v2/search?search=BNEE&per_page=50"]:
        r = get(u)
        if r is not None and r.status_code == 200:
            txt = r.text
            if u.endswith("wp-json/"):
                ns = re.findall(r'"namespaces":\[(.*?)\]', txt)
                print("      namespaces:", ns[:1])
                print("      dlm routes:", sorted(set(re.findall(r'"(/[^"]*(?:dlm|download)[^"]*)"', txt)))[:40])
            else:
                print("     ", re.sub(r"\s+", " ", txt[:3000]))
    links("https://arconel.gob.ec/balance-nacional-de-energia-electrica/",
          pat=re.compile(r"download|BNEE|\.xls", re.I), limit=60)


if __name__ == "__main__":
    for w in sys.argv[1:] or ["peru", "ecuador"]:
        try:
            globals()[w]()
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            print(f"!! {w} failed: {type(e).__name__}: {e}")
        sys.stdout.flush()
