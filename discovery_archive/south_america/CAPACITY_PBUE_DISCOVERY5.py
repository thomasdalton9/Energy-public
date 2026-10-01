"""
Discovery round 5 (after CAPACITY_PBUE_DISCOVERY4.py).

Round 4 found:
  Peru: COES annual statistics chapter 2 (2025) has sheet 'BD' - potencia efectiva (MW) per unit with
        TIPO DE GENERACION / TECNOLOGIA / TIPO DE RECURSO ENERGETICO (total 14,388 MW); the 2021 file name differs.
        File browser: POST /Portal/browser/vistadatos {baseDirectory, url, indicador, initialLink, orderFolder}.
  Bolivia: dashboard/potencia-historial monthly labels but values change only at year boundaries (= the year's
        value, even for future months) -> effectively annual; historico/potencia gives the same annual values
        2015-2025; historico/potencia/detalle per plant (thermal not split by fuel). CNDC annual 'Consumo de
        Diesel' workbooks (consdiesel_YYYY.xlsx) list diesel-burning plants.
  Ecuador: ARCONEL (arconel.gob.ec, valid TLS) publishes the monthly 'Balance Nacional de Energia Electrica'
        (BNEE_<mes>_<anio>.xls, potencia MW + energia) and 'Balance multianual 2015-2024' (.xlsx).

This round: COES folder listings; CNDC diesel workbook; ARCONEL media listing + BNEE / multianual layouts.
"""
import re
import sys
from urllib.parse import quote

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from CAPACITY_PBUE_DISCOVERY import get  # noqa: E402
from CAPACITY_PBUE_DISCOVERY2 import dump_excel  # noqa: E402

COES = "https://www.coes.org.pe/Portal/"


def coes_list(base, rel=None):
    r = get(COES + "browser/vistadatos", method="POST",
            data={"baseDirectory": base, "url": rel or base, "indicador": "", "initialLink": "", "orderFolder": ""})
    if r is None or r.status_code != 200:
        return []
    items = re.findall(r"(?:openDirectory|downloadBlob|abrirArchivo|visualizar\w*)\(\s*'([^']+)'", r.text)
    if not items:
        items = re.findall(r'data-url=["\']([^"\']+)["\']|url=([^"\'&]+)', r.text)
        items = [a or b for a, b in items]
    head = re.sub(r"\s+", " ", r.text[:600])
    print(f"     {rel or base}: {len(items)} items; raw head: {head}")
    for it in items[:60]:
        print(f"       . {it}")
    return items


def peru():
    print("\n================ PERU")
    for base in ["Operación/Estudios/Potencia Efectiva/", "Publicaciones/Boletines/",
                 "Publicaciones/Estadisticas Anuales/2021/", "Publicaciones/Estadisticas Anuales/2023/Excel/"]:
        items = coes_list(base)
        subs = [i for i in items if i.endswith("/")]
        for s in subs[-3:]:
            coes_list(base, s)
    url = COES + "browser/download?url=" + quote("Publicaciones/Estadisticas Anuales/2025/Excel/"
                                                 "Capítulo 02_ESTADO DE LA INFRAESTRUCTURA DEL SEIN.xlsx")
    r = get(url)
    if r is not None and r.status_code == 200:
        import io
        import pandas as pd
        bd = pd.read_excel(io.BytesIO(r.content), sheet_name="BD", header=3)
        print(bd.columns.tolist())
        for c in ["TIPO DE GENERACIÓN", "TECNOLOGÍA", "TIPO DE RECURSO ENERGÉTICO"]:
            if c in bd:
                print(bd.groupby(c)[bd.columns[7]].agg(["count", "sum"]).to_string())
        print(bd.groupby(["TIPO DE GENERACIÓN", "TIPO DE RECURSO ENERGÉTICO"])[bd.columns[7]].sum().to_string())
        sgi = pd.read_excel(io.BytesIO(r.content), sheet_name="SGI", header=None)
        print(sgi.iloc[:60, :14].to_string(max_colwidth=30))


def bolivia():
    print("\n================ BOLIVIA")
    r = get("https://www.cndc.bo/wp-content/uploads/mem/estadisticas/anual/2025/consdiesel_2025.xlsx")
    if r is not None and r.status_code == 200:
        dump_excel(r.content, "consdiesel_2025", rows=40)


def ecuador():
    print("\n================ ECUADOR")
    seen = []
    for q in ["BNEE", "Balance Nacional", "multianual", "Estadistica"]:
        for page in (1, 2, 3):
            r = get("https://arconel.gob.ec/wp-json/wp/v2/media", params={"search": q, "per_page": 100, "page": page,
                                                                         "_fields": "id,date,source_url,title"})
            if r is None or r.status_code != 200:
                break
            try:
                items = r.json()
            except ValueError:
                break
            for m in items:
                if m["source_url"] not in seen:
                    seen.append(m["source_url"])
                    print(f"     media {m['date'][:10]} {m['source_url']}")
            if len(items) < 100:
                break
    r = get("https://arconel.gob.ec/publicaciones-estadistica-del-sector-electrico-2/")
    if r is not None:
        for href in sorted(set(re.findall(r'href=["\']([^"\']+\.(?:xlsx?|pdf|zip))["\']', r.text, re.I)))[:80]:
            print("      pub:", href)
    for u in ["https://arconel.gob.ec/wp-content/uploads/downloads/2026/09/BNEE_junio_2026_revACH.xls",
              "https://arconel.gob.ec/wp-content/uploads/downloads/2025/05/Balance-multianual-2015-2024.xlsx"]:
        r = get(u)
        if r is not None and r.status_code == 200:
            dump_excel(r.content, u, rows=70)


if __name__ == "__main__":
    for w in sys.argv[1:] or ["peru", "bolivia", "ecuador"]:
        try:
            globals()[w]()
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            print(f"!! {w} failed: {type(e).__name__}: {e}")
        sys.stdout.flush()
