"""
Round 2 for Argentina gas net supply (see ARGENTINA_GAS_NETSUPPLY_DISCOVERY.py).

Round 1 found:
  - ENARGAS datos-estadisticos/GRT/GRT.xlsx "Gas Recibido por Transportistas y
    Distribuidoras de Productores y Otros Origenes" by Cuenca / Gasoducto,
    monthly, thousand m3 of 9300 kcal, 1993 to Jul-2026; "Otros Origenes" =
    LNG purchases, split by transporter since 2011.
  - datos-estadisticos/CPGNNC/CLP.xlsx: TGN/TGS fuel, GNNC (unaccounted gas) and
    losses, in % of total gas delivered by the transporters.
  - GET.xlsx (gas delivered by transporters), GETD.xlsx (total system
    deliveries by user type: the same numbers as the SE sector series).
  - daily import reports: dod-partes-exp-imp-consulta.php?tipo=importaciones
    (imp_dentro / imp_fuera return HTTP 500).
  - daily "Partes Diarios" (dod-partes-dist-trans.php, partes-diarios-listado.php)
    and daily PDF charts "Inyeccion Nacional por Gasoducto" / "Importaciones
    Diarias al Sistema de Transporte".
This prints the exact headers of GRT/CLP/GET/GETD, the daily import table
(tipo_list=importaciones) by year, the Partes Diarios form and AJAX call, the
visualisation page, and the SE Sesco gas balance zip contents.
"""
import io
import re
import zipfile

import pandas as pd
import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 180)
BASE = "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/"
DE = BASE + "datos-estadisticos/"
LIST = BASE + "partes-diarios-exp-imp-consulta-listado.php"
PAGE = BASE + "dod-partes-exp-imp-consulta.php"
SESCO_GAS = ("http://www.energia.gob.ar/contenidos/archivos/Reorganizacion/informacion_del_mercado/mercado_hidrocarburos/"
             "tablas_dinamicas/upstream/sescoweb_gas.zip")

pd.set_option("display.width", 300)
pd.set_option("display.max_columns", 40)
pd.set_option("display.max_rows", 300)
pd.set_option("display.max_colwidth", 40)

S = requests.Session()
S.headers.update(HEADERS)


def flat(s):
    return re.sub(r"\s+", " ", s)


def show_xlsx(name, sheets, head_rows=(8, 16)):
    r = S.get(DE + name, timeout=TIMEOUT)
    print(f"\n===== {name}: status={r.status_code} bytes={len(r.content)}")
    xl = pd.ExcelFile(io.BytesIO(r.content))
    for sh in sheets:
        df = xl.parse(sh, header=None)
        print(f"-- [{sh}] shape={df.shape}")
        print(df.iloc[head_rows[0]:head_rows[1]].to_string())
        dates = pd.to_datetime(df.iloc[:, 0], errors="coerce")
        recent = df[dates >= "2021-01-01"]
        print(f"   rows from 2021: {len(recent)}; first {dates[dates >= '2021-01-01'].min()} last {dates.max()}")
        print(recent.iloc[:3].to_string())
        print(recent.iloc[-14:].to_string())


def imports():
    print("\n===== daily import reports (tipo=importaciones) =====")
    r = S.get(PAGE, params={"tipo": "importaciones"}, timeout=TIMEOUT)
    html = r.text
    print(f"page bytes={len(html)}")
    for m in re.finditer(r"PD_\w+\([^)]*\)", html):
        print("   js call:", m.group(0))
    for m in re.finditer(r"<(input|select)[^>]*>", html, re.I):
        print("   form:", m.group(0)[:200])
    heads = [flat(re.sub(r"<[^>]+>", "", re.sub(r"<br\s*/?>", " | ", h))).strip()
             for h in re.findall(r"<th[^>]*>(.*?)</th>", html, re.S | re.I)]
    print("   page table heads:", heads)
    for tipo in ("importaciones", "imp", "importacion"):
        for d0, d1 in [("2021-01-01", "2021-12-30"), ("2022-01-01", "2022-12-30"), ("2023-01-01", "2023-12-30"),
                       ("2024-01-01", "2024-12-30"), ("2025-01-01", "2025-12-30"), ("2026-01-01", "2026-09-30")]:
            try:
                r = S.post(LIST, data={"fecha_desde": d0, "fecha_hasta": d1, "tipo_list": tipo}, timeout=TIMEOUT)
            except requests.RequestException as e:
                print(f"{tipo} {d0}: ERROR {e}")
                continue
            heads = [flat(re.sub(r"<[^>]+>", "", re.sub(r"<br\s*/?>", " | ", h))).strip()
                     for h in re.findall(r"<th[^>]*>(.*?)</th>", r.text, re.S | re.I)]
            rows = []
            for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", r.text, re.S | re.I):
                cells = [flat(re.sub(r"<[^>]+>", " ", c)).strip() for c in re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S | re.I)]
                if cells and re.match(r"\d{2}/\d{2}/\d{4}", cells[0]):
                    rows.append(cells)
            print(f"\n== {tipo} {d0}..{d1}: status={r.status_code} bytes={len(r.text)} rows={len(rows)} heads={heads}")
            if not rows:
                continue
            print("   first:", rows[0], " last:", rows[-1])
            df = pd.DataFrame([x[1:] for x in rows], index=pd.to_datetime([x[0] for x in rows], format="%d/%m/%Y"))
            df.columns = heads[1:len(rows[0])] if len(heads) >= len(rows[0]) else df.columns
            odd = sorted({v for v in df.values.ravel() if not re.fullmatch(r"-?[\d.]*,?\d*", str(v))})[:10]
            print("   non-numeric cells:", odd)
            df = df.apply(lambda c: pd.to_numeric(c.str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
                                                  errors="coerce"))
            print("   monthly sums (million m3) and days:")
            m = (df.resample("MS").sum() / 1000).round(1)
            m["days"] = df.resample("MS").size()
            print(m.to_string())
            print("   max daily (thousand m3):", df.max().round(0).to_dict())
        if tipo == "importaciones":
            break


def partes_diarios():
    print("\n===== Partes Diarios (dod-partes-dist-trans.php) =====")
    r = S.get(BASE + "dod-partes-dist-trans.php", timeout=TIMEOUT)
    html = r.text
    body = re.sub(r"<script.*?</script>", " ", html, flags=re.S | re.I)
    print(flat(re.sub(r"<[^>]+>", " ", body))[:2500])
    for m in re.finditer(r"<(input|select|option|a|form)\b[^>]*>([^<]{0,80})", html, re.I):
        t = m.group(0)
        if any(k in t.lower() for k in ("option", "select", "input", "onclick", "javascript", "form", "href=\"dod", "xls")):
            print("   ", flat(t)[:220])
    js = S.get(BASE + "funciones/js/funciones.js", timeout=TIMEOUT).text
    for fn in ("function PD_RefrescaListado(", "function CargarPuntosPorEmpresa", "function ExportarConsulta",
               "function PD_RefrescaListadoImportacionExportacion"):
        i = js.find(fn)
        if i >= 0:
            j = js.find("\nfunction ", i + 10)
            print(f"--- {fn}: {flat(js[i:j if j > 0 else i + 4000])[:3000]}")


def other_pages():
    for page in ("dod-visualizacion-de-datos.php", "dod-graficos-de-programacion-items.php?cat=6",
                 "dod-graficos-de-programacion-items.php?cat=7", "dod-graficos-de-programacion-items.php?cat=9",
                 "dod-transparencia-de-mercado.php", "dod-reporte-diario-sistema.php", "dod-entregas-inyecciones.php",
                 "datos-operativos.php"):
        r = S.get(BASE + page, timeout=TIMEOUT)
        print(f"\n===== {page}: status={r.status_code} bytes={len(r.text)}")
        hrefs = re.findall(r"(?:href|src)=[\"']([^\"'#]+)[\"']", r.text, re.I)
        keep = [h for h in hrefs if not re.search(r"\.(css|png|jpg|svg|ico|js)(\?|$)", h, re.I)
                and "facebook" not in h and "twitter" not in h]
        pdfs = [h for h in keep if h.lower().endswith(".pdf")]
        print(f"   {len(pdfs)} pdfs; first {pdfs[:2]} last {pdfs[-2:]}")
        print("   other links:", [h for h in keep if not h.lower().endswith(".pdf")][-60:])
        for m in re.finditer(r"<iframe[^>]*>", r.text, re.I):
            print("   iframe:", m.group(0)[:300])


def sesco():
    print("\n===== SE Sesco gas balances zip =====")
    try:
        r = S.get(SESCO_GAS, timeout=TIMEOUT)
    except requests.RequestException as e:
        print("ERROR", e)
        return
    print(f"status={r.status_code} bytes={len(r.content)} ct={r.headers.get('content-type')}")
    try:
        z = zipfile.ZipFile(io.BytesIO(r.content))
    except zipfile.BadZipFile:
        print(r.content[:300])
        return
    for info in z.infolist():
        print(f"   {info.filename} {info.file_size}")
    for info in z.infolist()[:3]:
        data = z.read(info.filename)
        if info.filename.lower().endswith((".csv", ".txt")):
            print(data[:1500].decode("latin-1"))
        elif info.filename.lower().endswith((".xlsx", ".xls", ".xlsb")):
            try:
                xl = pd.ExcelFile(io.BytesIO(data))
                print("   sheets:", xl.sheet_names)
                for sh in xl.sheet_names[:3]:
                    df = xl.parse(sh, header=None, nrows=25)
                    print(f"   [{sh}]")
                    print(df.iloc[:25, :14].to_string()[:4000])
            except Exception as e:  # noqa: BLE001
                print("   parse error", type(e).__name__, e)


def main():
    S.get(PAGE, params={"tipo": "exp_dentro"}, timeout=TIMEOUT)
    show_xlsx("GRT/GRT.xlsx", ["Cuenca", "Gasoducto"])
    show_xlsx("CPGNNC/CLP.xlsx", ["CLP"], head_rows=(4, 7))
    show_xlsx("GET/GET.xlsx", ["GET"])
    show_xlsx("GETD/GETD.xlsx", ["TipoUsuario"])
    imports()
    partes_diarios()
    other_pages()
    sesco()


if __name__ == "__main__":
    main()
