"""
Round 3 for Argentina gas net supply (see rounds 1-2).

Round 2 found:
  - GRT.xlsx 'Cuenca' columns: TGN Neuquina | TGN Noroeste | TGN Otros Origenes |
    TGS Neuquina | TGS San Jorge | TGS Austral | TGS Otros Origenes |
    Distribuidoras Propios | Distribuidoras Otros Origenes | Total general
    (thousand m3 9300 kcal, monthly 1993..Jul-2026). TGN 'Otros Origenes' (LNG)
    is 0 in 2026 although Escobar regasified ~475 million m3 in Jun-2026.
  - Daily import reports: POST tipo_list=importaciones -> Fecha, Bolivia,
    GNL B. Blanca, GNL Escobar, Gasandes, Norandino, Total (thousand m3/day), 2021+.
  - CLP.xlsx: fuel / GNNC / losses as % of gas delivered, per transporter (no volumes).
  - SE "Balances de Gas - desde 2009 - Sesco Web" zip: one xlsx pivot table
    (concepto: Aventado, Consumo en Yacimiento, Entregado a Generadores / Industrias /
    Licenciatarios de Distribucion / Otros Productores / TGN / TGS, Exportaciones
    Directas, ...), to Aug-2026.
This round looks for a DAILY injection series (Partes de Distribucion y Transporte,
Flujos de gas estimados, the daily 'Inyeccion Nacional por Gasoducto' PDFs) and
reads the SESCO pivot cache (national monthly totals per concepto, 2021+).
"""
import io
import re
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd
import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 240)
BASE = "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/"
SESCO_GAS = ("http://www.energia.gob.ar/contenidos/archivos/Reorganizacion/informacion_del_mercado/mercado_hidrocarburos/"
             "tablas_dinamicas/upstream/sescoweb_gas.zip")
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

pd.set_option("display.width", 300)
pd.set_option("display.max_columns", 40)
pd.set_option("display.max_rows", 300)

S = requests.Session()
S.headers.update(HEADERS)


def flat(s):
    return re.sub(r"\s+", " ", s)


def partes():
    print("===== Partes de Distribucion y Transporte (partes-diarios-listado.php) =====")
    S.get(BASE + "dod-partes-dist-trans.php", timeout=TIMEOUT)
    for d0, d1 in (("20260901", "20260905"), ("2026-09-01", "2026-09-05"), ("20210601", "20210603")):
        r = S.post(BASE + "partes-diarios-listado.php", data={"fecha_desde": d0, "fecha_hasta": d1}, timeout=TIMEOUT)
        print(f"\n-- {d0}..{d1}: status={r.status_code} bytes={len(r.text)}")
        print(flat(re.sub(r"<[^>]+>", " ", r.text))[:1500])
        links = re.findall(r"(?:href|src|onclick)=[\"']([^\"']+)[\"']", r.text, re.I)
        print("   links:", links[:30])
        if links:
            first = next((l for l in links if re.search(r"\.(pdf|xlsx?|csv)|descarga|ObtenerArchivo", l, re.I)), None)
            if first:
                u = first if first.startswith("http") else BASE + first.lstrip("/")
                try:
                    f = S.get(u, timeout=TIMEOUT)
                    print(f"   sample {u}: status={f.status_code} ct={f.headers.get('content-type')} bytes={len(f.content)} "
                          f"head={f.content[:120]!r}")
                    pdf_text(f.content)
                except requests.RequestException as e:
                    print("   sample ERROR", e)


def pdf_text(content, n=3000):
    if not content.startswith(b"%PDF"):
        return
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            print(f"   pdf pages={len(pdf.pages)}")
            print("   text:", flat(pdf.pages[0].extract_text() or "")[:n])
    except Exception as e:  # noqa: BLE001
        print("   pdf parse error", type(e).__name__, e)


def flujos():
    for op in ("historico", "2025", "2026"):
        r = S.get(BASE + "dod-flujos-gas-estimados.php", params={"opcion": op}, timeout=TIMEOUT)
        print(f"\n===== flujos-gas-estimados {op}: status={r.status_code} bytes={len(r.text)}")
        body = re.sub(r"<script.*?</script>|<style.*?</style>", " ", r.text, flags=re.S | re.I)
        print(flat(re.sub(r"<[^>]+>", " ", body))[:800])
        links = [l for l in re.findall(r"(?:href|src)=[\"']([^\"'#]+)[\"']", r.text, re.I)
                 if not re.search(r"\.(css|png|jpg|svg|ico|js)(\?|$)", l, re.I)]
        print("   links:", links[-25:])
        for m in re.finditer(r"<iframe[^>]*>", r.text, re.I):
            print("   iframe:", m.group(0)[:300])


def daily_pdfs():
    print("\n===== daily PDF charts =====")
    for path in ("datos-operativos-despacho/graficos-programacion/6/ING_20260826.pdf",
                 "datos-operativos-despacho/graficos-programacion/7/IMP_20260826.pdf",
                 "datos-operativos-despacho/graficos-programacion/8/DDS_20260826.pdf",
                 "datos-operativos-despacho/graficos-programacion/9/PEI_202607.pdf"):
        r = S.get(BASE + path, timeout=TIMEOUT)
        print(f"\n-- {path}: status={r.status_code} bytes={len(r.content)}")
        pdf_text(r.content, 2500)
    # how far back do the daily ING pdfs go (same naming)?
    for d in ("20210615", "20230615", "20250615", "20260615", "20260901", "20260928"):
        r = S.get(BASE + f"datos-operativos-despacho/graficos-programacion/6/ING_{d}.pdf", timeout=TIMEOUT)
        print(f"   ING_{d}.pdf: status={r.status_code} pdf={r.content[:4] == b'%PDF'}")


def sesco():
    print("\n===== SE Sesco gas balance pivot cache =====")
    r = S.get(SESCO_GAS, timeout=TIMEOUT)
    z = zipfile.ZipFile(io.BytesIO(r.content))
    name = z.namelist()[0]
    x = zipfile.ZipFile(io.BytesIO(z.read(name)))
    parts = [n for n in x.namelist() if "pivotCache" in n]
    print("pivot parts:", parts)
    defs = ET.fromstring(x.read("xl/pivotCache/pivotCacheDefinition1.xml"))
    fields = []
    for cf in defs.iter("{%s}cacheField" % NS["m"]):
        items = [it.get("v") for it in cf.find("m:sharedItems", NS) or []]
        fields.append((cf.get("name"), items))
        print(f"  field {cf.get('name')!r}: {len(items)} shared items, e.g. {items[:12]}")
    recs = ET.fromstring(x.read("xl/pivotCache/pivotCacheRecords1.xml"))
    rows = []
    for rec in recs:
        row = []
        for i, v in enumerate(rec):
            tag = v.tag.split("}")[1]
            if tag == "x":
                row.append(fields[i][1][int(v.get("v"))])
            elif tag == "m":
                row.append(None)
            else:
                row.append(v.get("v"))
        rows.append(row)
    df = pd.DataFrame(rows, columns=[f[0] for f in fields])
    print(f"records: {len(df)}")
    print(df.head(5).to_string())
    cols = {c.lower(): c for c in df.columns}
    year = next((cols[c] for c in cols if c.startswith("a") and "o" in c and len(c) <= 4), None)
    print("columns:", list(df.columns), "year col:", year)
    val = next((c for c in df.columns if "cantidad" in c.lower()), df.columns[-1])
    df[val] = pd.to_numeric(df[val], errors="coerce")
    keys = [c for c in df.columns if c.lower() in ("año", "anio", "ano", "mes", "concepto")]
    print("group keys:", keys)
    if len(keys) == 3:
        g = df.groupby(keys)[val].sum().unstack(keys[2])
        g = g[g.index.get_level_values(0).astype(str) >= "2021"]
        print((g / 1000).round(1).to_string())


def main():
    partes()
    flujos()
    daily_pdfs()
    sesco()


if __name__ == "__main__":
    main()
