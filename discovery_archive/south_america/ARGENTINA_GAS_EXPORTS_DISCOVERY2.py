"""
Round 2 for Argentina gas exports by destination. Round 1
(ARGENTINA_GAS_EXPORTS_DISCOVERY.py) found two candidate raw sources:

  1. ENARGAS "Gas Exportado Dentro del Sistema de Transporte"
     https://www.enargas.gob.ar/secciones/transporte-y-distribucion/datos-estadisticos/Expo/Exportaciones.xlsx
     monthly since 1997-02, columns by transporter (TGN/TGS) > country >
     export point (Uruguayana, GasAndes, Norandino, Petrouruguay, EGS Chile,
     Methanex YPF Chile, Gas Link Uruguay) + Total general.
  2. Secretaria de Energia "Refinacion y Comercializacion ... (Tablas
     Dinamicas)" CSV importaciones-exportaciones-a-partir-del-2016-.csv:
     company x product x country, monthly to 2026-08, product
     "Gas Natural(miles/m3)".

This prints the full ENARGAS layout (header + units + 2021 onwards) and
aggregates the SE CSV's natural gas exports by country per month, then
lists the dataset's resources and looks at the ENARGAS daily export
"partes" pages (dentro / fuera del sistema).
"""
import io
import re

import pandas as pd
import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 120)
ENARGAS_XLSX = "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/datos-estadisticos/Expo/Exportaciones.xlsx"
SE_CSV = ("http://datos.energia.gob.ar/dataset/5bdc436c-60d4-4c86-98ab-59834d047700/resource/"
          "ea145b70-c78f-4c94-8df0-27e3acfda3fc/download/importaciones-exportaciones-a-partir-del-2016-.csv")
CKAN_SHOW = "http://datos.energia.gob.ar/api/3/action/package_show"
PARTES = "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/dod-partes-exp-imp-consulta.php"

pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 40)


def enargas():
    print("===== ENARGAS Exportaciones.xlsx =====")
    r = requests.get(ENARGAS_XLSX, headers=HEADERS, timeout=TIMEOUT)
    print(f"status={r.status_code} bytes={len(r.content)} last-modified={r.headers.get('last-modified')}")
    xl = pd.ExcelFile(io.BytesIO(r.content))
    idx = xl.parse("Indice", header=None)
    for v in idx.values.ravel():
        if isinstance(v, str):
            print(f"  Indice: {v}")
    df = xl.parse("Exportaciones", header=None)
    for i in range(0, 14):
        vals = [str(v) for v in df.iloc[i].tolist() if pd.notna(v)]
        if vals:
            print(f"  row {i}: {vals}")
    body = df.iloc[14:].copy()
    body[0] = pd.to_datetime(body[0], errors="coerce")
    body = body.dropna(subset=[0])
    print(f"  date range {body[0].min()} .. {body[0].max()}, rows={len(body)}")
    print(body[body[0] >= "2021-01-01"].to_string(index=False))
    tail = df.iloc[14:].copy()
    tail = tail[pd.to_datetime(tail[0], errors="coerce").isna()].dropna(how="all")
    print("  non-date rows after header:")
    print(tail.to_string()[:2000])


def se_csv():
    print("\n===== SE importaciones-exportaciones-a-partir-del-2016 =====")
    r = requests.get(SE_CSV, headers=HEADERS, timeout=TIMEOUT, stream=True)
    print(f"status={r.status_code} content-length={r.headers.get('content-length')} "
          f"last-modified={r.headers.get('last-modified')}")
    raw = r.content
    print(f"bytes={len(raw)}")
    df = pd.read_csv(io.BytesIO(raw), encoding="utf-8-sig", low_memory=False)
    print(f"rows={len(df)} cols={list(df.columns)}")
    print("tipodecomercializacion:", df["tipodecomercializacion"].value_counts().to_dict())
    gas = df[df["producto"].astype(str).str.contains("Gas Natural", case=False)]
    print("gas productos:", gas["producto"].value_counts().to_dict())
    print("gas units:", gas["unidad"].value_counts().to_dict())
    exp = gas[(gas["tipodecomercializacion"] == "Exportación")
              & (gas["producto"].astype(str).str.startswith("Gas Natural(")) & (gas["anio"] >= 2021)]
    print("subtipos:", exp["subtipodecomercializacion"].value_counts().to_dict())
    print("paises (nonzero):", exp[exp["cantidad"] != 0]["pais"].value_counts().to_dict())
    print("empresas (nonzero):", exp[exp["cantidad"] != 0]["empresa"].value_counts().head(30).to_dict())
    exp = exp.assign(cantidad=pd.to_numeric(exp["cantidad"], errors="coerce"))
    piv = exp.pivot_table(index="indice_tiempo", columns="pais", values="cantidad", aggfunc="sum").fillna(0)
    piv["TOTAL"] = piv.sum(axis=1)
    print("monthly exports by country, thousand m3 (as published, unit miles/m3):")
    print((piv / 1000).round(1).to_string())
    print("\nfecha_data values for last months:", sorted(exp["fecha_data"].unique())[-6:])
    lng = gas[(gas["tipodecomercializacion"] == "Exportación") & (gas["producto"].astype(str).str.contains("Licuado"))
              & (gas["cantidad"] != 0) & (gas["anio"] >= 2021)]
    print("\nLNG export rows nonzero:", len(lng))
    if len(lng):
        print(lng.groupby(["indice_tiempo", "pais"])["cantidad"].sum().to_string()[:3000])
    imp = gas[(gas["tipodecomercializacion"] == "Importación") & (gas["anio"] >= 2024) & (gas["cantidad"] != 0)]
    print("\nimport gas by product/country since 2024 (rows):", imp.groupby(["producto", "pais"]).size().to_dict())


def ckan():
    print("\n===== dataset resources =====")
    r = requests.get(CKAN_SHOW, headers=HEADERS, timeout=TIMEOUT,
                     params={"id": "refinacion-y-comercializacion-de-petroleo-gas-y-derivados-tablas-dinamicas"})
    for res in r.json()["result"]["resources"]:
        print(f"  {res.get('name')} | {res.get('format')} | modified={res.get('last_modified') or res.get('modified')} "
              f"| {res.get('url')}")


def partes():
    print("\n===== ENARGAS partes diarios exp =====")
    for tipo in ("exp_dentro", "exp_fuera", "imp"):
        r = requests.get(PARTES, headers=HEADERS, timeout=TIMEOUT, params={"tipo": tipo})
        html = r.text
        print(f"tipo={tipo}: status={r.status_code} bytes={len(html)}")
        text = re.sub(r"<script.*?</script>|<style.*?</style>", " ", html, flags=re.S | re.I)
        text = re.sub(r"<[^>]+>", " ", text)
        text = re.sub(r"\s+", " ", text)
        i = text.find("Export")
        print("  text:", text[max(0, i - 200):i + 1500])
        for m in re.finditer(r"<(form|select|input)[^>]*>", html, re.I):
            print("  ", m.group(0)[:200])


def main():
    for fn in (enargas, ckan, partes, se_csv):
        try:
            fn()
        except Exception as e:
            print(f"{fn.__name__} FAILED: {type(e).__name__}: {str(e)[:300]}")


if __name__ == "__main__":
    main()
