"""
Round 4: Argentina. Round 3 found two raw CAMMESA sources of installed capacity:
  1. Secretaria de Energia open data, dataset publicaciones-cammesa, resource
     'Potencia Instalada' (potencia-instalada.csv): CAMMESA installed MW per
     power plant x machine type x month from 2015-10.
  2. CAMMESA's own 'Potencia Instalada.xlsx' (Sintesis Mensual > Estadisticas,
     microfe.cammesa.com static content, linked from
     cammesaweb.cammesa.com/download/potencia-instalada/): annual summary by
     machine type 2002-2026, autogenerators detail, and retirements by month.
This checks how current (1) is, the technology labels and monthly totals, and
the full layout of (2).

    python3 POWER_CAPACITY_PROBE4.py

Found (Oct 2026): the SE CSV stops at 2020-02. CAMMESA's workbook has year-end
MW by machine type 2002-2025 plus the current year, a 633-machine list for the
latest month (213 with FECHA HABILITACION) summing exactly to the current-year
total, and 57 retirements since 2021 with MES BAJA. -> ARGENTINA_CAMMESA_CAPACITY.py.
"""
import io

import pandas as pd
import requests

pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 40)
pd.set_option("display.max_rows", 400)
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"}
CSV = ("http://datos.energia.gob.ar/dataset/2b4dfee6-6fca-4e4d-9611-a12d65cd4aa8/resource/"
       "b05fbb16-7278-463f-8895-087e2495bfee/download/potencia-instalada.csv")
XLSX = ("https://microfe.cammesa.com/static-content/CammesaWeb/download-manager-files/Sintesis%20Mensual/"
        "Estadisticas/Potencia%20Instalada.xlsx")

print("=" * 30, "SE CSV", flush=True)
r = requests.get(CSV, headers=UA, timeout=(20, 300))
print(r.status_code, len(r.content), r.headers.get("last-modified"))
d = pd.read_csv(io.BytesIO(r.content), encoding="utf-8-sig", low_memory=False)
print(d.dtypes)
d["periodo"] = pd.to_datetime(d["periodo"], errors="coerce")
print("periodo", d["periodo"].min(), d["periodo"].max(), "rows", len(d))
print(d.groupby("periodo").size().tail(30).to_string())
for c in ["region", "tipo_maquina", "fuente_generacion", "tecnologia", "categoria_region"]:
    print(f"\n{c}:", d[c].value_counts().to_dict())
p = d.pivot_table(index="periodo", columns="tecnologia", values="potencia_instalada_mw", aggfunc="sum")
p["TOTAL"] = p.sum(axis=1)
print(p[p.index >= "2020-12-01"].round(0).to_string())
print(d.pivot_table(index="periodo", columns="tipo_maquina", values="potencia_instalada_mw", aggfunc="sum")
      .iloc[-3:].round(0).T.to_string())
print(d[d["periodo"] == d["periodo"].max()].sort_values("potencia_instalada_mw", ascending=False).head(15).to_string())

print("=" * 30, "CAMMESA xlsx", flush=True)
r = requests.get(XLSX, headers=UA, timeout=(20, 300))
print(r.status_code, len(r.content), r.headers.get("last-modified"))
xl = pd.ExcelFile(io.BytesIO(r.content))
for s in xl.sheet_names:
    t = xl.parse(s, header=None)
    print(f"\n--- {s} {t.shape}")
    if s.startswith("RESUMEN"):
        print(t.dropna(how="all").to_string(max_colwidth=22))
    else:
        print(t.dropna(how="all").head(8).to_string(max_colwidth=22))
        print(t.dropna(how="all").tail(5).to_string(max_colwidth=22))
        hdr_row = t.index[t.iloc[:, 0].astype(str).str.strip().eq("AÑO")]
        if len(hdr_row):
            b = xl.parse(s, header=hdr_row[0])
            for c in b.columns:
                if b[c].nunique() < 40:
                    print(f"   {c}: {b[c].value_counts().head(40).to_dict()}")
