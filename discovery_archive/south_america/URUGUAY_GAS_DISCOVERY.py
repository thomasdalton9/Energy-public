"""
One-off probe: what Uruguay natural gas demand data (by sector / tariff) is
downloadable, and in what layout.

Round 1 (pages + CKAN): catalogodatos.gub.uy has nothing for 'gas natural';
ben.miem.gub.uy no longer resolves; the URSEA statistics page has no gas
files. The MIEM 'Series estadisticas de gas natural' page offers one zip
of all attachments (download/node/field_documento/3815) plus a monthly
visualiser (visualpeb.miem.gub.uy/visualPEB/gas_natural).

Round 2: zip 3815 is a zip of zips - billing by tariff (Conecta Paysandu,
Conecta Sur, Montevideo Gas), customers by tariff, prices, and
importacion_gas_natural_por_gasoducto_m3. visualPEB also has a BEN
endpoint benDatosActividadesFuentes (annual consumption by activity).

Round 3 (this version): save every inner file plus the visualPEB pages /
BEN endpoint responses under uy_raw/ (the workflow pushes them to a
throwaway branch for offline inspection) and print sheet summaries.
Runs in GitHub Actions only.
"""
import io
import os
import re
import zipfile
from urllib.parse import urljoin

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 90)
RAW = "uy_raw"
ZIP = "https://www.gub.uy/ministerio-industria-energia-mineria/download/node/field_documento/3815"
VP = "https://visualpeb.miem.gub.uy/visualPEB/"
BEN_Q = ("benDatosActividadesFuentes?actividades=actividades&actividades=CONSUMO+FINAL+ENERG%C3%89TICO"
         "&actividades=fuentes&fuentes=Gas+Natural&anioDesde=1998&anioHasta=2025&unidades=ktep")
SECTORS = ["RESIDENCIAL", "COMERCIAL/SERVICIOS/SECTOR P%C3%9ABLICO", "TRANSPORTE", "INDUSTRIAL",
           "ACTIVIDADES PRIMARIAS", "CENTRALES EL%C3%89CTRICAS SERVICIO P%C3%9ABLICO"]


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=T, **kw)
        out(f"GET {url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)} bytes")
        return r
    except requests.RequestException as e:
        out(f"GET {url} -> ERROR {e}")
        return None


def save(name, content):
    path = os.path.join(RAW, re.sub(r"[^\w.\-]+", "_", name))
    with open(path, "wb") as f:
        f.write(content)
    return path


def walk_zip(content, prefix=""):
    z = zipfile.ZipFile(io.BytesIO(content))
    for n in z.namelist():
        data = z.read(n)
        out(f"  member {prefix}{n} ({len(data)} bytes, magic {data[:4]!r})")
        if data[:2] == b"PK" and not n.lower().endswith((".xlsx", ".xlsm")):
            try:
                walk_zip(data, prefix + n + "/")
                continue
            except zipfile.BadZipFile:
                pass
        p = save(prefix + n, data)
        try:
            sheets = pd.read_excel(p, sheet_name=None, header=None)
            for sn, raw in sheets.items():
                out(f"    sheet {sn!r} {raw.shape}")
        except Exception as e:
            out(f"    (not excel: {e}) head: {data[:300]!r}")


ROUND4 = [  # visualPEB form actions (round 4): annual gas flows by sector and BEN consumption by source/sector
    "gnFlujosSectores?anio=2025&unidades=m%C2%B3",
    "gnFlujosSectores?anio=2021&unidades=m%C2%B3",
    "gnFlujos?anio=2025&unidades=m%C2%B3",
    "gnParticipacion?anio=2025&unidades=m%C2%B3",
    "benConsumoPorFuenteOSector?fuenteOSector=fuente&anioDesde=2010&anioHasta=2025&unidades=ktep",
    "benConsumoFinalParaSector?sectores=industrial&anioDesde=2010&anioHasta=2025&unidades=ktep",
    "benInsumosElectricidad?anioDesde=2010&anioHasta=2025&unidades=ktep",
    "gnFlujosTarifasMes?mesDesde=Ene&anioDesde=2021&mesHasta=Dic&anioHasta=2026&unidades=m%C2%B3",
]


ROUND5 = [  # round 5: monthly flows by tariff need the hidden ultimoAnioSeleccionable field
    "gnFlujosTarifasMes?ultimoAnioSeleccionable=2026&mesDesde=Ene&anioDesde=2026&mesHasta=Ene&anioHasta=2026"
    "&unidades=m%C2%B3",
    "gnFlujosTarifasMes?ultimoAnioSeleccionable=2026&mesDesde=Mar&anioDesde=2026&mesHasta=Mar&anioHasta=2026"
    "&unidades=m%C2%B3",
    "gnFlujosTarifasMes?ultimoAnioSeleccionable=2026&mesDesde=Ene&anioDesde=2021&mesHasta=Dic&anioHasta=2026"
    "&unidades=m%C2%B3",
    "gnFacturacionMensualDistrib?distrib=Montevideo+Gas&mesDesde=Ene&anioDesde=2025&mesHasta=Dic&anioHasta=2026",
    "gnImportacionGasoductoMes?mesDesde=Ene&anioDesde=2025&mesHasta=Dic&anioHasta=2026",
]


def main():
    os.makedirs(RAW, exist_ok=True)
    rounds = {"4": ROUND4, "5": ROUND5}
    if os.environ.get("ROUND") in rounds:
        for i, q in enumerate(rounds[os.environ["ROUND"]]):
            r = get(VP + q)
            if r is not None:
                save(f"r{os.environ['ROUND']}_{i}_{q.split('?')[0]}.html", r.content)
        return
    r = get(ZIP)
    if r is not None and r.status_code == 200:
        walk_zip(r.content)
    for page in ["gas_natural", "ben", BEN_Q]:
        r = get(VP + page)
        if r is not None:
            save(f"visualpeb_{page[:40]}.html", r.content)
    # BEN gas by sector: try the same endpoint with each consuming activity
    for s in SECTORS:
        q = BEN_Q.replace("CONSUMO+FINAL+ENERG%C3%89TICO", s.replace(" ", "+"))
        r = get(VP + q)
        if r is not None:
            save(f"visualpeb_ben_{s[:12]}.html", r.content)
    # links/scripts the visualPEB gas page calls
    r = get(VP + "gas_natural")
    if r is not None:
        for u in sorted(set(re.findall(r"""(?:url|href|src|action)\s*[:=]\s*["']([^"']+)["']""", r.text))):
            out(f"  ref {u}")


if __name__ == "__main__":
    main()
