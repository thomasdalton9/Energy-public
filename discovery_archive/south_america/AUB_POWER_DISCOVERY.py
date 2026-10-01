"""
One-off probe for the daily power-generation-by-type pulls of Argentina
(CAMMESA), Uruguay (ADME) and Bolivia (CNDC) - checks history depth and
fuel detail before writing the scheduled scripts.

Argentina: the Programacion Diaria ZIP before 24-Oct-2024 holds only an
  Access .mdb (no CSV) - list its tables/columns with mdbtools for a 2021,
  2023 and recent day; print every CSV header in a recent ZIP (any per-unit
  fuel?); list PARTE_POST_OPERATIVO (actual, post-operation) attachments.
Uruguay: gpf.php for a few days in Jan-2021 and recently - sheet names,
  GPF columns, first rows.
Bolivia: the CNDC WordPress REST route index, the monthly document list
  (first/last months, document types) and the gen_dia header (plant
  names) for Jan-2021 and the latest month.

Usage: python3 AUB_POWER_DISCOVERY.py [argentina] [uruguay] [bolivia]
"""

print("STARTING", flush=True)

import datetime as dt
import io
import os
import re
import subprocess
import sys
import tempfile
import zipfile

import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0 Safari/537.36"}
WHICH = set(sys.argv[1:]) or {"argentina", "uruguay", "bolivia"}


# ---------------------------------------------------------------- Argentina
LOOKUP_URL = "https://api.cammesa.com/pub-svc/public/findDocumentosByNemoRango"
ATTACHMENT_URL = "https://api.cammesa.com/pub-svc/public/findAttachmentByNemoId"
TIME_FMT = "%Y-%m-%dT%H:%M:%S.000Z"


def cammesa_docs(nemo, day):
    r = requests.get(LOOKUP_URL, params={"fechadesde": day.strftime(TIME_FMT),
                                         "fechahasta": (day + dt.timedelta(days=1)).strftime(TIME_FMT),
                                         "nemo": nemo}, headers=UA, timeout=60)
    print(f"\n==== {nemo} {day}: {r.status_code} {len(r.content):,} bytes", flush=True)
    try:
        docs = r.json()
    except ValueError:
        print("  not JSON", r.text[:300])
        return []
    for d in docs[:4]:
        print(f"  doc id={d.get('id')} titulo={d.get('titulo')!r} fecha={d.get('fecha')} "
              f"adjuntos={[(a.get('id'), a.get('nombre')) for a in d.get('adjuntos', [])][:20]}", flush=True)
    return docs


def cammesa_get(doc, att, nemo):
    r = requests.get(ATTACHMENT_URL, params={"attachmentId": att["id"], "docId": doc["id"],
                                             "nemo": doc.get("nemo", nemo)}, headers=UA, timeout=180)
    print(f"  attachment {att['id']}: {r.status_code} {r.headers.get('Content-Type')} {len(r.content):,} bytes",
          flush=True)
    return r.content if r.ok else b""


def show_mdb(data, tag, max_tables=60):
    with tempfile.NamedTemporaryFile(suffix=".mdb", delete=False) as f:
        f.write(data)
        path = f.name
    try:
        tables = subprocess.run(["mdb-tables", "-1", path], capture_output=True, text=True, timeout=60).stdout.split("\n")
    except FileNotFoundError:
        print("  mdbtools not installed", flush=True)
        return
    tables = [t for t in tables if t.strip()]
    print(f"  [{tag}] {len(tables)} tables: {tables}", flush=True)
    for t in tables[:max_tables]:
        out = subprocess.run(["mdb-export", path, t], capture_output=True, text=True, timeout=120).stdout
        lines = out.splitlines()
        print(f"\n    -- table {t}: {len(lines) - 1} rows", flush=True)
        for line in lines[:4]:
            print(f"       {line[:400]}", flush=True)
        # look for fuel / type words in the table
        low = out.lower()
        hits = [w for w in ("combust", "gas natural", "gasoil", "fuel", "carbon", "tipo", "tecnolog") if w in low]
        if hits:
            print(f"       keywords: {hits}", flush=True)
    os.unlink(path)


def show_zip(data, tag):
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        print(f"  [{tag}] not a zip: {data[:100]!r}", flush=True)
        if data[:4] == b"\x00\x01\x00\x00":
            show_mdb(data, tag)
        return
    for info in zf.infolist():
        print(f"  [{tag}] member {info.filename} {info.file_size:,} bytes", flush=True)
    for info in zf.infolist():
        body = zf.read(info.filename)
        name = info.filename.lower()
        if name.endswith((".mdb", ".accdb")):
            show_mdb(body, f"{tag}:{info.filename}")
        elif name.endswith((".csv", ".txt")):
            lines = body.decode("utf-8-sig", errors="replace").splitlines()
            print(f"\n    -- {info.filename}: {len(lines)} lines", flush=True)
            for line in lines[:3]:
                print(f"       {line[:400]}", flush=True)
        elif name.endswith(".zip"):
            show_zip(body, f"{tag}:{info.filename}")


def argentina():
    today = dt.date.today()
    for day in [dt.date(2021, 1, 15), dt.date(2023, 6, 15), today - dt.timedelta(days=2)]:
        docs = cammesa_docs("PROGRAMACION_DIARIA", day)
        want = day.strftime("PD%y%m%d.zip")
        for doc in docs:
            for att in doc.get("adjuntos", []):
                if att.get("id") == want:
                    show_zip(cammesa_get(doc, att, "PROGRAMACION_DIARIA"), f"PD {day}")
    for day in [dt.date(2021, 1, 15), today - dt.timedelta(days=3)]:
        docs = cammesa_docs("PARTE_POST_OPERATIVO", day)
        for doc in docs[:1]:
            for att in doc.get("adjuntos", [])[:12]:
                body = cammesa_get(doc, att, "PARTE_POST_OPERATIVO")
                if body[:2] == b"PK" or body[:4] == b"\x00\x01\x00\x00":
                    show_zip(body, f"PPO {day} {att['id']}")
                elif body:
                    text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", body.decode("latin-1", errors="replace")))
                    print(f"    text: {text[:1500]}", flush=True)
    # other candidate report codes with actual (not programmed) generation
    for nemo in ["POST_OPERATIVO", "INFORME_MENSUAL", "INFORME_MENSUAL_DATOS", "BASE_DATOS_INFORME_MENSUAL",
                 "GENERACION_REAL", "TRANSACCIONES_ECONOMICAS", "DTE"]:
        cammesa_docs(nemo, today - dt.timedelta(days=20))


# ---------------------------------------------------------------- Uruguay
def uruguay():
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "south_america"))
    import URUGUAY_ADME as U
    s = requests.Session()
    s.headers.update(UA)
    today = dt.date.today()
    for a, b in [(dt.date(2021, 1, 1), dt.date(2021, 1, 3)), (today - dt.timedelta(days=4), today - dt.timedelta(days=2))]:
        page = s.get(U.GPF_URL, params={"fecha_ini": a.strftime("%d/%m/%Y"), "fecha_fin": b.strftime("%d/%m/%Y"),
                                        "send": "MOSTRAR"}, timeout=180)
        links = re.findall(r"""href=['"]([^'"]*\.(?:ods|xlsx?|csv))['"]""", page.text)
        print(f"\n==== ADME gpf {a}..{b}: {page.status_code}, links {links}", flush=True)
        link = next((l for l in links if l.endswith("_horario.ods")), None)
        if not link:
            continue
        ods = s.get(U.BASE_URL + link if link.startswith("/") else link, timeout=180)
        sheets = U.read_ods(ods.content)
        for name, rows in sheets.items():
            print(f"  sheet {name!r}: {len(rows)} rows", flush=True)
            for row in rows[:8]:
                print(f"     {row[:30]}", flush=True)


# ---------------------------------------------------------------- Bolivia
def bolivia():
    import openpyxl
    r = requests.get("https://www.cndc.bo/wp-json/cndc/v1", headers=UA, timeout=60)
    print(f"\n==== CNDC route index: {r.status_code}", flush=True)
    try:
        for route, info in r.json().get("routes", {}).items():
            print(f"  ROUTE {route} {info.get('methods')} args={list((info.get('endpoints') or [{}])[0].get('args', {}))}",
                  flush=True)
    except ValueError:
        print(r.text[:1000])
    for tipo in ("diaria", "mensual", "anual"):
        r = requests.get("https://www.cndc.bo/wp-json/cndc/v1/estadisticas/categorias", params={"tipo": tipo},
                         headers=UA, timeout=60)
        print(f"\n==== categorias {tipo}: {r.status_code} {r.text[:1500]}", flush=True)
    r = requests.get("https://www.cndc.bo/wp-json/cndc/v1/estadisticas/documentos",
                     params={"categoria_id": 155, "agrupado": "true"}, headers=UA, timeout=60)
    grupos = r.json().get("grupos", [])
    print(f"\n==== documentos: {len(grupos)} months, first {grupos[0].get('periodo')} last {grupos[-1].get('periodo')}",
          flush=True)
    for g in grupos[:1] + grupos[-1:]:
        print(f"  {g.get('periodo')}: {[(d['tipo_doc'], d['archivo_url']) for d in g.get('docs', [])]}", flush=True)
    by_period = {g.get("periodo"): g for g in grupos}
    for g in [grupos[0], by_period.get("2021-01"), by_period.get("2023-07")]:
        if not g:
            continue
        docs = {d["tipo_doc"]: d["archivo_url"] for d in g.get("docs", [])}
        url = next((u for t, u in docs.items() if t.startswith("Generaci")), None)
        print(f"\n==== gen_dia {g.get('periodo')}: {url}", flush=True)
        if not url:
            continue
        x = requests.get(url, headers=UA, timeout=120)
        wb = openpyxl.load_workbook(io.BytesIO(x.content), data_only=True)
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            print(f"  sheet {ws.title}: {len(rows)} rows", flush=True)
            for row in rows[:9] + rows[-4:]:
                print(f"     {row}", flush=True)
    # daily dashboard endpoints people see on the homepage
    for path in ["generacion/diaria", "generacion/tipo", "despacho/diario", "operacion/diaria", "demanda/diaria",
                 "generacion", "historico/generacion/diaria", "tiempo-real/generacion", "generacion/fuente"]:
        r = requests.get(f"https://www.cndc.bo/wp-json/cndc/v1/{path}", headers=UA, timeout=60)
        print(f"  try {path}: {r.status_code} {r.text[:300]!r}", flush=True)


# Round 1 (Oct-2026): the route index lists daily/by-technology endpoints (rt/generacion, dashboard/genbruta/
# detalle, historico/generacion/central); the monthly documentos list returns only the last 20 months unless
# asked by year. Round 2 ('bolivia2') probes those.
def bolivia2():
    base = "https://www.cndc.bo/wp-json/cndc/v1/"
    probes = [
        ("rt/fechas", {}), ("rt/generacion", {"fecha": "2026-09-28"}), ("rt/generacion", {"fecha": "2021-03-10"}),
        ("movil/fechas", {}), ("movil/generacion-datos", {}), ("fecha-operacion", {}),
        ("dashboard/energia", {"modo": "mensual", "anio": 2026, "mes": 8}),
        ("dashboard/energia", {"modo": "diario", "anio": 2026, "mes": 8}),
        ("dashboard/genbruta/detalle", {"modo": "diario", "anio": 2026, "mes": 8}),
        ("dashboard/genbruta/detalle", {"modo": "mensual", "anio": 2026, "mes": 8}),
        ("dashboard/genbruta/detalle", {"tecnologia": "Termoelectrica", "modo": "diario", "anio": 2021, "mes": 3}),
        ("historico/generacion/detalle", {"anio": 2025}),
        ("historico/generacion/central", {"anio": 2026, "mes": 8}),
        ("historico/generacion/central", {"anio": 2021, "mes": 1}),
        ("ga/generacion-bruta", {}),
        ("inyecsti", {"fecha": "2026-09-28"}),
        ("estadisticas/documentos", {"categoria_id": 265, "agrupado": "true"}),
        ("estadisticas/documentos", {"categoria_id": 155, "anio": 2021, "agrupado": "true"}),
        ("estadisticas/documentos", {"categoria_id": 225, "desde": "2020-12-01", "hasta": "2021-03-01"}),
    ]
    for path, params in probes:
        try:
            r = requests.get(base + path, params=params, headers=UA, timeout=90)
            print(f"\n==== {path} {params}: {r.status_code} {len(r.content):,} bytes\n  {r.text[:2500]}", flush=True)
        except requests.RequestException as e:
            print(f"\n==== {path} {params}: {type(e).__name__} {e}", flush=True)


# Round 1 also found that PARTE_POST_OPERATIVO (POyymmdd.zip, ~9 MB: HTML pages + POyymmdd.mdb) carries the
# actual (post-operation) VALORES_GENERADORES.ENERGIA per unit and hour, GENERADORES (TIPO/SUBTIPO per unit)
# and COMBUSTIBLE_PORCENTAJE_DET (fuel % per unit and hour). Round 3 ('argentina2') checks it across 2021-2026:
# energy by TIPO/SUBTIPO, fuel codes and their descriptions, and one-call listing of a month of documents.
def mdb_table(path, table):
    import csv as _csv
    out = subprocess.run(["mdb-export", path, table], capture_output=True, text=True, timeout=300).stdout
    return list(_csv.DictReader(io.StringIO(out)))


def argentina2():
    from collections import defaultdict
    r = requests.get(LOOKUP_URL, params={"fechadesde": "2021-01-01T00:00:00.000Z", "fechahasta": "2021-02-01T00:00:00.000Z",
                                         "nemo": "PARTE_POST_OPERATIVO"}, headers=UA, timeout=120)
    docs = r.json() if r.ok else []
    print(f"\n==== PO listing Jan-2021 in one call: {r.status_code}, {len(docs) if isinstance(docs, list) else docs}",
          flush=True)
    if isinstance(docs, list):
        print("  ", [(d.get("fecha"), [a.get("id") for a in d.get("adjuntos", [])]) for d in docs[:40]], flush=True)
    for day in [dt.date(2021, 1, 15), dt.date(2022, 7, 15), dt.date(2024, 6, 15), dt.date.today() - dt.timedelta(days=3)]:
        docs = cammesa_docs("PARTE_POST_OPERATIVO", day)
        want = day.strftime("PO%y%m%d.zip")
        hit = [(d, a) for d in (docs if isinstance(docs, list) else []) for a in d.get("adjuntos", []) if a.get("id") == want]
        if not hit:
            print(f"  no {want}", flush=True)
            continue
        body = cammesa_get(*hit[0], "PARTE_POST_OPERATIVO")
        zf = zipfile.ZipFile(io.BytesIO(body))
        mdb = next(n for n in zf.namelist() if n.lower().endswith(".mdb"))
        with tempfile.NamedTemporaryFile(suffix=".mdb", delete=False) as f:
            f.write(zf.read(mdb))
            path = f.name
        tables = subprocess.run(["mdb-tables", "-1", path], capture_output=True, text=True).stdout.split()
        print(f"  {mdb}: tables {tables}", flush=True)
        gens = {g["GRUPO"]: g for g in mdb_table(path, "GENERADORES")}
        vals = mdb_table(path, "VALORES_GENERADORES")
        hours = sorted({int(v["HORA"]) for v in vals})
        print(f"  VALORES_GENERADORES hours {hours[:3]}..{hours[-3:]} ({len(hours)})", flush=True)
        by_type = defaultdict(float)
        energy = {}
        for v in vals:
            g = gens.get(v["GRUPO"], {})
            e = float(v["ENERGIA"] or 0)
            by_type[(g.get("TIPO"), g.get("SUBTIPO"), g.get("INTERCAMBIO"))] += e
            energy[(v["GRUPO"], v["HORA"])] = e
        for k, e in sorted(by_type.items(), key=lambda kv: -kv[1]):
            print(f"    {k}: {e:,.0f}", flush=True)
        print(f"    TOTAL {sum(by_type.values()):,.0f}", flush=True)
        missing = sorted({v["GRUPO"] for v in vals if v["GRUPO"] not in gens})[:20]
        print(f"    units without GENERADORES row: {missing}", flush=True)
        by_fuel = defaultdict(float)
        for c in mdb_table(path, "COMBUSTIBLE_PORCENTAJE_DET"):
            e = energy.get((c["GRUPO"], c["HORA"]), 0)
            by_fuel[(c["COMB"], gens.get(c["GRUPO"], {}).get("SUBTIPO"))] += e * float(c["PORCENTAJE"] or 0) / 100
        print("    thermal energy by (fuel code, subtipo):", {k: round(v) for k, v in sorted(by_fuel.items(), key=lambda kv: -kv[1])},
              flush=True)
        if "CVP" in tables:
            print("    CVP fuel codes:", sorted({(c["COMB"], c["COMB_DESC"]) for c in mdb_table(path, "CVP")}), flush=True)
        if "ENERGIAS_RENOVABLES" in tables:
            ren = defaultdict(float)
            for x in mdb_table(path, "ENERGIAS_RENOVABLES"):
                ren[gens.get(x["GRUPO"], {}).get("SUBTIPO")] += float(x["EGENERADA"] or 0)
            print("    ENERGIAS_RENOVABLES.EGENERADA by subtipo:", {k: round(v) for k, v in ren.items()}, flush=True)
        os.unlink(path)


for name, fn in [("argentina", argentina), ("uruguay", uruguay), ("bolivia", bolivia), ("bolivia2", bolivia2),
                 ("argentina2", argentina2)]:
    if name in WHICH:
        try:
            fn()
        except Exception as e:  # keep going to the next country
            print(f"\n#### {name} FAILED: {type(e).__name__}: {e}", flush=True)
print("\nDONE", flush=True)
