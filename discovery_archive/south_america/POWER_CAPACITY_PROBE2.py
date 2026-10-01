"""
Round 2 of POWER_CAPACITY_PROBE.py.

Round 1 found: ANEEL SIGA has each operating plant's first start date
(DatEntradaOperacao) but no retirement date; CNE publishes
Capacidad_Instalada_Generacion.xlsx monthly (current snapshot, unit level);
XM CapEfecNeta per plant per day works back to 2021.

This round checks:
  brazil: ANEEL unit-level commercial-operation releases (liberacao para
          operacao comercial, detalhado), the historic operating-plant count
          table, annual additions, and the distributed generation (MMGD)
          parquet schema and totals by connection year.
  chile:  the CNE workbook's 'base' sheet columns, status values and totals;
          whether past monthly uploads are still at their old
          wp-content/uploads/YYYY/MM/ URLs; energiaabierta.cl.

    python3 POWER_CAPACITY_PROBE2.py brazil|chile

Found (Oct 2026): ANEEL's unit-release list gives yearly additions close to
SIGA's plant start dates (2021-25: 7.6/8.3/10.3/10.8/7.5 GW); ANEEL also
publishes year-end MW by plant type (empreendimento-operacao-historico.csv,
used as a check sheet). MMGD parquet: 4.66 million systems, 54 GW, dated by
DthAtualizaCadastralEmpreend. CNE 'base' sheet: one row per SEN plant with
fecha_puesta_servicio_central and potencia_neta_mw, no retired plants; no older
uploads at old wp-content URLs; Wayback CDX returned 503.
"""
import io
import re
import sys

import requests

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"}


def get(url, **kw):
    try:
        r = requests.get(url, headers=kw.pop("headers", UA), timeout=kw.pop("timeout", (15, 300)), **kw)
        print(f"GET {r.url[:200]} -> {r.status_code} {r.headers.get('content-type')} {len(r.content):,} B", flush=True)
        return r
    except requests.RequestException as e:
        print(f"GET {url} FAILED {type(e).__name__}: {e}", flush=True)
        return None


def hdr(s):
    print(f"\n{'=' * 78}\n{s}\n{'=' * 78}", flush=True)


def brazil():
    import pandas as pd
    pd.set_option("display.width", 250)
    pd.set_option("display.max_columns", 40)
    base = "https://dadosabertos.aneel.gov.br/dataset/"
    hdr("liberacao operacao comercial - detalhado")
    r = get(base + "2b2ace01-5692-4636-8c99-2a84ff094f4f/resource/75419902-c692-498b-a6ef-85f6d4beb5b2/download/"
            "unidades-geradoras-liberadas-operacao-comercial-detalhado.csv")
    for enc in ("utf-8", "latin-1"):
        try:
            d = pd.read_csv(io.BytesIO(r.content), sep=";", encoding=enc, dtype=str, low_memory=False)
            break
        except UnicodeDecodeError:
            continue
    print("columns:", list(d.columns), len(d))
    print(d.head(4).T.to_string())
    for c in d.columns:
        if d[c].nunique() < 40:
            print(f"\n{c}: {d[c].value_counts(dropna=False).head(20).to_dict()}")
    for c in [c for c in d.columns if c.lower().startswith("dat")]:
        s = pd.to_datetime(d[c], errors="coerce")
        print(f"{c}: {s.min()} .. {s.max()}, null {s.isna().sum()}")
        print(s.dt.year.value_counts().sort_index().tail(8).to_dict())
    pot = [c for c in d.columns if "Pot" in c]
    print("power columns:", pot)
    if pot:
        kw = pd.to_numeric(d[pot[0]].str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
                           errors="coerce")
        dc = [c for c in d.columns if c.lower().startswith("dat")]
        if dc:
            y = pd.to_datetime(d[dc[-1]], errors="coerce").dt.year
            print(f"sum {pot[0]} by year of {dc[-1]} (GW):", (kw.groupby(y).sum() / 1e6).round(2).tail(8).to_dict())

    hdr("empreendimento-operacao-historico")
    r = get(base + "306a6fdb-beb9-4296-bf18-77fa0e076ef1/resource/e61fd029-5e78-43be-bed4-873b7b11f04c/download/"
            "empreendimento-operacao-historico.csv")
    print(r.content.decode("utf-8", errors="replace")[:3000])
    hdr("acrescimo-anual-potencia-instalada")
    r = get(base + "4bf07812-e421-48f3-8e48-eec0040764ba/resource/3b3eb013-8831-4ca9-a582-01010ff4e0fb/download/"
            "acrescimo-anual-potencia-instalada.csv")
    print(r.content.decode("utf-8", errors="replace")[:1500])

    hdr("GD parquet")
    import pyarrow.parquet as pq
    r = get(base + "5e0fafd2-21b9-4d5b-b622-40438d40aba2/resource/cd29f6eb-e08d-4db7-b6fb-ed6e3b682d27/download/"
            "empreendimento-geracao-distribuida.parquet", timeout=(15, 600))
    f = pq.ParquetFile(io.BytesIO(r.content))
    print(f.schema_arrow)
    print("rows", f.metadata.num_rows)
    t = f.read().to_pandas()
    print(t.head(3).T.to_string())
    for c in t.columns:
        if t[c].nunique() < 30:
            print(f"\n{c}: {t[c].value_counts(dropna=False).head(15).to_dict()}")
    pot = [c for c in t.columns if "Potencia" in c]
    dts = [c for c in t.columns if c.startswith(("Dat", "Dth"))]
    print("pot", pot, "dates", dts)
    for c in dts:
        s = pd.to_datetime(t[c], errors="coerce")
        print(f"{c}: {s.min()} .. {s.max()} null {s.isna().sum()}")
    if pot:
        kw = pd.to_numeric(t[pot[0]].astype(str).str.replace(",", ".", regex=False), errors="coerce")
        print(f"total {pot[0]}: {kw.sum() / 1e6:.2f} GW")
        if "SigTipoGeracao" in t:
            print((kw.groupby(t["SigTipoGeracao"]).sum() / 1e6).round(2).to_dict())
        for c in dts:
            y = pd.to_datetime(t[c], errors="coerce").dt.to_period("Y")
            print(f"by {c} year (GW cum):", (kw.groupby(y).sum().cumsum() / 1e6).round(2).tail(8).to_dict())


def chile():
    import openpyxl
    import pandas as pd
    pd.set_option("display.width", 250)
    pd.set_option("display.max_columns", 50)
    hdr("CNE base sheet")
    url = "https://www.cne.cl/wp-content/uploads/2026/09/Capacidad_Instalada_Generacion.xlsx"
    r = get(url)
    b = pd.read_excel(io.BytesIO(r.content), sheet_name="base")
    print(list(b.columns), len(b))
    print(b.head(3).T.to_string())
    for c in b.columns:
        if b[c].nunique() < 45:
            print(f"\n{c}: {b[c].value_counts(dropna=False).head(45).to_dict()}")
    num = [c for c in b.columns if re.search(r"pot|mw|capac", str(c), re.I)]
    print("numeric-ish:", num)
    for c in num:
        print(c, pd.to_numeric(b[c], errors="coerce").sum())
    if "sistema" in b and num:
        p = pd.to_numeric(b[num[0]], errors="coerce")
        for k in ["tipo_tecnologia", "tecnologia", "tipo_energia", "combustible"]:
            if k in b:
                print(f"\n{num[0]} by sistema x {k}:\n",
                      p.groupby([b["sistema"], b[k]]).sum().round(1).to_string())
    wb = openpyxl.load_workbook(io.BytesIO(r.content), read_only=True, data_only=True)
    ws = wb["SEN"]
    rows = list(ws.iter_rows(values_only=True))
    print("SEN header rows:", rows[1][:40])
    for row in rows[-40:]:
        vals = [v for v in row if v is not None]
        if vals:
            print("   tail:", vals[:20])

    hdr("Past monthly uploads at old URLs")
    names = ["Capacidad_Instalada_Generacion", "Capacidad_Instalada_Generacion-1", "Capacidad-Instalada-Generacion",
             "Capacidad_Instalada_de_Generacion"]
    for y in range(2020, 2027):
        for m in range(1, 13):
            for n in names:
                for ext in ("xlsx",):
                    u = f"https://www.cne.cl/wp-content/uploads/{y}/{m:02d}/{n}.{ext}"
                    try:
                        h = requests.head(u, headers=UA, timeout=(10, 30), allow_redirects=True)
                    except requests.RequestException:
                        continue
                    if h.status_code == 200:
                        print(f"  FOUND {u} {h.headers.get('content-length')}", flush=True)
    hdr("Wayback CDX for the CNE file")
    r = get("http://web.archive.org/cdx/search/cdx", params={"url": "cne.cl/wp-content/uploads/*Capacidad*",
                                                            "output": "json", "limit": 500,
                                                            "filter": "statuscode:200"}, timeout=(15, 120))
    if r is not None and r.ok:
        try:
            for row in r.json()[:200]:
                print("  ", row)
        except ValueError:
            print(r.text[:2000])
    hdr("energiaabierta.cl")
    for u in ["http://energiaabierta.cl/?s=capacidad+instalada",
              "https://energiaabierta.cl/visualizaciones/capacidad-instalada-de-generacion/",
              "http://datos.energiaabierta.cl/dataviews/",
              "https://api.energiaabierta.cl/"]:
        r = get(u, timeout=(10, 60))
        if r is not None and r.ok:
            links = sorted(set(re.findall(r"https?://[^\"'\s<>]+(?:capacidad|Capacidad)[^\"'\s<>]*", r.text)))
            print("   ", links[:30])


if __name__ == "__main__":
    {"brazil": brazil, "chile": chile}[sys.argv[1]]()
