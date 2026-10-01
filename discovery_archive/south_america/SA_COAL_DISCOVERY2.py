"""
South America coal - discovery round 2 (after SA_COAL_DISCOVERY.py found):
  - datos.gov.co r85m-vv6c "ANM Volumen de Explotacion de Minerales Asociados a Pagos de Regalias"
    (quarterly by municipality / department / mineral): coal mineral names, units, totals by year
  - DANE anex-EXPORTACIONES-SerieCafeCarbonPetroleoNotradicionales-<mes><yyyy>.xlsx (monthly coal exports)
  - Energy Institute: energyinst.org returns 403 to scripts -> OWID ETL copy of the Statistical Review
  - Brazil EPE BEN historical series; Argentina BEN xlsx; Chile energiaabierta BNE; Peru MINEM
"""
import io
import re
import sys

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "es,en;q=0.8"}
S = requests.Session()
S.headers.update(H)
pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 30)
pd.set_option("display.max_colwidth", 40)


def get(url, **kw):
    kw.setdefault("timeout", (15, 120))
    try:
        r = S.get(url, **kw)
        print(f"  GET {r.url[:200]} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)} bytes")
        return r
    except Exception as e:
        print(f"  GET {url} -> {type(e).__name__}: {str(e)[:150]}")
        return None


def links(html, pat, base=""):
    out = []
    for m in re.finditer(r'href=["\']([^"\']+)["\']', html, re.I):
        h = m.group(1)
        if re.search(pat, h, re.I):
            if h.startswith("/") and base:
                h = base + h
            out.append(h)
    return list(dict.fromkeys(out))


def section(t):
    print("\n" + "=" * 100 + f"\n{t}\n" + "=" * 100, flush=True)


def soql(q):
    r = get("https://www.datos.gov.co/resource/r85m-vv6c.json", params={"$query": q})
    if r is not None and r.status_code == 200:
        return pd.DataFrame(r.json())
    if r is not None:
        print(f"     {r.text[:300]}")
    return pd.DataFrame()


def anm():
    section("ANM r85m-vv6c")
    r = get("https://www.datos.gov.co/api/views/r85m-vv6c.json")
    if r is not None and r.status_code == 200:
        m = r.json()
        print(f"     name {m.get('name')}; rowsUpdatedAt {m.get('rowsUpdatedAt')}; desc {str(m.get('description'))[:600]}")
        for c in m.get("columns", []):
            print(f"     col {c.get('fieldName')} ({c.get('dataTypeName')}): {c.get('name')}")
    print(soql("SELECT recurso_natural, unidad_medida, count(*) GROUP BY recurso_natural, unidad_medida ORDER BY recurso_natural LIMIT 500").to_string())
    print(soql("SELECT periodicidad2, contraprestacion, count(*) GROUP BY periodicidad2, contraprestacion").to_string())
    vol = "volumenes_de_explotacion"
    print(soql(f"SELECT recurso_natural, a_o_liquidado, trimestre, sum({vol}) AS t, count(*) WHERE recurso_natural like '%CARB%' "
               "GROUP BY recurso_natural, a_o_liquidado, trimestre ORDER BY a_o_liquidado, trimestre LIMIT 2000").to_string())
    print(soql(f"SELECT nombre_departamento, a_o_liquidado, sum({vol}) AS t WHERE recurso_natural like '%CARB%' AND a_o_liquidado >= '2021' "
               "GROUP BY nombre_departamento, a_o_liquidado ORDER BY nombre_departamento, a_o_liquidado LIMIT 2000").to_string())
    print(soql(f"SELECT * WHERE recurso_natural like '%CARB%' AND a_o_liquidado = '2025' AND nombre_departamento like 'La Guajira%' LIMIT 20").to_string())
    print(soql(f"SELECT periodo_liquidado2, trimestre, periodicidad2, count(*) WHERE recurso_natural like '%CARB%' "
               "GROUP BY periodo_liquidado2, trimestre, periodicidad2 ORDER BY periodo_liquidado2 LIMIT 100").to_string())


def dane():
    section("DANE coal exports series")
    for u in ["https://www.dane.gov.co/files/operaciones/EXPORTACIONES/anex-EXPORTACIONES-SerieCafeCarbonPetroleoNotradicionales-jul2026.xlsx",
              "https://www.dane.gov.co/files/operaciones/EXPORTACIONES/anex-EXPORTACIONES-SerieCafeCarbonPetroleoNotradicionales-ago2026.xlsx"]:
        r = get(u)
        if r is None or r.status_code != 200:
            continue
        xl = pd.ExcelFile(io.BytesIO(r.content))
        print(f"     sheets {xl.sheet_names}")
        for s in xl.sheet_names:
            df = xl.parse(s, header=None)
            print(f"\n     --- {s} {df.shape}")
            print(df.head(16).to_string(max_colwidth=30))
            print("     ...")
            # rows from 2020 on
            yr = df[df.apply(lambda r: r.astype(str).str.contains(r"^202[0-6]").any(), axis=1)]
            print(yr.head(8).to_string(max_colwidth=30))
            print(df.tail(14).to_string(max_colwidth=30))


def owid():
    section("Energy Institute via OWID ETL")
    for u in ["https://www.energyinst.org/__data/assets/excel_doc/0020/1540550/EI-Stats-Review-All-Data.xlsx",
              "https://www.energyinst.org/__data/assets/file/0003/1540551/EI-Stats-Review-ALL-data.xlsx"]:
        get(u)
    r = get("https://api.github.com/repos/owid/etl/contents/etl/steps/data/garden/energy_institute")
    vers = []
    if r is not None and r.status_code == 200:
        vers = sorted(d["name"] for d in r.json() if d["type"] == "dir")
        print(f"     versions: {vers}")
    for v in reversed(vers[-3:]):
        base = f"https://catalog.ourworldindata.org/garden/energy_institute/{v}/statistical_review_of_world_energy/statistical_review_of_world_energy"
        r = get(base + ".csv")
        if r is None or r.status_code != 200:
            continue
        df = pd.read_csv(io.BytesIO(r.content), low_memory=False)
        cc = [c for c in df.columns if "coal" in c and ("prod" in c or "export" in c)]
        print(f"     {v}: shape {df.shape}; coal cols {cc}")
        sub = df[df["country"].isin(["Colombia", "Brazil", "Venezuela", "Chile", "Argentina", "Peru",
                                     "Other South and Central America (EI)", "South and Central America (EI)"]) & (df["year"] >= 2019)]
        cols_ = [c for c in cc if c.endswith("_mt") or "tonne" in c] or cc[:4]
        print(sub[["country", "year"] + cols_].to_string())
        others = df[(df["year"] == 2024) & df[cols_[0]].notna()]["country"].tolist() if cols_ else []
        print(f"     countries with coal production (mt) 2024: {[c for c in others if 'America' in c or c in ('Chile','Peru','Argentina','Venezuela','Mexico')]}")
        break


def brazil():
    section("Brazil EPE BEN")
    r = get("https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes/BEN-Series-Historicas-Completas")
    if r is not None and r.status_code == 200:
        ls = links(r.text, r"\.xlsx?|\.zip|sites/pt/publicacoes", "https://www.epe.gov.br")
        for h in ls[:50]:
            print(f"     {h}")
        for h in [x for x in ls if re.search(r"\.xlsx?$", x, re.I)][:12]:
            if re.search(r"produ|carv|cap|fontes|energia|bal|consolid", h, re.I):
                rr = get(h)
                if rr is None or rr.status_code != 200:
                    continue
                try:
                    xl = pd.ExcelFile(io.BytesIO(rr.content))
                except Exception as e:
                    print(f"     {e}")
                    continue
                print(f"     sheets {xl.sheet_names[:40]}")
                for s in xl.sheet_names[:40]:
                    df = xl.parse(s, header=None)
                    for i, row in df.iterrows():
                        t = " | ".join(str(v) for v in row.values if str(v) != "nan")
                        if re.search(r"carv", t, re.I):
                            print(f"     [{s} r{i}] {t[:400]}")


def argentina():
    section("Argentina BEN")
    for u in ["https://www.energia.gob.ar/contenidos/archivos/Reorganizacion/informacion_del_mercado/publicaciones/energia_en_gral/balances_2025/balance_2025_v0_h.xlsx",
              "https://www.argentina.gob.ar/sites/default/files/balance_2024_v0_h.xlsx"]:
        r = get(u)
        if r is None or r.status_code != 200:
            continue
        xl = pd.ExcelFile(io.BytesIO(r.content))
        print(f"     sheets {xl.sheet_names}")
        for s in xl.sheet_names:
            df = xl.parse(s, header=None)
            print(f"     --- {s} {df.shape}")
            print(df.head(8).to_string(max_colwidth=22)[:3000])
            for i, row in df.iterrows():
                t = " | ".join(str(v) for v in row.values if str(v) != "nan")
                if re.search(r"carb|produc|unidad|ktep|tonel", t, re.I):
                    print(f"     [{s} r{i}] {t[:400]}")
        break
    r = get("https://datos.energia.gob.ar/dataset/balances-energeticos")
    if r is not None and r.status_code == 200:
        for h in links(r.text, r"\.csv|\.xlsx|download", "https://datos.energia.gob.ar")[:30]:
            print(f"     {h}")


def chile():
    section("Chile")
    for u in ["https://energiaabierta.cl/categorias-estadistica/balance-energetico/",
              "https://energiaabierta.cl/visualizaciones/national-energy-balance/?lang=en",
              "https://www.cne.cl/estadisticas/hidrocarburo/",
              "https://www.cne.cl/estadisticas/"]:
        r = get(u)
        if r is not None and r.status_code == 200:
            for h in links(r.text, r"balance|bne|xls|carb|csv|dataset")[:40]:
                print(f"     {h}")


def peru():
    section("Peru")
    for u in ["https://www.gob.pe/busquedas?term=boletin%20estadistico%20minero&institucion=minem",
              "https://www.gob.pe/busquedas?term=produccion%20minera%20no%20metalica",
              "https://www.gob.pe/institucion/minem/colecciones/16233-anuario-minero",
              "https://www.gob.pe/institucion/minem/colecciones/16232-boletin-estadistico-minero"]:
        r = get(u)
        if r is not None and r.status_code == 200:
            ls = links(r.text, r"minem/(informes|colecciones|normas)|\.xlsx?|\.pdf", "https://www.gob.pe")
            for h in ls[:40]:
                print(f"     {h}")


if __name__ == "__main__":
    for w in sys.argv[1:] or ["anm", "dane", "owid", "brazil", "argentina", "chile", "peru"]:
        try:
            globals()[w]()
        except Exception as e:
            print(f"{w} failed: {type(e).__name__}: {e}")
