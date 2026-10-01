"""
South America coal - discovery round 3:
  - ANM r85m-vv6c: all coal rows since 2019 pulled and summed locally (volumes are text "1,234.5" / "- 0")
    -> totals by year / department / big projects, to see if the royalty volumes cover Cerrejon, Drummond, etc.
  - EI Statistical Review: the all-data xlsx as mirrored in OWID's ETL snapshot store (energyinst.org returns 403)
  - Brazil EPE BEN chapter 2 (html-unescaped link): coal production in 10^3 t
  - Argentina BEN 2021-2025 (coal production, ktep; any calorific factor in the file)
  - Chile BNE on datos.energiaabierta.cl; Peru MINEM collections
"""
import html
import io
import re
import sys
from urllib.parse import quote

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "es,en;q=0.8"}
S = requests.Session()
S.headers.update(H)
pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 40)
pd.set_option("display.max_rows", 400)


def get(url, **kw):
    kw.setdefault("timeout", (15, 180))
    try:
        r = S.get(url, **kw)
        print(f"  GET {r.url[:220]} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)} bytes")
        return r
    except Exception as e:
        print(f"  GET {url[:220]} -> {type(e).__name__}: {str(e)[:200]}")
        return None


def links(text, pat, base=""):
    out = []
    for m in re.finditer(r'href=["\']([^"\']+)["\']', text, re.I):
        h = html.unescape(m.group(1))
        if re.search(pat, h, re.I):
            if h.startswith("/") and base:
                h = base + h
            out.append(h)
    return list(dict.fromkeys(out))


def section(t):
    print("\n" + "=" * 100 + f"\n{t}\n" + "=" * 100, flush=True)


def num(s):
    s = str(s).strip()
    if s in ("", "-", "- 0", "nan", "None"):
        return 0.0
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return float("nan")


def anm():
    section("ANM coal rows")
    rows, off = [], 0
    while True:
        r = get("https://www.datos.gov.co/resource/r85m-vv6c.json",
                params={"$where": "recurso_natural like '%CARB%' AND recurso_natural != 'CARBONATO DE CALCIO' AND a_o_liquidado >= '2019'",
                        "$limit": 50000, "$offset": off, "$order": ":id"})
        if r is None or r.status_code != 200:
            break
        d = r.json()
        rows += d
        if len(d) < 50000:
            break
        off += 50000
    df = pd.DataFrame(rows)
    print(f"     rows {len(df)}; cols {list(df.columns)}")
    df["t"] = df["volumenes_de_explotacion"].map(num)
    df["roy"] = df["regalias_pagadas"].map(num)
    bad = df[df["t"].isna()]["volumenes_de_explotacion"].unique()[:20]
    print(f"     unparsed volumes: {bad}")
    print(f"     mineral names: {df['recurso_natural'].value_counts().to_dict()}; min_agrupado2: {df['min_agrupado2'].value_counts().to_dict()}")
    print(f"     contraprestacion: {df['contraprestacion'].value_counts().to_dict()}")
    reg = df[df["contraprestacion"] == "REGALIAS"]
    print("\n     Mt by year x contraprestacion")
    print((df.groupby(["a_o_liquidado", "contraprestacion"])["t"].sum() / 1e6).unstack().round(3).to_string())
    print("\n     Mt by year x quarter (regalias)")
    print((reg.groupby(["a_o_liquidado", "trimestre"])["t"].sum() / 1e6).unstack().round(3).to_string())
    print("\n     Mt by year x periodo (regalias)")
    print((reg.groupby(["a_o_liquidado", "periodo_liquidado2"])["t"].sum() / 1e6).unstack().round(3).to_string())
    print("\n     Mt by department x year (regalias)")
    print((reg.groupby(["nombre_departamento", "a_o_liquidado"])["t"].sum() / 1e6).unstack().round(3).to_string())
    print("\n     Mt by mineral x year (regalias)")
    print((reg.groupby(["recurso_natural", "a_o_liquidado"])["t"].sum() / 1e6).unstack().round(3).to_string())
    print("\n     big projects (regalias, Mt by year)")
    p = (reg.groupby(["nombre_departamento", "nombre_del_proyecto", "a_o_liquidado"])["t"].sum() / 1e6).unstack()
    print(p[p.max(axis=1) > 0.5].round(3).to_string())
    print("\n     projects with royalties but zero volume (top 15 by royalties 2024)")
    z = reg[(reg["t"] == 0)].groupby(["nombre_departamento", "nombre_del_proyecto", "a_o_liquidado"])["roy"].sum().unstack()
    print((z.sort_values(z.columns[-2] if len(z.columns) > 1 else z.columns[0], ascending=False).head(15) / 1e9).round(1).to_string())
    print("\n     La Guajira 2024 rows sample")
    print(reg[(reg["nombre_departamento"] == "La Guajira") & (reg["a_o_liquidado"] == "2024")][
        ["nombre_municipio", "nombre_del_proyecto", "periodo_liquidado2", "trimestre", "volumenes_de_explotacion", "regalias_pagadas"]].head(40).to_string())
    print(f"\n     latest periods: {reg.groupby('a_o_liquidado')['periodo_liquidado2'].apply(lambda s: sorted(set(s))).tail(3).to_dict()}")


def ei():
    section("EI via OWID snapshots")
    r = get("https://api.github.com/repos/owid/etl/contents/snapshots/energy_institute")
    vers = sorted(d["name"] for d in r.json() if d["type"] == "dir") if r is not None and r.status_code == 200 else []
    print(f"     snapshot versions: {vers}")
    for v in reversed(vers[-2:]):
        r = get(f"https://api.github.com/repos/owid/etl/contents/snapshots/energy_institute/{v}")
        if r is None or r.status_code != 200:
            continue
        files = [d["name"] for d in r.json()]
        print(f"     {v}: {files}")
        for f in files:
            if not f.endswith(".dvc") or not re.search(r"statistical_review", f):
                continue
            rr = get(f"https://raw.githubusercontent.com/owid/etl/master/snapshots/energy_institute/{v}/{f}")
            if rr is None or rr.status_code != 200:
                continue
            print("     " + rr.text.replace("\n", "\n     ")[:2500])
            md5 = re.search(r"md5:\s*([0-9a-f]{32})", rr.text)
            if md5 and f.endswith(".xlsx.dvc") or (md5 and "xlsx" in f):
                h = md5.group(1)
                for u in [f"https://snapshots.owid.io/{h[:2]}/{h[2:]}", f"https://owid-snapshots.nyc3.digitaloceanspaces.com/{h[:2]}/{h[2:]}"]:
                    x = get(u)
                    if x is not None and x.status_code == 200 and len(x.content) > 100000:
                        try:
                            xl = pd.ExcelFile(io.BytesIO(x.content))
                        except Exception as e:
                            print(f"     not xlsx: {e}")
                            continue
                        print(f"     sheets: {xl.sheet_names}")
                        for s in [s for s in xl.sheet_names if re.search(r"coal", s, re.I)]:
                            df = xl.parse(s, header=None)
                            print(f"     --- {s} {df.shape}")
                            print(df.iloc[:5, [0] + list(range(df.shape[1] - 8, df.shape[1]))].to_string(max_colwidth=25))
                            for _, row in df.iterrows():
                                if re.search(r"Colombia|Brazil|Venezuela|Chile|Argentina|Peru|S\. & Cent|Mexico", str(row.iloc[0])):
                                    print(f"       {row.iloc[0]}: {row.values[-12:].tolist()}")
                        return


def brazil():
    section("Brazil EPE BEN chapter 2")
    r = get("https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes/BEN-Series-Historicas-Completas")
    if r is None:
        return
    ls = links(r.text, r"Cap.tulo 2|Cap.tulo 4", "https://www.epe.gov.br")
    for h in ls:
        u = quote(h, safe=":/")
        x = get(u)
        if x is None or x.status_code != 200:
            continue
        xl = pd.ExcelFile(io.BytesIO(x.content))
        print(f"     sheets: {xl.sheet_names}")
        for s in xl.sheet_names:
            df = xl.parse(s, header=None)
            title = " ".join(str(v) for v in df.iloc[:12, :3].values.ravel() if str(v) != "nan")[:200]
            if not re.search(r"carv|coal", title + s, re.I):
                continue
            print(f"\n     --- {s}: {title}")
            print(df.iloc[:40, [0, 1] + list(range(df.shape[1] - 6, df.shape[1]))].to_string(max_colwidth=40))


def argentina():
    section("Argentina BEN")
    for y, u in [(2025, "https://www.argentina.gob.ar/sites/default/files/balance_2025_v0_h.xlsx"),
                 (2025, "https://www.energia.gob.ar/contenidos/archivos/Reorganizacion/informacion_del_mercado/publicaciones/energia_en_gral/balances_2025/balance_2025_v0_h.xlsx"),
                 (2024, "https://www.argentina.gob.ar/sites/default/files/balance_2024_v0_h.xlsx"),
                 (2023, "https://www.argentina.gob.ar/sites/default/files/balance_2023_v0_h.xlsx"),
                 (2022, "https://www.energia.gob.ar/contenidos/archivos/Reorganizacion/informacion_del_mercado/publicaciones/energia_en_gral/balances_2022/balance_2022_V0_horizontal.xlsx"),
                 (2021, "http://www.energia.gob.ar/contenidos/archivos/Reorganizacion/informacion_del_mercado/publicaciones/energia_en_gral/balances_2021/balance_2021_V1.xlsx")]:
        r = get(u)
        if r is None or r.status_code != 200:
            continue
        try:
            xl = pd.ExcelFile(io.BytesIO(r.content))
        except Exception as e:
            print(f"     {e}")
            continue
        for s in xl.sheet_names:
            df = xl.parse(s, header=None)
            for i, row in df.iterrows():
                t = " | ".join(str(v) for v in row.values if str(v) != "nan")
                if re.search(r"carb.n mineral|kcal|tonelada|UNIDADES", t, re.I):
                    print(f"     {y} [{s} r{i}] {t[:300]}")
    r = get("https://www.argentina.gob.ar/economia/energia/planeamiento-energetico/balances-energeticos")
    if r is not None and r.status_code == 200:
        for h in links(r.text, r"\.xlsx|\.pdf|\.csv", "https://www.argentina.gob.ar")[:30]:
            print(f"     {h}")


def chile():
    section("Chile BNE")
    for u in ["http://datos.energiaabierta.cl/dataviews/238711/balance-nacional-de-energia/",
              "https://energiaabierta.cl/categorias-estadistica/balance-energetico/?_sft_etiquetas-estadistica=bne",
              "https://www.energia.gob.cl/balance-nacional-de-energia",
              "https://www.energia.gob.cl/documentos?search_api_fulltext=balance%20nacional%20de%20energia"]:
        r = get(u)
        if r is not None and r.status_code == 200:
            for h in links(r.text, r"csv|xls|download|datastream|dataview|bne|balance")[:40]:
                print(f"     {h}")
            for m in re.finditer(r"(Carb[oó]n[^<]{0,200})", r.text):
                print(f"     txt: {m.group(1)[:200]}")
                break


def peru():
    section("Peru")
    for u in ["https://www.gob.pe/institucion/minem/colecciones",
              "https://www.gob.pe/institucion/minem/informes-publicaciones",
              "https://www.gob.pe/busquedas.json?term=anuario+minero&institucion[]=minem",
              "https://www.minem.gob.pe/_estadistica.php?idSector=1&idEstadistica=13015"]:
        r = get(u)
        if r is not None and r.status_code == 200:
            for h in links(r.text, r"anuario|bolet|estad|minera|carb", "https://www.gob.pe")[:40]:
                print(f"     {h}")


if __name__ == "__main__":
    for w in sys.argv[1:] or ["anm", "ei", "brazil", "argentina", "chile", "peru"]:
        try:
            globals()[w]()
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"{w} failed: {type(e).__name__}: {e}")
