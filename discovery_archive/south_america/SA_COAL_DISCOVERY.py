"""
South America coal production / exports - source discovery (run in GitHub Actions; the sources are
not reachable from the editing sandbox). Probes:
  1. datos.gov.co (Socrata) catalogue: ANM coal production / royalty datasets, columns, sample rows
  2. ANM Colombia and SIMCO pages: links to coal production statistics
  3. DANE exports page: monthly annex xlsx, coal rows (tonnes)
  4. UN Comtrade preview: Colombia HS 2701 exports (monthly)
  5. Energy Institute Statistical Review: all-data xlsx link and coal production sheets
  6. Brazil: ANM CFEM royalties open data (coal sold by month), EPE BEN
  7. Argentina: Series de Tiempo API search for coal; Peru MINEM; Chile CNE
"""
import io
import re
import sys
import time

import requests

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "es,en;q=0.8"}
S = requests.Session()
S.headers.update(H)


def get(url, **kw):
    kw.setdefault("timeout", (15, 90))
    try:
        r = S.get(url, **kw)
        print(f"  GET {url} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)} bytes")
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


def socrata():
    section("1. datos.gov.co catalogue")
    seen = {}
    for q in ["carbon", "carbón", "produccion carbon", "producción de minerales", "regalias", "ANM mineria produccion",
              "exportaciones carbon"]:
        r = get("https://api.us.socrata.com/api/catalog/v1", params={"domains": "www.datos.gov.co", "q": q, "limit": 40})
        if r is None or r.status_code != 200:
            continue
        for res in r.json().get("results", []):
            d = res["resource"]
            if d["id"] in seen:
                continue
            seen[d["id"]] = d
            print(f"  [{q}] {d['id']} | {d['name'][:90]} | {d.get('attribution')} | updated {d.get('data_updated_at')} "
                  f"| type {d.get('type')}")
    cand = [i for i, d in seen.items() if re.search(r"carb|miner|regal|produc", d["name"], re.I) and d.get("type") == "dataset"]
    for i in cand[:25]:
        d = seen[i]
        print(f"\n  --- {i} {d['name']}")
        print(f"      cols: {d.get('columns_name', [])[:30]}")
        r = get(f"https://www.datos.gov.co/resource/{i}.json", params={"$limit": 3})
        if r is not None and r.status_code == 200:
            for row in r.json()[:3]:
                print(f"      {str(row)[:400]}")
        r = get(f"https://www.datos.gov.co/resource/{i}.json", params={"$select": "count(*)"})
        if r is not None and r.status_code == 200:
            print(f"      rows: {r.text[:100]}")


def anm_colombia():
    section("2. ANM Colombia / SIMCO")
    for u in ["https://www.anm.gov.co/", "https://www.anm.gov.co/?q=informacion-estadistica",
              "https://www.anm.gov.co/?q=produccion-de-minerales", "https://www.anm.gov.co/datos-abiertos",
              "https://www.anm.gov.co/?q=datos_abiertos",
              "https://www.simco.gov.co/", "https://www1.upme.gov.co/simco/Cifras-Sectoriales/Paginas/carbon.aspx",
              "https://www1.upme.gov.co/simco/Cifras-Sectoriales/Paginas/Informacion-estadistica-minera.aspx",
              "https://www.anm.gov.co/estadisticas-mineras", "https://www.anm.gov.co/?q=estadisticas-mineras"]:
        r = get(u)
        if r is None or r.status_code != 200:
            continue
        for h in links(r.text, r"carb|produc|estad|xls|regal|cifra", "https://www.anm.gov.co" if "anm" in u else "")[:60]:
            print(f"     {h}")


def dane():
    section("3. DANE exports")
    for u in ["https://www.dane.gov.co/index.php/estadisticas-por-tema/comercio-internacional/exportaciones",
              "https://www.dane.gov.co/index.php/estadisticas-por-tema/comercio-internacional/exportaciones/exportaciones-historicos"]:
        r = get(u)
        if r is None or r.status_code != 200:
            continue
        ls = links(r.text, r"\.xls", "https://www.dane.gov.co")
        for h in ls[:60]:
            print(f"     {h}")
        anex = [h for h in ls if re.search(r"anex", h, re.I)]
        if anex:
            try_xlsx(anex[0], r"carb|hulla|2701|tonel")
            if len(anex) > 1:
                try_xlsx(anex[1], r"carb|hulla|2701|tonel")


def try_xlsx(url, pat, max_rows=60):
    r = get(url)
    if r is None or r.status_code != 200:
        return None
    import pandas as pd
    try:
        xl = pd.ExcelFile(io.BytesIO(r.content))
    except Exception as e:
        print(f"     not excel: {e}")
        return None
    print(f"     sheets: {xl.sheet_names[:60]}")
    n = 0
    for s in xl.sheet_names:
        try:
            df = xl.parse(s, header=None)
        except Exception:
            continue
        for i, row in df.iterrows():
            txt = " | ".join(str(v) for v in row.values if str(v) != "nan")
            if re.search(pat, txt, re.I):
                print(f"     [{s} r{i}] {txt[:300]}")
                n += 1
                if n > max_rows:
                    return xl
    return xl


def comtrade():
    section("4. UN Comtrade Colombia HS 2701")
    for period in ["202401", "202506", "202507"]:
        r = get("https://comtradeapi.un.org/public/v1/preview/C/M/HS",
                params={"reporterCode": 170, "period": period, "cmdCode": "2701", "flowCode": "X"})
        if r is not None and r.status_code == 200:
            data = r.json().get("data", [])
            w = [d for d in data if d.get("partnerCode") == 0]
            print(f"     {period}: {len(data)} rows; world: {[(d.get('netWgt'), d.get('primaryValue')) for d in w]}")
        time.sleep(2)


def energy_institute():
    section("5. Energy Institute Statistical Review")
    for u in ["https://www.energyinst.org/statistical-review/resources-and-data-downloads",
              "https://www.energyinst.org/statistical-review"]:
        r = get(u)
        if r is None or r.status_code != 200:
            continue
        ls = links(r.text, r"\.xlsx|\.csv|excel_doc|__data/assets", "https://www.energyinst.org")
        for h in ls[:40]:
            print(f"     {h}")
        xl = [h for h in ls if re.search(r"all.?data|consolidated|panel", h, re.I)]
        for h in xl[:3]:
            if h.endswith(".xlsx") or "excel" in h:
                x = try_xlsx(h, r"^Colombia|^Brazil|^Venezuela|^Chile|^Argentina|^Peru|Central & S", max_rows=0)
                if x is not None:
                    import pandas as pd
                    for s in [s for s in x.sheet_names if re.search(r"coal", s, re.I)]:
                        df = x.parse(s, header=None)
                        print(f"     --- {s}: shape {df.shape}; header rows: {df.iloc[2:4, :3].values.tolist()}")
                        for i, row in df.iterrows():
                            if str(row.iloc[0]).strip() in ("Colombia", "Brazil", "Venezuela", "Chile", "Argentina", "Peru",
                                                            "Total S. & Cent. America", "Other S. & Cent. America"):
                                print(f"       {row.iloc[0]}: {[v for v in row.values[-8:]]}")
                        if df.shape[0] > 2:
                            print(f"       years row: {df.iloc[2].values[-8:].tolist()}")
                break


def brazil():
    section("6. Brazil")
    for u in ["https://app.anm.gov.br/dadosabertos/ARRECADACAO/CFEM_Arrecadacao.csv",
              "https://app.anm.gov.br/dadosabertos/ARRECADACAO/"]:
        try:
            with S.get(u, stream=True, timeout=(15, 60)) as r:
                print(f"  GET {u} -> {r.status_code} {r.headers.get('content-type')} len {r.headers.get('content-length')}")
                if r.status_code == 200:
                    buf = b""
                    for chunk in r.iter_content(65536):
                        buf += chunk
                        if len(buf) > 200000:
                            break
                    txt = buf.decode("latin-1", errors="replace")
                    lines = txt.splitlines()
                    for ln in lines[:4]:
                        print(f"     {ln[:400]}")
                    coal = [ln for ln in lines if re.search(r"CARV", ln)]
                    print(f"     coal lines in first 200kB: {len(coal)}")
                    for ln in coal[:5]:
                        print(f"     {ln[:400]}")
                    if "<a" in txt:
                        for h in links(txt, r"\.csv|\.zip")[:40]:
                            print(f"     {h}")
        except Exception as e:
            print(f"  {u}: {type(e).__name__} {str(e)[:120]}")
    for u in ["https://dados.gov.br/api/publico/conjuntos-dados?isPrivado=false&nomeConjuntoDados=cfem",
              "https://www.gov.br/anm/pt-br/centrais-de-conteudo/publicacoes/serie-estatisticas-e-economia-mineral/anuario-mineral",
              "https://www.gov.br/anm/pt-br/assuntos/economia-mineral/publicacoes/anuario-mineral",
              "https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes/balanco-energetico-nacional-ben",
              "https://www.carvaomineral.com.br/"]:
        r = get(u)
        if r is None or r.status_code != 200:
            continue
        for h in links(r.text, r"anuario|anu%C3%A1rio|carv|xls|ben|estat|csv", "")[:40]:
            print(f"     {h}")


def others():
    section("7. Argentina / Peru / Chile")
    r = get("https://apis.datos.gob.ar/series/api/search/", params={"q": "carbon", "limit": 30})
    if r is not None and r.status_code == 200:
        for d in r.json().get("data", []):
            f = d.get("field", {})
            print(f"     {f.get('id')} | {f.get('description')} | {f.get('units')} | {d.get('dataset', {}).get('title')} "
                  f"| {f.get('time_index_start')}..{f.get('time_index_end')} {f.get('frequency')}")
    for u in ["https://www.gob.pe/institucion/minem/colecciones/16233-anuario-minero",
              "https://www.gob.pe/busquedas?term=anuario%20minero&institucion=minem",
              "https://www.gob.pe/busquedas?term=carbon&institucion=minem",
              "https://www.minem.gob.pe/_estadistica.php?idSector=1&idEstadistica=12000",
              "https://www.energiaabierta.cl/?s=carbon",
              "https://www.cne.cl/estadisticas/energia/",
              "https://www.cne.cl/normativas/energia/balance-energetico/",
              "https://energiaabierta.cl/visualizaciones/balance-de-energia/",
              "https://www.argentina.gob.ar/economia/energia/hidrocarburos/balances-energeticos",
              "http://www.ycrt.gob.ar/"]:
        r = get(u)
        if r is None or r.status_code != 200:
            continue
        for h in links(r.text, r"anuario|carb|xls|balance|BEN|produc|csv|estad")[:40]:
            print(f"     {h}")


if __name__ == "__main__":
    which = sys.argv[1:] or ["socrata", "anm_colombia", "dane", "comtrade", "energy_institute", "brazil", "others"]
    for w in which:
        try:
            globals()[w]()
        except Exception as e:
            print(f"{w} failed: {type(e).__name__}: {e}")
