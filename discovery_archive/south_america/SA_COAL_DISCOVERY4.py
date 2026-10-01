"""
South America coal - discovery round 4:
  - Brazil EPE BEN chapter 2 (single sheet 'Capitulo 2'): coal production rows + year header
  - Argentina BEN 2022 / 2025 (www.energia.gob.ar fails TLS chain verification -> plain-http mirror check),
    methodological document for the coal calorific value used to express coal in toe
  - Chile: Sernageomin "Anuario de la Mineria de Chile" (coal production in tonnes), Cochilco yearbook
  - Peru: gob.pe search JSON for MINEM "Anuario Minero" / "Boletin Estadistico Minero"
"""
import html
import io
import json
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


def get(url, **kw):
    kw.setdefault("timeout", (15, 120))
    try:
        r = S.get(url, **kw)
        print(f"  GET {r.url[:220]} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)} bytes")
        return r
    except Exception as e:
        print(f"  GET {url[:220]} -> {type(e).__name__}: {str(e)[:400]}")
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


def pdf_grep(content, pat, maxhits=40):
    import pdfplumber
    n = 0
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        print(f"     pages: {len(pdf.pages)}")
        for i, p in enumerate(pdf.pages):
            t = p.extract_text() or ""
            for ln in t.splitlines():
                if re.search(pat, ln, re.I):
                    print(f"     p{i + 1}: {ln[:250]}")
                    n += 1
                    if n >= maxhits:
                        return


def brazil():
    section("Brazil EPE BEN chapter 2")
    r = get("https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes/BEN-Series-Historicas-Completas")
    ls = links(r.text, r"Cap.tulo 2", "https://www.epe.gov.br")
    x = get(quote(ls[0], safe=":/"))
    df = pd.read_excel(io.BytesIO(x.content), header=None)
    print(f"     shape {df.shape}")
    for i, row in df.iterrows():
        t = " | ".join(str(v) for v in row.values if str(v) != "nan")
        if re.search(r"carv|coal|tabela|table|10.? ?t|unidade|unit", t, re.I) and len(t) < 2000:
            print(f"     r{i}: {t[:330]}")
    # year header rows
    for i, row in df.iterrows():
        vals = [v for v in row.values if isinstance(v, (int, float)) and 1970 <= v <= 2030]
        if len(vals) > 20:
            print(f"     year row r{i}: first {vals[:3]} last {vals[-3:]} (cols {list(row.values).index(vals[0])}..)")
            break


def argentina():
    section("Argentina BEN")
    for u in ["https://www.energia.gob.ar/contenidos/archivos/Reorganizacion/informacion_del_mercado/publicaciones/energia_en_gral/balances_2025/balance_2025_v0_h.xlsx",
              "http://www.energia.gob.ar/contenidos/archivos/Reorganizacion/informacion_del_mercado/publicaciones/energia_en_gral/balances_2025/balance_2025_v0_h.xlsx",
              "http://www.energia.gob.ar/contenidos/archivos/Reorganizacion/informacion_del_mercado/publicaciones/energia_en_gral/balances_2022/balance_2022_V0_horizontal.xlsx"]:
        r = get(u)
        if r is None or r.status_code != 200:
            continue
        df = pd.read_excel(io.BytesIO(r.content), header=None)
        for i, row in df.iterrows():
            t = " | ".join(str(v) for v in row.values if str(v) != "nan")
            if re.search(r"carb.n mineral|UNIDADES", t, re.I):
                print(f"     r{i}: {t[:300]}")
    for u in ["http://www.energia.gob.ar/contenidos/archivos/Reorganizacion/informacion_del_mercado/publicaciones/energia_en_gral/balances_2016/documento-metodologico-balance-energetico-nacional-final-2015.pdf",
              "http://www.energia.gob.ar/contenidos/archivos/Reorganizacion/informacion_del_mercado/publicaciones/energia_en_gral/balances_2021/sintesisbalancesenergeticos2021v1.pdf"]:
        r = get(u)
        if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
            pdf_grep(r.content, r"carb[oó]n mineral|kcal/kg|poder calor|Rio Turbio|R[ií]o Turbio")


def chile():
    section("Chile")
    for u in ["https://www.sernageomin.cl/anuario-de-la-mineria-de-chile/",
              "https://www.sernageomin.cl/anuario-de-mineria/",
              "https://www.cochilco.cl/web/anuario-de-estadisticas-del-cobre-y-otros-minerales/",
              "https://www.cochilco.cl/web/estadisticas/"]:
        r = get(u)
        if r is not None and r.status_code == 200:
            for h in links(r.text, r"anuario|\.pdf|\.xlsx?")[:40]:
                print(f"     {h}")


def peru():
    section("Peru")
    for term in ["anuario minero", "boletin estadistico minero", "produccion minera no metalica"]:
        r = get("https://www.gob.pe/busquedas.json", params={"term": term, "institucion[]": "minem"})
        if r is None or r.status_code != 200:
            continue
        try:
            d = r.json()
        except Exception:
            print(f"     not json: {r.text[:200]}")
            continue
        items = d.get("data", {}).get("attributes", {}).get("results") if isinstance(d.get("data"), dict) else None
        txt = json.dumps(d)[:200000]
        urls = sorted(set(re.findall(r'(/institucion/minem/(?:informes-publicaciones|colecciones)/[^"\\?#]+)', txt)))
        print(f"     [{term}] keys {list(d.keys())[:10]}; {len(urls)} urls")
        for u in urls[:40]:
            print(f"       {u}")
    r = get("https://www.gob.pe/institucion/minem/colecciones/1452-anuario-minero")
    for u in ["https://www.gob.pe/institucion/minem/colecciones/1452-anuario-minero",
              "https://www.gob.pe/institucion/minem/colecciones/1453-boletin-estadistico-minero"]:
        r = get(u)
        if r is not None and r.status_code == 200:
            for h in links(r.text, r"informes-publicaciones|\.xlsx?|\.pdf", "https://www.gob.pe")[:20]:
                print(f"     {h}")


if __name__ == "__main__":
    for w in sys.argv[1:] or ["brazil", "argentina", "chile", "peru"]:
        try:
            globals()[w]()
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"{w} failed: {type(e).__name__}: {e}")
