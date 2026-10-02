"""Discovery step 2 (GitHub Actions): (a) every file on EPE's 'revisoes quadrimestrais da carga' page,
with the layout of one monthly-forecast spreadsheet; (b) ANEEL open-data RALIE / capacity-addition resources
and a sample of their columns; (c) ONS 'carga-mensal' resources."""
import re, io, html, json, requests, pandas as pd
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data-probe/1.0"}
def get(url, **kw):
    r = requests.get(url, headers=UA, timeout=90, **kw); return r

print("===== A. EPE quadrimestral load reviews: all files")
r = get("https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes/revisoes-quadrimestrais-da-carga")
page = html.unescape(r.text)
files = sorted(set(m.group(1) for m in re.finditer(r'href="([^"]+PublicacoesArquivos[^"]+)"', page)))
print(len(files), "files")
for f in files: print("  ", f)
xls = [f for f in files if re.search(r"\.xlsx?$", f, re.I)]
print("\nspreadsheets:", len(xls))
for f in xls: print("  ", f)
print("\n--- layout of the monthly-forecast spreadsheets (first 3 found with 'mens' or 'Previs' in the name)")
shown = 0
for f in xls:
    if not re.search(r"mens|previs", f, re.I): continue
    url = "https://www.epe.gov.br" + f if f.startswith("/") else f
    try:
        b = get(url).content; x = pd.ExcelFile(io.BytesIO(b)); print("\n", url, "sheets:", x.sheet_names)
        for s in x.sheet_names[:4]:
            df = x.parse(s, header=None); print(f"  sheet '{s}' shape {df.shape}"); print(df.head(12).to_string()[:2500])
        shown += 1
    except Exception as e: print("  failed:", url, e)
    if shown >= 3: break

print("\n===== B. ANEEL open data: RALIE and capacity additions")
for name in ["ralie-relatorio-de-acompanhamento-da-expansao-da-oferta-de-geracao-de-energia-eletrica", "acrescimo-da-potencia-instalada", "liberacao-para-operacao-comercial-de-empreendimentos-de-geracao"]:
    r = get(f"https://dadosabertos.aneel.gov.br/api/3/action/package_show?id={name}")
    try:
        p = r.json()["result"]; print("\n", name, "| updated", p.get("metadata_modified"), "| notes:", (p.get("notes") or "")[:300].replace("\n", " "))
        for res in p["resources"]: print("   ", res.get("format"), res.get("name"), res.get("url"), res.get("last_modified"))
        csvs = [res for res in p["resources"] if (res.get("format") or "").upper() == "CSV"]
        if csvs:
            u = csvs[0]["url"]; b = get(u).content[:400000]
            for enc in ("latin-1", "utf-8"):
                try: df = pd.read_csv(io.BytesIO(b), sep=";", encoding=enc, on_bad_lines="skip", nrows=5); break
                except Exception as e: df = None
            if df is not None: print("    columns:", list(df.columns)); print(df.head(3).to_string()[:2000])
    except Exception as e: print(name, "failed:", e, r.text[:200])

print("\n===== C. ONS open data: carga-mensal")
r = get("https://dados.ons.org.br/api/3/action/package_show?id=carga-mensal")
try:
    p = r.json()["result"]; print((p.get("notes") or "")[:300].replace("\n", " "))
    for res in p["resources"][:8]: print("   ", res.get("format"), res.get("name"), res.get("url"))
except Exception as e: print("failed", e)
print("\nDONE")
