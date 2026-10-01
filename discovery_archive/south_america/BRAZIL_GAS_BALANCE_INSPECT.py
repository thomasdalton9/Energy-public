"""
Download MME's historical gas balance annex (historico-balanco-boletim.xlsx,
linked from the monthly gas bulletin's "anexos" page) - SA_GAS_DEMAND_
DISCOVERY2.py only got the /view HTML wrapper - and dump its structure:
sheet names, shapes, header rows, and every row whose label mentions a
demand segment, to see whether it carries consumption by segment
(industrial, automotive, residential, commercial, power, cogeneration)
monthly from 2021.
"""
import io

import pandas as pd
import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
BASE = ("https://www.gov.br/mme/pt-br/assuntos/secretarias/petroleo-gas-natural-e-biocombustiveis/publicacoes-1/"
        "boletim-mensal-de-acompanhamento-da-industria-de-gas-natural/anexos/historico-balanco-boletim.xlsx")
pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 40)

content = None
for url in [BASE, BASE + "/@@download/file", BASE + "/at_download/file"]:
    r = requests.get(url, headers=HEADERS, timeout=(10, 90))
    print(f"GET {url} -> {r.status_code} {r.headers.get('content-type', '')[:60]} {len(r.content)}B", flush=True)
    if r.ok and r.content[:2] == b"PK":
        content = r.content
        break
if content is None:
    raise SystemExit("no xlsx body found")

xl = pd.ExcelFile(io.BytesIO(content))
print("sheets:", xl.sheet_names, flush=True)
seg = "consumo|demanda|segment|industri|automot|gnv|residen|comerci|termel|gera|cogera|refin|fertil|outros|total"
for s in xl.sheet_names:
    df = xl.parse(s, header=None)
    print(f"\n==================== sheet {s!r} shape {df.shape}", flush=True)
    print(df.head(12).to_string(max_colwidth=22)[:5000], flush=True)
    lab = df.iloc[:, :3].fillna("").astype(str).fillna("").agg(" | ".join, axis=1)
    hits = df[lab.str.contains(seg, case=False, regex=True)]
    print(f"--- {len(hits)} rows with segment-like labels (first 3 cols + last 6 cols):", flush=True)
    cols = list(df.columns[:3]) + list(df.columns[-6:])
    print(hits[cols].head(60).to_string(max_colwidth=40)[:8000], flush=True)
