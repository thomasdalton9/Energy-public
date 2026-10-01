"""
Round 4: the ANP pipeline-movement files for Jan-May 2022 parse to zero
"Volume Realizado" rows in BRAZIL_GAS.py. Print their links, raw header
lines, variable names and a few raw rows, next to a 2022-06 file that parses.
"""
import io
import re

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
PAGE = ("https://www.gov.br/anp/pt-br/centrais-de-conteudo/dados-abertos/"
        "dados-consolidados-movimentacao-de-gas-natural-em-gasodutos-de-transporte")
html = requests.get(PAGE, headers=H, timeout=(10, 90)).text
links = sorted(set(re.findall(r'href="([^"]+/2022/[^"]+)"', html)))
print("2022 links:", *links, sep="\n  ", flush=True)
for u in links:
    if not re.search(r"janeiro|fevereiro|maio|junho", u):
        continue
    c = requests.get(u, headers=H, timeout=(10, 120)).content
    try:
        t = c.decode("utf-8-sig")
    except UnicodeDecodeError:
        t = c.decode("latin-1")
    print(f"\n=== {u.rsplit('/', 1)[-1]} {len(c)}B, first bytes {c[:8]!r}", flush=True)
    for ln in t.splitlines()[:4]:
        print("  RAW:", ln[:700], flush=True)
    sep = ";" if t[:3000].count(";") > t[:3000].count(",") else ","
    df = pd.read_csv(io.StringIO(t), sep=sep, dtype=str, on_bad_lines="skip")
    print("  sep", repr(sep), "shape", df.shape, flush=True)
    print("  cols", list(df.columns)[:16], "...", list(df.columns)[-3:], flush=True)
    var = [c for c in df.columns if "vari" in c.lower()][0]
    print("  variables:", df[var].value_counts().head(40).to_dict(), flush=True)
    for col in df.columns[:12]:
        if df[col].nunique() <= 15:
            print(f"  distinct {col!r}: {df[col].value_counts().to_dict()}", flush=True)
print("\ndone")
