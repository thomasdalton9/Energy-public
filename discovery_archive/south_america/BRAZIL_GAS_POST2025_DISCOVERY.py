"""
Where is Brazil gas demand BY SEGMENT published after June 2025 (where
MME's historical-balance xlsx annex stops)? The Observatorio "Boletim do
Gas" dashboard redirects to a gov.br login (BRAZIL_BOLETIM_GAS_DISCOVERY.py).
Checks the remaining public routes:
  1. ANP open data "comercializacao de gas natural" - CSV columns, latest month,
     whether there's any end-use segment field.
  2. ABEGAS "estatisticas de consumo" - titles/dates of the newest posts, and
     the text of the newest one (do they still publish monthly by segment?).
  3. MME bulletin PDFs for Jul-Sep 2025 - do they exist and contain a demand-
     by-segment table (pdfplumber text around "Industrial"/"Geracao").
"""
import io
import re

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 90)


def out(*a):
    print(*a, flush=True)


out("########## 1. ANP comercializacao de gas natural ##########")
page = "https://www.gov.br/anp/pt-br/centrais-de-conteudo/dados-abertos/comercializacao-de-gas-natural"
r = requests.get(page, headers=H, timeout=T)
links = sorted(set(re.findall(r'href="([^"]+\.csv[^"]*)"', r.text, re.I)))
out(page, r.status_code, f"{len(links)} csv links")
import pandas as pd
for l in links:
    out(" ", l)
    try:
        c = requests.get(l, headers=H, timeout=T)
        txt = c.content.decode("latin-1")
        df = pd.read_csv(io.StringIO(txt), sep=None, engine="python")
        out(f"    shape {df.shape}; cols {list(df.columns)}")
        out(df.tail(4).to_string(max_colwidth=30)[:1500])
        for col in df.columns:
            if df[col].dtype == object and df[col].nunique() < 30:
                out(f"    distinct {col}: {sorted(map(str, df[col].dropna().unique()))[:30]}")
    except Exception as e:
        out("    ERR", type(e).__name__, str(e)[:200])

out("\n########## 2. ABEGAS consumption posts ##########")
for u in ["https://www.abegas.org.br/arquivos/category/consumo", "https://www.abegas.org.br/estatisticas-de-consumo"]:
    r = requests.get(u, headers=H, timeout=T)
    posts = re.findall(r'<a[^>]+href="(https://www\.abegas\.org\.br/arquivos/\d+)"[^>]*>([^<]{15,200})</a>', r.text)
    dates = re.findall(r'(\d{1,2}\s+de\s+\w+\s+de\s+20\d\d|\d{2}/\d{2}/20\d\d)', r.text)
    out(u, r.status_code, f"{len(posts)} post links; dates seen: {dates[:15]}")
    for href, title in list(dict.fromkeys(posts))[:20]:
        out("   ", href, "|", title.strip())
newest = "https://www.abegas.org.br/arquivos/99939"
r = requests.get(newest, headers=H, timeout=T)
text = re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", r.text, flags=re.S | re.I)
text = re.sub(r"\s+", " ", text)
i = text.find("Consumo de g")
out("\nnewest annual release text:", text[i:i + 2500])

out("\n########## 3. MME bulletin PDFs Jul-Sep 2025 ##########")
base = ("https://www.gov.br/mme/pt-br/assuntos/secretarias/petroleo-gas-natural-e-biocombustiveis/publicacoes-1/"
        "boletim-mensal-de-acompanhamento-da-industria-de-gas-natural/2025/")
for name in ["07-boletim-de-acompanhamento-da-industria-de-gas-natural-julho-de-2025.pdf",
             "08-boletim-de-acompanhamento-da-industria-de-gas-natural-agosto-de-2025-pdf.pdf",
             "09-boletim-de-acompanhamento-da-industria-de-gas-natural-setembro-em-elaboracao-de-2025-pdf.pdf"]:
    r = requests.get(base + name, headers=H, timeout=T)
    out(f"\n{name}: {r.status_code} {r.headers.get('content-type','')[:30]} {len(r.content)}B")
    if r.content[:4] != b"%PDF":
        continue
    import pdfplumber
    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
        out(f"  {len(pdf.pages)} pages")
        for pi, pg in enumerate(pdf.pages):
            t = pg.extract_text() or ""
            if re.search(r"Gera..o El.trica", t) and re.search(r"Industrial", t) and re.search(r"Automotivo", t):
                out(f"  -- page {pi + 1} has the segment table:")
                lines = [ln for ln in t.splitlines() if re.search(r"Industrial|Automotivo|Residencial|Comercial|Gera..o|Cogera|Outros|DEMANDA|jun|jul|ago|set", ln, re.I)]
                out("   " + "\n   ".join(lines[:30]))
                break
