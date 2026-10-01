"""
The MME Sep-2025 gas bulletin (the last public one) has a demand-by-segment
table on page 3 headed "Malha Interligada" (interconnected grid only - its
power-generation figures are far below the national xlsx). Dump the header
line and DEMANDA TOTAL / Geracao Eletrica lines of EVERY table in the PDF so
we can see whether a national ("Brasil") table exists and which one matches
historico-balanco-boletim.xlsx for Jun-2025 (Industrial 39.25, Geracao
Eletrica 19.01, DEMANDA TOTAL 66.69).
"""
import io
import re

import pdfplumber
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
URL = ("https://www.gov.br/mme/pt-br/assuntos/secretarias/petroleo-gas-natural-e-biocombustiveis/publicacoes-1/"
       "boletim-mensal-de-acompanhamento-da-industria-de-gas-natural/2025/"
       "09-boletim-de-acompanhamento-da-industria-de-gas-natural-setembro-em-elaboracao-de-2025-pdf.pdf")
pdf_bytes = requests.get(URL, headers=H, timeout=(10, 120)).content
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    for i, page in enumerate(pdf.pages):
        t = page.extract_text() or ""
        lines = t.splitlines()
        hits = [j for j, ln in enumerate(lines) if re.search(r"DEMANDA TOTAL|OFERTA TOTAL|Gera..o El.trica|Produ..o Nacional|Importa", ln, re.I)]
        if not hits:
            continue
        print(f"\n===== page {i + 1} =====", flush=True)
        heads = [ln for ln in lines if re.search(r"(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)-\d\d", ln)]
        for h in heads[:4]:
            print("  HEAD:", h[:220], flush=True)
        for ln in lines:
            if re.search(r"^(Brasil|Malha|Sistema|Isolad|Nordeste|Sudeste|Sul|Norte|Centro)", ln.strip(), re.I):
                print("  SECTION:", ln[:200], flush=True)
        for j in hits:
            print("  ROW:", lines[j][:240], flush=True)
