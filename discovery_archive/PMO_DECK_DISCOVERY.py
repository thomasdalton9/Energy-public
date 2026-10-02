"""Discovery probe (GitHub Actions): where can the monthly NEWAVE deck inputs - the load forecast
(SISTEMA.DAT 'mercado de energia') and the expansion schedule (EXPT/EXPH, CONFHD) - be fetched
without a SINTEGRE login? Prints status and candidate links for each source."""
import re, json, requests, sys
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data-probe/1.0", "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8"}

def get(url, timeout=45, **kw):
    try:
        r = requests.get(url, headers=UA, timeout=timeout, allow_redirects=True, **kw)
        return r.status_code, r
    except Exception as e:
        return f"ERR {type(e).__name__}: {str(e)[:120]}", None

def links(html, pat, base=""):
    out = []
    for m in re.finditer(r'href=["\']([^"\']+)["\']', html, flags=re.I):
        h = m.group(1)
        if re.search(pat, h, flags=re.I): out.append(h if h.startswith("http") else base + h)
    return sorted(set(out))

print("===== 1. CCEE price decks (expect 403 from runners)")
for u in ["https://www.ccee.org.br/precos/decks", "https://www.ccee.org.br/web/guest/precos/decks-de-preco", "https://www.ccee.org.br/api/", "https://dadosabertos.ccee.org.br/"]:
    s, r = get(u); print(u, "->", s, (len(r.text) if r is not None else ""))

print("\n===== 2. ONS: PMO pages, acervo, PEN")
for u in ["https://www.ons.org.br/paginas/energia-no-futuro/programacao-mensal-da-operacao",
          "https://www.ons.org.br/Paginas/Noticias/PMO.aspx",
          "https://www.ons.org.br/paginas/sobre-o-sin/o-sistema-em-numeros",
          "https://www.ons.org.br/AcervoDigitalDocumentosEPublicacoes/",
          "https://www.ons.org.br/paginas/energia-no-futuro/planejamento-da-operacao",
          "https://www.ons.org.br/paginas/energia-no-futuro/plano-da-operacao-energetica",
          "https://sintegre.ons.org.br/"]:
    s, r = get(u); print(u, "->", s)
    if r is not None and r.status_code == 200:
        for l in links(r.text, r"pmo|sumario|sum%C3%A1rio|carga|newave|deck|pen[-_ ]|\.pdf|\.xlsx|\.zip", "https://www.ons.org.br")[:40]: print("    ", l)

print("\n===== 3. ONS open data CKAN catalogue: datasets mentioning previsao / carga / expansao / pmo")
for q in ["previsao carga", "carga mensal", "expansao", "pmo", "newave", "programacao mensal", "cronograma", "capacidade instalada"]:
    s, r = get(f"https://dados.ons.org.br/api/3/action/package_search?q={requests.utils.quote(q)}&rows=20")
    if r is not None and r.status_code == 200:
        try:
            res = r.json()["result"]; print(f"  q='{q}': {res['count']} datasets ->", [p["name"] for p in res["results"]])
        except Exception as e: print(f"  q='{q}': parse error {e}")
    else: print(f"  q='{q}': {s}")

print("\n===== 4. EPE: load projections and quadrimestral reviews")
for u in ["https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes/projecoes-da-demanda-de-energia-eletrica",
          "https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes/revisoes-quadrimestrais-da-carga",
          "https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes/revisao-quadrimestral-das-projecoes-da-demanda-de-energia-eletrica-do-sistema-interligado-nacional",
          "https://www.epe.gov.br/pt/areas-de-atuacao/energia-eletrica/expansao-da-geracao",
          "https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes/plano-decenal-de-expansao-de-energia-pde",
          "https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes"]:
    s, r = get(u); print(u, "->", s)
    if r is not None and r.status_code == 200:
        for l in links(r.text, r"carga|demanda|quadrimestr|\.xlsx|\.xls|\.pdf|\.zip", "https://www.epe.gov.br")[:40]: print("    ", l)

print("\n===== 5. MME / CMSE and ANEEL expansion monitoring")
for u in ["https://www.gov.br/mme/pt-br/assuntos/secretarias/energia-eletrica/cmse",
          "https://www.gov.br/aneel/pt-br/centrais-de-conteudos/relatorios-e-indicadores/geracao",
          "https://dadosabertos.aneel.gov.br/api/3/action/package_search?q=expansao+geracao&rows=10",
          "https://dadosabertos.aneel.gov.br/api/3/action/package_search?q=siga&rows=10"]:
    s, r = get(u); print(u, "->", s)
    if r is not None and r.status_code == 200 and "api/3" in u:
        try: res = r.json()["result"]; print("    ", res["count"], [p["name"] for p in res["results"]])
        except Exception as e: print("     parse error", e)

print("\n===== 6. GitHub-hosted deck mirrors / tooling (public)")
for u in ["https://api.github.com/search/repositories?q=newave+deck+pmo&sort=updated", "https://api.github.com/search/repositories?q=sintegre+deck&sort=updated"]:
    s, r = get(u); print(u, "->", s)
    if r is not None and r.status_code == 200:
        try: print("    ", [(i["full_name"], i["description"][:60] if i["description"] else "") for i in r.json()["items"][:10]])
        except Exception as e: print("     parse error", e)
print("\nDONE")
