"""
One-off discovery: where does Colombia gas SUPPLY (production by field,
SPEC LNG regasification, Venezuela imports) live?
1. BMC Gestor del Mercado monthly report PDFs: dump pages that mention
   supply/production/injection keywords, with their tables.
2. ANH / MinEnergia / SPEC pages: status and spreadsheet links.
Manual-only (discovery_archive/workflows/colombia_gas_supply_discovery.yml).
"""
import io
import re
import sys

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 120)
REPORTS = sys.argv[1:] or [
    "https://www.bmcbec.com.co/sites/default/files/2026-09/Informe%20Mensual%202026%20Agosto.pdf",
    "https://www.bmcbec.com.co/sites/default/files/2023-02/1_Informe%20Mensual%202023%20Enero.pdf",
    "https://www.bmcbec.com.co/sites/default/files/2021-02/Informe%20Mensual%20Enero%202021.pdf",
]
KW = re.compile(r"oferta|producci|inyecc|campo|regasific|SPEC|importa|Cusiana|Cupiagua|Ballena|Chuchupa|Jobo|"
                r"Venezuela|suministro|fuente", re.I)


def out(*a):
    print(*a, flush=True)


def dump_pdf(u, full_first):
    import pdfplumber
    out(f"\n==================== {u}")
    try:
        c = requests.get(u, headers=H, timeout=T).content
    except Exception as e:
        out("ERR", e)
        return
    if c[:4] != b"%PDF":
        out("not a pdf", c[:80])
        return
    with pdfplumber.open(io.BytesIO(c)) as pdf:
        out(f"{len(pdf.pages)} pages")
        for i, page in enumerate(pdf.pages):
            t = page.extract_text() or ""
            first = t.strip().splitlines()[:2]
            hits = sorted(set(m.group(0).lower() for m in KW.finditer(t)))
            out(f"-- p{i + 1}: {first} hits={hits}")
            if not hits:
                continue
            nl = 90 if full_first else 45
            for ln in t.splitlines()[:nl]:
                out("    |", ln[:170])
            try:
                tabs = page.extract_tables()
            except Exception:
                tabs = []
            for j, tb in enumerate(tabs[:4]):
                out(f"    table {j}: {len(tb)} rows")
                for row in tb[:30]:
                    out("     T", [re.sub(r"\s+", " ", str(x))[:16] if x else "" for x in row][:16])


def probe(u):
    try:
        r = requests.get(u, headers=H, timeout=T)
        links = re.findall(r'href="([^"]+\.(?:xlsx?|csv|pdf|zip)[^"]*)"', r.text, re.I)
        out(f"\n{r.status_code} {len(r.content)}B {u} -> {r.url}")
        n = 0
        for ln in links:
            if re.search(r"gas|produc|fiscal|regas|import", ln, re.I):
                out("   ", ln[:200])
                n += 1
                if n > 40:
                    break
    except Exception as e:
        out(f"\nERR {u}: {type(e).__name__} {str(e)[:150]}")


if __name__ == "__main__":
    for k, u in enumerate(REPORTS):
        dump_pdf(u, k == 0)
    for u in [
        "https://www.anh.gov.co/es/operaciones-y-regal%C3%ADas/sistemas-integrados-operaciones/estad%C3%ADsticas-de-producci%C3%B3n/",
        "https://www.anh.gov.co/es/operaciones-y-regal%C3%ADas/sistemas-integrados-operaciones/estadisticas-de-produccion/",
        "https://www.minenergia.gov.co/es/servicio-al-ciudadano/estadisticas/hidrocarburos/",
        "https://www.speclng.com/",
        "https://www.bmcbec.com.co/informes/informes-anuales",
    ]:
        probe(u)
