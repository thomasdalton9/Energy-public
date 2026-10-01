"""Inspect older CNE monthly reports (2021-mid 2025), whose fuel-imports layout
CHILE_GAS_IMPORTS.py doesn't parse yet, and look for the report missing at the
usual URL (RMensual_v202605). Prints every page line that mentions gas, imports,
LNG or origin countries, plus any tables on those pages.
Not reachable from the editing sandbox; runs in GitHub Actions."""
import io, re, requests, pdfplumber

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
KEY = re.compile(r"gas|GNL|import|miles de|corresponde|Combustible|Argentina|Trinidad|Estados Unidos|Aduana|COMEX", re.I)


def get(u):
    try:
        r = requests.get(u, headers=H, timeout=(10, 120))
        return r.status_code, r.content
    except requests.RequestException as e:
        return str(e)[:80], b""


def dump(y, m, full_pages=()):
    u = f"https://www.cne.cl/wp-content/uploads/{y}/{m:02d}/RMensual_v{y}{m:02d}.pdf"
    st, c = get(u)
    print("\n==========", u, st, len(c), flush=True)
    if c[:4] != b"%PDF":
        return
    with pdfplumber.open(io.BytesIO(c)) as pdf:
        print("pages", len(pdf.pages))
        for i, p in enumerate(pdf.pages):
            t = p.extract_text() or ""
            lines = t.splitlines()
            print(f"--- page {i+1}: {len(t)} chars; head: {' | '.join(lines[:2])[:150]}")
            if (i + 1) in full_pages:
                for ln in lines:
                    print("    F", ln[:220])
                continue
            hits = [ln for ln in lines if KEY.search(ln)]
            for ln in hits[:40]:
                print("    ", ln[:220])
            if re.search(r"Importaci", t, re.I):
                for tb in p.extract_tables()[:4]:
                    for row in tb[:20]:
                        print("    T", [str(x).replace("\n", " ")[:18] if x else "" for x in row][:10])


# old layout across the range, plus the first working (new-layout) report for comparison
for y, m in [(2021, 3), (2022, 6), (2024, 6), (2025, 9), (2025, 10)]:
    dump(y, m)

# the report for data month Mar 2026 was not at the usual URL: search the WordPress media library
print("\n========== looking for RMensual_v202605")
for q in ["RMensual_v202605", "RMensual v202605", "RMensual_v2026", "Reporte Mensual Mayo 2026"]:
    st, c = get(f"https://www.cne.cl/wp-json/wp/v2/media?search={q}&per_page=30")
    print("media search", q, st, len(c))
    for s in re.findall(r'"source_url":"([^"]+)"', c.decode("utf-8", "replace")):
        print("    ", s.replace("\\/", "/"))
for folder in ["2026/05", "2026/06", "2026/04"]:
    for name in ["RMensual_v202605.pdf", "RMensual_v202605-1.pdf", "RMensual_v202605-2.pdf", "RMensual_V202605.pdf",
                 "RMensual_v202605_.pdf", "Rmensual_v202605.pdf", "RMensual_v20265.pdf"]:
        st, c = get(f"https://www.cne.cl/wp-content/uploads/{folder}/{name}")
        print("   try", folder, name, st, c[:4])
