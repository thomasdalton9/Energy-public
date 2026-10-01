"""Print the fuel-imports section of older CNE monthly reports (2021-mid 2025), whose layout
CHILE_GAS_IMPORTS.py doesn't parse yet. Not reachable from the editing sandbox."""
import io, re, requests, pdfplumber
H = {"User-Agent": "Mozilla/5.0"}
for y, m in [(2021, 6), (2023, 6), (2025, 6)]:
    u = f"https://www.cne.cl/wp-content/uploads/{y}/{m:02d}/RMensual_v{y}{m:02d}.pdf"
    c = requests.get(u, headers=H, timeout=(10, 120)).content
    print("\n==========", u, flush=True)
    with pdfplumber.open(io.BytesIO(c)) as pdf:
        for i, p in enumerate(pdf.pages):
            t = p.extract_text() or ""
            if re.search(r"Importaci", t, re.I) and re.search(r"gas", t, re.I):
                print(f"--- page {i+1}")
                for ln in t.splitlines():
                    if re.search(r"gas|GNL|import|miles|ton|corresponde|Combustible|Carb|Crudo|Total", ln, re.I):
                        print("   ", ln[:200])
                for tb in p.extract_tables()[:3]:
                    for row in tb[:15]:
                        print("    T", [str(x)[:14] if x else "" for x in row][:10])
