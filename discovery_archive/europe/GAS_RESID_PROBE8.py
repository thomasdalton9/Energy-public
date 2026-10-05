"""Probe 8 (gas residual, Spain): try pdfplumber extraction modes on Enagas bulletins with letter-spaced tables; print the BARCELONA..Total rows of the regasification page."""
import io, os, re, sys, json
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "europe"))
import pdfplumber
import GAS_TSO_SOUTHEAST_DAILY as G
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results", "residual")
h = {"User-Agent": G.UA}
files = {}
for y, m in [(2022, 3), (2023, 4), (2023, 7), (2024, 2)]:
    r = G.get(G.ENAGAS_PAGE, params={"category": "", "month": m, "year": y}, headers=h)
    for x in re.findall(r'href="(/content/dam[^"]+\.pdf)"', r.text):
        files[x.split("/")[-1]] = x
want = ["Monthly_Bulletin_June23.pdf", "Boletín%20Estadístico_ene24_ingles.pdf", "Monthly-Bulletin-Gas-March-2023.pdf", "Monthly-Bulletin-Gas-february-2022.pdf"]
rep = {}
for w in want:
    c = G.get("https://www.enagas.es" + files[w], headers=h).content
    with pdfplumber.open(io.BytesIO(c)) as pdf:
        pg = next((p for p in pdf.pages if re.search(r"SHIPS UNLOADED", p.extract_text() or "")), None)
        if pg is None: continue
        out = {}
        out["xtol8"] = pg.extract_text(x_tolerance=8)[:1800]
        out["layout"] = pg.extract_text(layout=True, x_tolerance=8)[:2500]
        words = pg.extract_words(x_tolerance=8, y_tolerance=4)
        rows = {}
        for wd in words: rows.setdefault(round(wd["top"] / 4), []).append(wd)
        out["rows"] = [" ".join(x["text"] for x in sorted(v, key=lambda z: z["x0"])) for k, v in sorted(rows.items())][:60]
        rep[w] = out
    print(w, flush=True)
json.dump(rep, open(os.path.join(OUT, "enagas_trucks_modes.json"), "w"), ensure_ascii=False, indent=1)
