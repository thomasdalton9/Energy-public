"""Probe 6 (gas residual, Spain): inspect a few Enagas bulletins that lack a parsed LNG-truck figure: per page text length, keyword lines, and the images on the page."""
import io, os, re, sys, json
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "europe"))
import pdfplumber
import GAS_TSO_SOUTHEAST_DAILY as G
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results", "residual")
h = {"User-Agent": G.UA}
want = ["Monthly_Bulletin_June23.pdf", "Boletín%20Estadístico_ene24_ingles.pdf", "Monthly-Bulletin-Gas-March-2023.pdf", "Monthly-Bulletin-Gas-february-2022.pdf", "Boletín%20Estadístico_oct24_ingles.pdf"]
files = {}
for y, m in [(2022, 3), (2023, 4), (2023, 7), (2024, 2), (2024, 11)]:
    r = G.get(G.ENAGAS_PAGE, params={"category": "", "month": m, "year": y}, headers=h)
    for x in re.findall(r'href="(/content/dam[^"]+\.pdf)"', r.text):
        files[x.split("/")[-1]] = x
rep = {}
for w in want:
    if w not in files: print("no", w); continue
    c = G.get("https://www.enagas.es" + files[w], headers=h).content
    with pdfplumber.open(io.BytesIO(c)) as pdf:
        info = []
        for i, pg in enumerate(pdf.pages):
            t = pg.extract_text() or ""
            kw = [ln[:160] for ln in t.split("\n") if re.search(r"truck|cisterna|regasif|ships|loaded|unload|Total", ln, re.I)][:8]
            info.append({"p": i, "chars": len(t), "images": len(pg.images), "kw": kw})
        rep[w] = info
    print(w, len(rep[w]), flush=True)
json.dump(rep, open(os.path.join(OUT, "enagas_trucks_pages.json"), "w"), ensure_ascii=False, indent=1)
