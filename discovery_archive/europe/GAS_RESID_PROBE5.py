"""Probe 5 (gas residual, Spain): Enagas bulletins whose LNG-truck figure the parser misses - dump the text of the regasification-plants page (and a no-space version)."""
import io, os, re, sys, json
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "europe"))
import pandas as pd, pdfplumber, requests
import GAS_TSO_SOUTHEAST_DAILY as G
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results", "residual")
h = {"User-Agent": G.UA}
files = {}
for y in range(2021, 2027):
    for m in range(1, 13):
        try:
            r = G.get(G.ENAGAS_PAGE, params={"category": "", "month": m, "year": y}, headers=h)
            for x in re.findall(r'href="(/content/dam[^"]+\.pdf)"', r.text):
                files.setdefault(x, (y, m))
        except Exception as e:
            pass
print("bulletins", len(files), flush=True)
rep = {}
for f, ym in sorted(files.items(), key=lambda kv: kv[1]):
    try:
        c = G.get("https://www.enagas.es" + f, headers=h).content
    except Exception as e:
        continue
    t = G.parse_enagas_trucks(c)
    if t == t:
        print(ym, f.split("/")[-1], "ok", t, flush=True); continue
    with pdfplumber.open(io.BytesIO(c)) as pdf:
        pages = [(pg.extract_text() or "") for pg in pdf.pages]
    hit = [i for i, p in enumerate(pages) if re.search(r"regasif|REGASIF|truck|TRUCK|cisterna|CISTERNA", re.sub(r"\s+", "", p), re.I) and re.search(r"BARCELONA|Barcelona", p)]
    rep[f.split("/")[-1]] = {"ym": ym, "pages": len(pages), "hit": hit, "text": [pages[i][:1800] for i in hit[:2]]}
    print(ym, f.split("/")[-1], "MISS", hit, flush=True)
json.dump(rep, open(os.path.join(OUT, "enagas_trucks_missing.json"), "w"), ensure_ascii=False, indent=1)
