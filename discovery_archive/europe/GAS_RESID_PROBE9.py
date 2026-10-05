"""Probe 9 (gas residual, Spain): run the Enagas truck parser over every bulletin and print month, truck figures and the source file."""
import os, re, sys, json
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "europe"))
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
        except Exception:
            pass
rows = []
for f, ym in sorted(files.items(), key=lambda kv: kv[1]):
    try:
        c = G.get("https://www.enagas.es" + f, headers=h).content
        res = G.parse_enagas_bulletin(c)
        row = [list(ym), f.split("/")[-1], str(res[0])[:7] if res else None, G.parse_enagas_trucks(c), G.parse_enagas_trucks(c, "prev")]
    except Exception as e:
        row = [list(ym), f.split("/")[-1], "ERR " + type(e).__name__ + str(e)[:80]]
    rows.append(row); print(row, flush=True)
json.dump(rows, open(os.path.join(OUT, "enagas_trucks_all.json"), "w"), ensure_ascii=False, indent=0)
