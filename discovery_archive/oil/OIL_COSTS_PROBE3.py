"""Oil cost curve probe 3 (Actions): EIA STEO regional oil series list + values, Dallas Fed breakeven.xlsx, COP / TPL / HPK / EOG decks."""
import os, re, io, time, json, requests
from urllib.parse import urljoin
OUT = "discovery_archive/results/oil"
os.makedirs(OUT + "/decks", exist_ok=True)
from curl_cffi import requests as cr
import pdfplumber, openpyxl
KEY = os.environ.get("EIA_API_KEY", "")
def get(u):
    try:
        return cr.get(u, impersonate="chrome", timeout=120)
    except Exception as e:
        print("ERR", u, str(e)[:120], flush=True)
# --- EIA STEO series with names
r = requests.get("https://api.eia.gov/v2/steo/facet/seriesId", params={"api_key": KEY}, timeout=120)
print("steo facet", r.status_code, flush=True)
if r.status_code == 200:
    fac = r.json()["response"]["facets"]
    open(f"{OUT}/eia_steo_series.txt", "w").write("\n".join(f"{f['id']}\t{f.get('name','')}" for f in fac))
    hits = [f for f in fac if re.search(r"permian|bakken|eagle ford|niobrara|anadarko|appalachia|haynesville|tight|lower 48|crude oil production|CORIPUS|COPR", (f["id"] + " " + f.get("name", "")), re.I)]
    for f in hits[:120]: print(f["id"], f.get("name"), flush=True)
# values for candidate ids found by name
ids = [f["id"] for f in fac if re.search(r"crude oil.*(permian|bakken|eagle|niobrara|anadarko|appalach|haynes|region)|permian.*crude|(permian|bakken|eagle ford|niobrara).*oil", f.get("name", ""), re.I)] if r.status_code == 200 else []
print("candidate ids", ids, flush=True)
out = []
for i in ids[:40]:
    d = requests.get("https://api.eia.gov/v2/steo/data", params={"api_key": KEY, "frequency": "monthly", "data[0]": "value", "facets[seriesId][]": i, "start": "2025-01", "sort[0][column]": "period", "sort[0][direction]": "asc", "length": 60}, timeout=60)
    if d.status_code == 200:
        for x in d.json()["response"]["data"]:
            out.append(f"{i}\t{x['period']}\t{x['value']}\t{x.get('seriesDescription','')}\t{x.get('unit','')}")
open(f"{OUT}/eia_steo_oil_regions.txt", "w").write("\n".join(out))
# --- Dallas Fed breakeven history workbook
for n, u in {"dallas_breakeven_xlsx": "https://www.dallasfed.org/-/media/documents/research/surveys/des/documents/breakeven.xlsx"}.items():
    x = get(u)
    if x is not None:
        print(n, x.status_code, len(x.content), flush=True)
        if x.status_code == 200 and x.content[:2] == b"PK":
            wb = openpyxl.load_workbook(io.BytesIO(x.content), data_only=True)
            lines = []
            for ws in wb:
                lines.append(f"## SHEET {ws.title}")
                for row in ws.iter_rows(values_only=True):
                    if any(v is not None for v in row): lines.append(" | ".join("" if v is None else str(v) for v in row))
            open(f"{OUT}/{n}.txt", "w").write("\n".join(lines)[:300000])
# --- decks
DECKS = {"COP_2q26_deck": "https://static.conocophillips.com/files/resources/2q26_earnings_release_deck.pdf",
         "COP_2q26_supp": "https://static.conocophillips.com/files/resources/2q26_supplemental-information-final.pdf",
         "TPL_2q26_deck": "https://d1io3yog0oux5.cloudfront.net/_44c0a1327e37db57b36e014028684610/texaspacific/db/706/6716/pdf/2Q+2026+TPL+Investor+Presentation+vF.pdf"}
for n, u in DECKS.items():
    x = get(u)
    if x is None or x.status_code != 200: print(n, "fail", None if x is None else x.status_code, flush=True); continue
    with pdfplumber.open(io.BytesIO(x.content)) as p:
        t = "\n".join((q.extract_text() or "") + f"\n[[page {i+1}]]" for i, q in enumerate(p.pages))
    open(f"{OUT}/decks/{n}.txt", "w").write(f"URL: {u}\nCHARS: {len(t)}\n" + t[:250000]); print(n, len(t), flush=True)
    try:
        import fitz
        d = fitz.open(stream=x.content, filetype="pdf")
        os.makedirs(f"{OUT}/decks/pages_{n}", exist_ok=True)
        for pi, pg in enumerate(d):
            if re.search(r"break-?even|inventory|per lateral|EUR", pg.get_text(), re.I) and pi < 60:
                pg.get_pixmap(dpi=70).save(f"{OUT}/decks/pages_{n}/p{pi+1}.png")
    except Exception as e: print("render", e, flush=True)
for n, u in {"HPK_pres": "https://ir.highpeakenergy.com/news-events/presentations", "EOG_pres": "https://investors.eogresources.com/events-and-presentations",
             "OXY_pres": "https://www.oxy.com/investors/events-presentations/", "DVN_pres": "https://www.devonenergy.com/investors/presentations",
             "FANG_pres": "https://ir.diamondbackenergy.com/", "PR_pres": "https://permianres.com/investors/"}.items():
    x = get(u)
    if x is None: continue
    print(n, x.status_code, len(x.content), flush=True)
    if x.status_code == 200:
        open(f"{OUT}/decks/{n}_links.txt", "w").write("\n".join(sorted(set(urljoin(u, h) for h in re.findall(r'href="([^"]+)"', x.text)))))
