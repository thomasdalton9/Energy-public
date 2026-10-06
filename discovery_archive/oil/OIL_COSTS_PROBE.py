"""Oil cost curve probe 1 (Actions): SEC filings of US shale oil operators (10-K, 10-Q, 2026 8-K exhibits) -> keyword windows; Fed surveys; EIA cost pages.
Only keyword windows are saved (breakeven, per lateral foot, EUR, inventory, maintenance...), to discovery_archive/results/oil/."""
import os, re, io, time, json, requests
OUT = "discovery_archive/results/oil"
os.makedirs(OUT, exist_ok=True)
UA = {"User-Agent": "Mozilla/5.0 (research; thomas.dalton@cantab.net)"}
from curl_cffi import requests as cr
import pdfplumber
KEY = re.compile(r"break-?\s?even|per lateral foot|/lateral|lateral foot|\bEUR\b|estimated ultimate|inventory|locations|maintenance capital|maintenance capex|free cash flow (at|break)|IP-?30|IP-?90|first[- ]year decline|payout|\bIRR\b|PV-?10|WTI (price|of|at|oil)|oil cut|oil mix|% oil|well cost|cost per well|D&C|drilling and completion|net (inventory )?(drilling )?locations", re.I)
CIK = {"DVN": 1090012, "FANG": 1539838, "PR": 1658566, "OXY": 797468, "COP": 1163165, "EOG": 821189, "CTRA": 858470, "MTDR": 1520006,
       "CHRD": 1486159, "CIVI": 1509589, "OVV": 1792580, "CVX": 93410, "XOM": 34088, "APA": 1841666, "SM": 893538, "CRGY": 1866175,
       "HPK": 1792849, "VTLE": 1528129, "CLR": 732834, "HES": 4447, "MGY": 1698990, "NOG": 1104485, "REPX": 1384195, "RIG_PERM": 1001614, "WDS_TPL": 1811074}
def clean(x):
    return re.sub(r"[ \t\xa0]+", " ", re.sub(r"<[^>]+>", "\n", re.sub(r"<(script|style)[\s\S]*?</\1>", " ", x)))
def windows(t, w=350, cap=180000):
    out, last = [], -1
    for m in KEY.finditer(t):
        if m.start() < last: continue
        a, b = max(0, m.start() - w), min(len(t), m.end() + w)
        out.append(re.sub(r"\s+", " ", t[a:b])); last = b
    s = "\n---\n".join(out)
    return s[:cap]
def save(name, u, t, full=False):
    body = t[:300000] if full else windows(t)
    open(f"{OUT}/{name}.txt", "w").write(f"URL: {u}\nCHARS: {len(t)}\n" + body)
def get(u, imp=False):
    for k in range(2):
        try:
            x = cr.get(u, impersonate="chrome", timeout=90) if imp else requests.get(u, headers=UA, timeout=90)
            return x
        except Exception as e:
            print("ERR", u, e, flush=True); time.sleep(2)
def text_of(x):
    if x.content[:5] == b"%PDF-":
        with pdfplumber.open(io.BytesIO(x.content)) as p:
            return "\n".join((q.extract_text() or "") + f"\n[[page {i+1}]]" for i, q in enumerate(p.pages))
    return clean(x.text)
def doc(name, u, imp=False, full=False):
    x = get(u, imp)
    if x is None: print(f"[{name}] FAILED", flush=True); return
    print(f"[{name}] {x.status_code} {len(x.content)}", flush=True)
    if x.status_code == 200:
        save(name, u, text_of(x), full)
    time.sleep(0.4)
# ---------------- SEC
for tk, cik in CIK.items():
    try:
        r = requests.get(f"https://data.sec.gov/submissions/CIK{cik:010d}.json", headers=UA, timeout=60)
        if r.status_code != 200: print(tk, "submissions", r.status_code, flush=True); continue
        rec = r.json()["filings"]["recent"]
    except Exception as e:
        print(tk, "ERR", e, flush=True); continue
    got = {"10-K": 0, "10-Q": 0}
    n8 = 0
    for form, d, acc, prim in zip(rec["form"], rec["filingDate"], rec["accessionNumber"], rec["primaryDocument"]):
        a = acc.replace("-", ""); base = f"https://www.sec.gov/Archives/edgar/data/{cik}/{a}"
        if form in got and got[form] < 1 and d >= "2025-12-01":
            got[form] += 1; doc(f"{tk}_{form}_{d}", f"{base}/{prim}")
        elif form == "8-K" and d >= "2026-01-01" and n8 < 7:
            idx = get(base + "/")
            if idx is None or idx.status_code != 200: continue
            links = sorted(set(re.findall(r'href="(/Archives/edgar/data/%d/%s/[^"]+\.(?:htm|pdf))"' % (cik, a), idx.text)))
            for l in links:
                n = l.rsplit("/", 1)[1]
                if re.search(r"ex[-_]?99|pres|deck|investor", n, re.I):
                    n8 += 1; doc(f"{tk}_8K_{d}_{n[:40]}", "https://www.sec.gov" + l)
# ---------------- Fed surveys, EIA, others
PAGES = {"dallasfed_2601": "https://www.dallasfed.org/research/surveys/des/2026/2601",
         "dallasfed_2602": "https://www.dallasfed.org/research/surveys/des/2026/2602",
         "dallasfed_2603": "https://www.dallasfed.org/research/surveys/des/2026/2603",
         "dallasfed_2504": "https://www.dallasfed.org/research/surveys/des/2025/2504",
         "dallasfed_2503": "https://www.dallasfed.org/research/surveys/des/2025/2503",
         "dallasfed_2501": "https://www.dallasfed.org/research/surveys/des/2025/2501",
         "kcfed_energy": "https://www.kansascityfed.org/surveys/energy-survey/",
         "eia_upstream_costs": "https://www.eia.gov/analysis/studies/drilling/pdf/upstream.pdf",
         "eia_upstream_costs_page": "https://www.eia.gov/analysis/studies/drilling/",
         "eia_ogsm": "https://www.eia.gov/outlooks/aeo/assumptions/pdf/OGSM_Assumptions.pdf",
         "eia_ogsm_page": "https://www.eia.gov/outlooks/aeo/assumptions/pub/oilgas_supply.php",
         "eia_dpr": "https://www.eia.gov/petroleum/drilling/",
         "eia_steo": "https://www.eia.gov/outlooks/steo/",
         "iea_oil2026": "https://www.iea.org/reports/oil-2026",
         "opec_woo": "https://www.opec.org/world-oil-outlook-2025.html",
         "enverus_blog": "https://www.enverus.com/blog/",
         "rystad_press": "https://www.rystadenergy.com/news"}
for n, u in PAGES.items():
    doc(n, u, imp=True)
# Dallas Fed: dump links to tables / charts on the Q pages
for q in ("2601", "2602", "2603"):
    x = get(f"https://www.dallasfed.org/research/surveys/des/2026/{q}", True)
    if x is not None and x.status_code == 200:
        open(f"{OUT}/dallasfed_{q}_links.txt", "w").write("\n".join(sorted(set(re.findall(r'href="([^"]+)"', x.text)))))
        open(f"{OUT}/dallasfed_{q}_fulltext.txt", "w").write(clean(x.text)[:400000])
