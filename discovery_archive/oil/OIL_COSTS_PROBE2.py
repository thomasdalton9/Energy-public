"""Oil cost curve probe 2 (Actions): Dallas Fed breakeven page + chart workbooks; operator IR presentation pages -> latest decks (full text)."""
import os, re, io, time, requests
from urllib.parse import urljoin
OUT = "discovery_archive/results/oil"
os.makedirs(OUT + "/decks", exist_ok=True)
from curl_cffi import requests as cr
import pdfplumber, openpyxl
def get(u):
    try:
        return cr.get(u, impersonate="chrome", timeout=90)
    except Exception as e:
        print("ERR", u, str(e)[:120], flush=True)
def clean(x):
    return re.sub(r"[ \t\xa0]+", " ", re.sub(r"<[^>]+>", "\n", re.sub(r"<(script|style)[\s\S]*?</\1>", " ", x)))
# ---- Dallas Fed
for n, u in {"dallas_breakeven_page": "https://www.dallasfed.org/research/surveys/des/data/breakeven",
             "dallas_breakeven_page2": "https://www.dallasfed.org/research/surveys/des/data/breakeven.aspx"}.items():
    x = get(u)
    if x is not None:
        print(n, x.status_code, len(x.content), flush=True)
        if x.status_code == 200:
            open(f"{OUT}/{n}.txt", "w").write(clean(x.text)[:200000])
            open(f"{OUT}/{n}_links.txt", "w").write("\n".join(sorted(set(re.findall(r'href="([^"]+)"', x.text)))))
for n, u in {"des26q1_charts": "https://www.dallasfed.org/-/media/Documents/research/surveys/DES/2026/2601/des26q1_charts.xlsx",
             "des26q3_charts1": "https://www.dallasfed.org/-/media/Documents/research/surveys/DES/2026/2603/des26q3_charts1.xlsx",
             "des26q2_charts": "https://www.dallasfed.org/-/media/Documents/research/surveys/DES/2026/2602/des26q2_charts.xlsx",
             "des25q1_charts": "https://www.dallasfed.org/-/media/Documents/research/surveys/DES/2025/2501/des25q1_charts.xlsx"}.items():
    x = get(u)
    if x is None: continue
    print(n, x.status_code, len(x.content), flush=True)
    if x.status_code == 200 and x.content[:2] == b"PK":
        wb = openpyxl.load_workbook(io.BytesIO(x.content), data_only=True)
        lines = []
        for ws in wb:
            lines.append(f"## SHEET {ws.title}")
            for row in ws.iter_rows(values_only=True):
                if any(v is not None for v in row):
                    lines.append(" | ".join("" if v is None else str(v) for v in row))
        open(f"{OUT}/{n}.txt", "w").write("\n".join(lines)[:300000])
# ---- IR presentation pages
IR = {"FANG": ["https://ir.diamondbackenergy.com/events-and-presentations", "https://ir.diamondbackenergy.com/presentations", "https://www.diamondbackenergy.com/investors/events-and-presentations"],
      "PR": ["https://permianres.com/investors/events-presentations/", "https://permianres.com/investors/"],
      "DVN": ["https://www.devonenergy.com/investors/presentations", "https://www.devonenergy.com/investors/events-and-presentations"],
      "OXY": ["https://www.oxy.com/investors/events-presentations/", "https://www.oxy.com/investors/"],
      "COP": ["https://www.conocophillips.com/investor-relations/presentations/", "https://www.conocophillips.com/investor-relations/"],
      "EOG": ["https://investors.eogresources.com/presentations", "https://investors.eogresources.com/"],
      "CTRA": ["https://ir.coterra.com/events-and-presentations/", "https://www.coterra.com/investors/"],
      "MTDR": ["https://ir.matadorresources.com/presentations/", "https://investors.matadorresources.com/events-and-presentations"],
      "CHRD": ["https://ir.chordenergy.com/events-and-presentations", "https://ir.chordenergy.com/"],
      "CIVI": ["https://investors.civitasresources.com/events-and-presentations", "https://investors.civitasresources.com/"],
      "OVV": ["https://www.ovintiv.com/investors/presentations/", "https://www.ovintiv.com/investors/"],
      "APA": ["https://investor.apacorp.com/events-and-presentations", "https://investor.apacorp.com/"],
      "SM": ["https://investors.sm-energy.com/events-and-presentations", "https://investors.sm-energy.com/"],
      "CRGY": ["https://ir.crescentenergyco.com/events-and-presentations", "https://ir.crescentenergyco.com/"],
      "HPK": ["https://ir.highpeakenergy.com/events-and-presentations", "https://ir.highpeakenergy.com/"],
      "MGY": ["https://ir.magnoliaoilgas.com/events-and-presentations", "https://ir.magnoliaoilgas.com/"],
      "NOG": ["https://investors.northernoil.com/events-and-presentations", "https://investors.northernoil.com/"],
      "CVX": ["https://www.chevron.com/investors/presentations"],
      "XOM": ["https://corporate.exxonmobil.com/investors/investor-presentations"],
      "TPL": ["https://texaspacific.com/investors/"]}
PAT = re.compile(r'href="([^"]+)"', re.I)
for tk, urls in IR.items():
    cand = []
    for u in urls:
        x = get(u)
        if x is None: continue
        print(tk, u, x.status_code, len(x.content), flush=True)
        if x.status_code != 200: continue
        open(f"{OUT}/decks/{tk}_irpage_{abs(hash(u))%10000}_links.txt", "w").write("\n".join(sorted(set(PAT.findall(x.text)))))
        for h in PAT.findall(x.text):
            hl = h.lower()
            if (".pdf" in hl or "static-files" in hl or "q4cdn" in hl or "doc_presentations" in hl or "/files/" in hl) and not re.search(r"proxy|esg|sustain|10-k|10k|annual-report|code-of|charter|bylaws|governance|policy|climate", hl):
                cand.append(urljoin(u, h.replace("&amp;", "&")))
        time.sleep(0.5)
    seen, n = set(), 0
    for c in cand:
        if c in seen: continue
        seen.add(c)
        if n >= 3: break
        x = get(c)
        if x is None or x.status_code != 200 or x.content[:5] != b"%PDF-":
            print("  skip", c[:120], None if x is None else (x.status_code, x.content[:5]), flush=True); continue
        try:
            with pdfplumber.open(io.BytesIO(x.content)) as p:
                t = "\n".join((q.extract_text() or "") + f"\n[[page {i+1}]]" for i, q in enumerate(p.pages))
        except Exception as e:
            print("  pdf err", e, flush=True); continue
        n += 1
        open(f"{OUT}/decks/{tk}_deck{n}.txt", "w").write(f"URL: {c}\nCHARS: {len(t)}\n" + t[:220000])
        print("  saved", tk, n, c[:120], len(t), flush=True)
        time.sleep(0.5)
