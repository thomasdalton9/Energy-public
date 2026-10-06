"""Probe 4: render company-presentation pages with the inventory / breakeven / cost charts; Expand EDGAR earnings releases; press summaries."""
import os, re, io, time, requests
OUT = "discovery_archive/results/haynesville"
UA = {"User-Agent": "Mozilla/5.0 (research; thomas.dalton@cantab.net)"}
from curl_cffi import requests as cr
import fitz, pdfplumber
def pdf(u):
    x = cr.get(u, impersonate="chrome", timeout=90); print("PDF", u[-40:], x.status_code, len(x.content), flush=True); return x.content
os.makedirs(f"{OUT}/pres_pages", exist_ok=True)
for nm, u, pages in (("crk2610", "https://investors.comstockresources.com/static-files/5a596a22-02f6-4b49-a9cc-93ecfe0179c0", [13, 14, 15, 16, 21, 22, 30]),
                     ("crk2312", "https://investors.comstockresources.com/static-files/0f5f7e1b-b8bd-41f0-a34c-3846b8301eb8", [4, 5, 6, 7, 8, 9, 10, 11]),
                     ("exe3q25", "https://investors.expandenergy.com/static-files/b269b415-dfe3-4eca-b66c-250be02bcbae", [4, 12, 13])):
    d = fitz.open(stream=pdf(u), filetype="pdf")
    for p in pages:
        if p <= len(d): d[p-1].get_pixmap(dpi=100).save(f"{OUT}/pres_pages/{nm}_p{p}.png")
def txt(name, u, imp=True):
    try:
        x = cr.get(u, impersonate="chrome", timeout=60) if imp else requests.get(u, headers=UA, timeout=60)
    except Exception as e:
        print(f"[{name}] ERR {e}", flush=True); return
    print(f"[{name}] {x.status_code} {len(x.content)}", flush=True)
    if x.status_code != 200: return
    if x.content[:5] == b"%PDF-":
        with pdfplumber.open(io.BytesIO(x.content)) as p: t = "\n".join((q.extract_text() or "") + f"\n[[page {i+1}]]" for i, q in enumerate(p.pages))
    else:
        t = re.sub(r"[ \t\xa0]+", " ", re.sub(r"<[^>]+>", "\n", re.sub(r"<(script|style)[\s\S]*?</\1>", " ", x.text)))
    open(f"{OUT}/{name}.txt", "w").write(f"URL: {u}\n" + t[:1500000])
for nm, u in {"marketbeat_crk_q2_2026": "https://www.marketbeat.com/instant-alerts/comstock-resources-q2-earnings-call-highlights-2026-07-31/",
              "nasdaq_crk_q2": "https://www.nasdaq.com/articles/comstock-resources-q2-earnings-call-highlights",
              "exe_p17681": "https://www.expandenergy.com/?p=17681",
              "ogj_exe_q2_2026": "https://www.ogj.com/general-interest/companies/news/55394461/expand-energy-sticks-to-full-year-production-goal-touts-twin-eagle-marketing-purchases-prospects",
              "worldoil_exe_q4": "https://worldoil.com/news/2026/2/18/expand-energy-reports-strong-q4-targets-higher-gas-output-in-2026/",
              "exe_8k_20260630": "https://www.sec.gov/Archives/edgar/data/0000895126/000089512626000046/exe-ex_991x20260630x8kxpr.htm",
              "eia_tie_56361_b": "https://www.eia.gov/todayinenergy/detail.php?id=56361"}.items():
    txt(nm, u); time.sleep(0.5)
# Expand 8-K exhibits 2026 (underscore-named)
r = requests.get("https://data.sec.gov/submissions/CIK0000895126.json", headers=UA, timeout=60).json()["filings"]["recent"]
for form, d, acc in zip(r["form"], r["filingDate"], r["accessionNumber"]):
    if form != "8-K" or d < "2026-01-01": continue
    a = acc.replace("-", ""); base = f"https://www.sec.gov/Archives/edgar/data/895126/{a}"
    idx = requests.get(base + "/", headers=UA, timeout=60)
    links = sorted(set(re.findall(r'href="(/Archives/edgar/data/895126/%s/[^"]+\.(?:htm|pdf))"' % a, idx.text)))
    print(d, [l.rsplit("/", 1)[1] for l in links], flush=True)
    for l in links:
        n = l.rsplit("/", 1)[1]
        if re.search(r"ex[-_]?99|pres", n, re.I) and not os.path.exists(f"{OUT}/EXE_8K_{d}_{n[:40]}.txt"):
            txt(f"EXE_8K_{d}_{n[:40]}", "https://www.sec.gov" + l, imp=False); time.sleep(0.3)
