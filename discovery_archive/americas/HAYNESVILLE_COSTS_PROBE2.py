"""Discovery: public sources of Haynesville breakeven economics. Downloads candidate documents (opened in Actions),
extracts text to discovery_archive/results/haynesville/*.txt and logs reachability + keyword hits. Manual only."""
import os, re, sys, io, json, time, requests
OUT = "discovery_archive/results/haynesville"
os.makedirs(OUT, exist_ok=True)
UA = {"User-Agent": "Mozilla/5.0 (research; thomas.dalton@cantab.net)", "Accept": "*/*"}
KW = re.compile(r"break-?\s?even|NPV-?10|Haynesville|Bossier|EUR|per lateral foot|well cost|inventory|locations", re.I)
STRONG = re.compile(r"break-?\s?even|NPV-?10|\bEUR\b|lateral foot|\$/Mcf|/MMBtu", re.I)
def get(u, t=60):
    for a in range(2):
        try:
            return requests.get(u, headers=UA, timeout=t)
        except Exception as e:
            err = e; time.sleep(2)
    return err
def text_of(r, u):
    ct = r.headers.get("content-type", "").lower()
    if "pdf" in ct or u.lower().endswith(".pdf") or r.content[:5] == b"%PDF-":
        try:
            import pdfplumber
            with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                return "\n".join((p.extract_text() or "") + f"\n[[page {i+1}]]" for i, p in enumerate(pdf.pages))
        except Exception as e:
            return f"PDFERR {e}"
    from html import unescape
    t = re.sub(r"<(script|style)[\s\S]*?</\1>", " ", r.text)
    t = re.sub(r"<(br|/p|/tr|/div|/li|/h\d)[^>]*>", "\n", t)
    t = re.sub(r"<[^>]+>", " ", t)
    return re.sub(r"[ \t\xa0]+", " ", unescape(t))
def report(name, u):
    r = get(u)
    if not hasattr(r, "status_code"):
        print(f"[{name}] ERR {r} {u}", flush=True); return None
    n = len(r.content)
    if r.status_code != 200:
        print(f"[{name}] HTTP {r.status_code} {u}", flush=True); return None
    t = text_of(r, u)
    open(f"{OUT}/{name}.txt", "w").write(f"URL: {u}\nFETCHED: {time.strftime('%Y-%m-%d')}\n\n" + t[:1500000])
    hits = [l.strip() for l in t.splitlines() if STRONG.search(l) and re.search("haynes|bossier|breakeven|break-even|NPV", l, re.I)]
    print(f"[{name}] OK {n} bytes, {len(t)} chars, {len(hits)} strong lines {u}", flush=True)
    for l in hits[:6]: print("     ", l[:220], flush=True)
    return t

URLS = {
 "crk_pres_2026_10": "https://investors.comstockresources.com/static-files/5a596a22-02f6-4b49-a9cc-93ecfe0179c0",
 "crk_pres_2023_12": "https://investors.comstockresources.com/static-files/0f5f7e1b-b8bd-41f0-a34c-3846b8301eb8",
 "crk_q4_2025_release": "https://investors.comstockresources.com/static-files/c8faaf24-c415-4027-bc38-953660f4d9d5",
 "crk_q4_2024_release": "https://investors.comstockresources.com/static-files/1cbc0e8c-5c94-4d82-acc0-e39b4618be49",
 "crk_ir_home": "https://investors.comstockresources.com/",
 "crk_ir_presentations": "https://investors.comstockresources.com/events-and-presentations",
 "exe_3q25_pres": "https://investors.expandenergy.com/static-files/b269b415-dfe3-4eca-b66c-250be02bcbae",
 "exe_local_matters_4q25": "https://www.expandenergy.com/2026/03/10/local-matters-4q25-haynesville/",
 "exe_ir_home": "https://investors.expandenergy.com/",
 "exe_ir_pres": "https://investors.expandenergy.com/events-and-presentations",
 "beg_ogj_study": "https://ogj.com/general-interest/article/17236900/study-forecasts-gradual-haynesville-production-recovery-before-final-decline",
 "beg_haynesville_play": "https://www.beg.utexas.edu/files/content/beg/research/shale/Haynesville%20Shale%20Gas%20Play.pdf",
 "worldoil_2026_haynesville": "https://worldoil.com/magazine/2026/august/features/haynesville-shale-calculated-re-engagement-characterizes-operator-behavior/",
 "evaluate_energy": "https://info.evaluateenergy.com/haynesville-producers-eye-rising-gas-prices/",
 "dallasfed_2601": "https://www.dallasfed.org/research/surveys/des/2026/2601",
 "dallasfed_2501_full": "https://www.dallasfed.org/research/surveys/des/2025/2501",
 "dallasfed_2604": "https://www.dallasfed.org/research/surveys/des/2026/2603",
 "eia_aeo2025_ogsm": "https://www.eia.gov/outlooks/aeo/assumptions/pdf/OGSM_Assumptions.pdf",
 "eia_decline_curves": "https://www.eia.gov/analysis/drilling/curve_analysis/",
 "eia_dpr": "https://www.eia.gov/petroleum/drilling/pdf/dpr-full.pdf",
 "eia_steo_haynesville": "https://www.eia.gov/outlooks/steo/",
}
for k, u in URLS.items():
    report(k, u); time.sleep(0.5)
# EDGAR: every 8-K for CRK and EXE in the last ~14 months, list all exhibits, fetch any presentation / ex99 exhibit not yet seen
CIKS = {"CRK": 23194, "EXE": 895126}
for tk, cik in CIKS.items():
    r = get(f"https://data.sec.gov/submissions/CIK{cik:010d}.json")
    if not hasattr(r, "status_code") or r.status_code != 200:
        print(f"[edgar {tk}] failed", flush=True); continue
    rec = r.json()["filings"]["recent"]
    rows = [x for x in zip(rec["form"], rec["filingDate"], rec["accessionNumber"], rec["primaryDocument"]) if x[0] == "8-K" and x[1] >= "2025-01-01"]
    print(f"[edgar {tk}] {len(rows)} 8-Ks since 2025", flush=True)
    for form, d, acc, doc in rows:
        a = acc.replace("-", ""); base = f"https://www.sec.gov/Archives/edgar/data/{cik}/{a}"
        idx = get(f"{base}/")
        if not hasattr(idx, "status_code") or idx.status_code != 200: continue
        links = sorted(set(re.findall(r'href="(/Archives/edgar/data/%d/%s/[^"]+\.(?:htm|pdf))"' % (cik, a), idx.text)))
        print(f"[{tk} {d}] exhibits: {[l.rsplit('/',1)[1] for l in links]}", flush=True)
        for l in links:
            nm = l.rsplit("/", 1)[1]
            if re.search(r"ex-?99|ex-?10|pres", nm, re.I) and not os.path.exists(f"{OUT}/{tk}_8K_{d}_{nm[:40]}.txt"):
                report(f"{tk}_8K_{d}_{nm[:40]}", "https://www.sec.gov" + l); time.sleep(0.3)
