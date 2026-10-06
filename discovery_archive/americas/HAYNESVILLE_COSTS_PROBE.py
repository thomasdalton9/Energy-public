"""Discovery: public sources of Haynesville breakeven economics. Downloads candidate documents (opened in Actions),
extracts text to discovery_archive/results/haynesville/*.txt and logs reachability + keyword hits. Manual only."""
import os, re, sys, io, json, time, requests
OUT = "discovery_archive/results/haynesville"
os.makedirs(OUT, exist_ok=True)
UA = {"User-Agent": "Mozilla/5.0 (research; thomas.dalton@cantab.net)", "Accept": "*/*"}
KW = re.compile(r"break-?\s?even|NPV-?10|Haynesville|Bossier|EUR|per lateral foot|well cost|inventory|locations", re.I)
STRONG = re.compile(r"break-?\s?even|NPV-?10|\bEUR\b|lateral foot|\$/Mcf|/MMBtu", re.I)
URLS = {
 "eia_today_67944": "https://www.eia.gov/todayinenergy/detail.php?id=67944",
 "eia_upstream_costs": "https://www.eia.gov/analysis/studies/drilling/pdf/upstream.pdf",
 "eia_ogsm_assump": "https://www.eia.gov/outlooks/aeo/assumptions/pdf/OGSM_Assumptions.pdf",
 "eia_drilling": "https://www.eia.gov/petroleum/drilling/",
 "beg_haynesville_decades": "https://www.beg.utexas.edu/files/content/beg/research/shale/Haynesville%20Shale%20to%20remain%20a%20major%20producer%20for%20decades,%20researchers%20f.pdf",
 "beg_able_competitor": "https://www.beg.utexas.edu/files/content/beg/ext-aff/16-02/Haynesville An Able Price Competitor With Northeast NatGas in 'Plenty of Areas'.pdf",
 "dallasfed_des_index": "https://www.dallasfed.org/research/surveys/des",
 "dallasfed_2501": "https://dallasfed.org/research/surveys/des/2025/2501",
 "kcfed_energy": "https://www.kansascityfed.org/surveys/energy-survey/",
 "crk_ir": "https://www.comstockresources.com/investors/presentations",
 "crk_ir2": "https://ir.comstockresources.com/",
 "exe_ir": "https://investors.expandenergy.com/",
 "exe_presentations": "https://investors.expandenergy.com/news-events/presentations",
 "ogj_economic_envelopes": "https://ogj.com/general-interest/companies/article/17227022/louisiana-haynesville-shale2-economic-operating-envelopes-characterized-for-haynesville-shale",
 "iea_gmr": "https://www.iea.org/reports/gas-market-report-q3-2026",
 "oies": "https://www.oxfordenergy.org/publications/",
 "baker": "https://www.bakerinstitute.org/research",
 "crk_8k_2024_a": "https://www.sec.gov/Archives/edgar/data/23194/000095017024014590/crk-ex99_1.htm",
}
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
for k, u in URLS.items():
    report(k, u); time.sleep(0.5)
# EDGAR crawl
CIKS = {"CRK": 23194, "EXE": 895126}
for tk, cik in CIKS.items():
    r = get(f"https://data.sec.gov/submissions/CIK{cik:010d}.json")
    if not hasattr(r, "status_code") or r.status_code != 200:
        print(f"[edgar {tk}] submissions failed {r if not hasattr(r,'status_code') else r.status_code}", flush=True); continue
    rec = r.json()["filings"]["recent"]
    rows = list(zip(rec["form"], rec["filingDate"], rec["accessionNumber"], rec["primaryDocument"]))
    print(f"[edgar {tk}] {len(rows)} recent filings, newest {rows[0][:2]}", flush=True)
    n10k = n10q = n8k = 0
    for form, d, acc, doc in rows:
        a = acc.replace("-", "")
        base = f"https://www.sec.gov/Archives/edgar/data/{cik}/{a}"
        if form == "10-K" and n10k < 1:
            n10k += 1; report(f"{tk}_10K_{d}", f"{base}/{doc}")
        elif form == "10-Q" and n10q < 1:
            n10q += 1; report(f"{tk}_10Q_{d}", f"{base}/{doc}")
        elif form == "8-K" and n8k < 8:
            n8k += 1
            idx = get(f"{base}/")
            if not hasattr(idx, "status_code") or idx.status_code != 200:
                print(f"[{tk} 8-K {d}] index failed", flush=True); continue
            links = sorted(set(re.findall(r'href="(/Archives/edgar/data/%d/%s/[^"]+\.(?:htm|pdf))"' % (cik, a), idx.text)))
            for l in links:
                if re.search(r"ex-?99|ex-?10|presentation|investor", l, re.I):
                    report(f"{tk}_8K_{d}_{l.rsplit('/',1)[1][:40]}", "https://www.sec.gov" + l)
                    time.sleep(0.3)
        time.sleep(0.3)
        if n10k and n10q and n8k >= 8: break
