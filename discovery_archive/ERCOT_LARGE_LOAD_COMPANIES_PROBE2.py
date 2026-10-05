"""One-off probe (5 Oct 2026): Texas utilities' large-load queue / energised MW by category in 10-K, 10-Q and earnings-release exhibits (EX-99). Log only."""
import re, time, requests
H = {"User-Agent": "Energy research thomas.dalton@cantab.net"}
NUM = re.compile(r"\b\d[\d,.]*\s?(GW|MW|megawatts?|gigawatts?)\b", re.I)
TOP = re.compile(r"data.?cent|crypto|large.?load|queue|interconnection request|industrial|hydrogen|oil and gas|oil & gas|LNG|load growth|signed|energi[sz]ed|in service|letters? of agreement|ESA|requests", re.I)
COS = {"Oncor": 1193311, "CenterPoint": 1130310, "AEP": 4904, "Sempra": 1032208, "TXNM": 1108426, "Vistra": 1692819, "NRG": 1013871, "Entergy": 65984, "Xcel": 72903}
def j(u):
    for k in range(3):
        try:
            r = requests.get(u, headers=H, timeout=60)
            return r if r.status_code == 200 else None
        except Exception: time.sleep(2)
def sents(t, n=30):
    t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t)).replace("&#8217;", "'").replace("&#8220;", '"').replace("&#8221;", '"').replace("&#160;", " ")
    out, seen = [], set()
    for s in re.split(r"(?<=[.;])\s", t):
        if NUM.search(s) and TOP.search(s) and re.search(r"Texas|ERCOT|Oncor|Houston|AEP Texas|TNMP|data cent|large load", s, re.I):
            if s[:70] in seen: continue
            seen.add(s[:70]); out.append(s.strip()[:520])
    return out[:n]
for name, cik in COS.items():
    r = j(f"https://data.sec.gov/submissions/CIK{cik:010d}.json")
    if r is None: print("##", name, "none"); continue
    d = r.json()["filings"]["recent"]
    rows = [(d["form"][i], d["filingDate"][i], d["accessionNumber"][i], d["primaryDocument"][i], d["items"][i]) for i in range(len(d["form"]))]
    pick = [x for x in rows if x[0] in ("10-K", "10-Q")][:3] + [x for x in rows if x[0] == "8-K" and "2.02" in x[4]][:3]
    print(f"\n######## {name}: {[(x[0], x[1]) for x in pick]}")
    for form, dt, acc, doc, it in pick:
        base = f"https://www.sec.gov/Archives/edgar/data/{cik}/{acc.replace('-', '')}/"
        docs = [base + doc]
        if form == "8-K":
            ix = j(base + "index.json")
            if ix is not None:
                docs = [base + f["name"] for f in ix.json()["directory"]["item"] if re.search(r"ex-?99|ex99|d\d+dex99|ex_99", f["name"], re.I) and f["name"].endswith((".htm", ".html"))][:3] or docs
        for u in docs:
            rr = j(u)
            if rr is None: continue
            hs = sents(rr.text, 30 if form != "8-K" else 25)
            print(f"-- {form} {dt} {u}: {len(hs)} hits")
            for s in hs: print("   *", s)
            time.sleep(0.3)
print("DONE")
