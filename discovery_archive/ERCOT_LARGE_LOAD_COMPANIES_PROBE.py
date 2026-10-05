"""One-off probe (5 Oct 2026): SEC filings (latest 10-K/10-Q/8-K) of Texas utilities and ERCOT crypto/data-centre companies: sentences with MW/GW
of large-load queue / operating capacity. Plus PUCT 58317 / EIA / Brattle quick fetches. Output is the log only."""
import re, sys, json, time
import requests
H = {"User-Agent": "Energy research thomas.dalton@cantab.net", "Accept-Encoding": "gzip, deflate"}
NUM = re.compile(r"\b\d[\d,.]*\s?(GW|MW|megawatts?|gigawatts?)\b", re.I)
TOPIC = re.compile(r"data.?cent|crypto|bitcoin|large.?load|hash|hydrogen|oil and gas|industrial|energi[sz]ed|in service|operating|ERCOT", re.I)
COS = {  # name: cik
 "Oncor Electric Delivery": 1193311, "CenterPoint Energy": 1130310, "AEP": 4904, "Sempra": 1032208, "PNM Resources/TXNM": 1108426,
 "Riot Platforms": 1167419, "Cipher Mining": 1819989, "Core Scientific": 1839341, "TeraWulf": 1083301, "Galaxy Digital": 1859392, "MARA": 1507605,
 "CleanSpark": 827876, "Bitdeer": 1899123, "Hut 8": 1964789, "IREN": 1878848, "Applied Digital": 1144879, "Bitfarms": 1812477, "Crusoe": 0,
 "NRG": 1013871, "Vistra": 1692819, "Talen": 1622536, "Constellation": 1868275, "Cipher": 1819989, "Soluna": 1074902, "Bit Digital": 1710350,
 "Hive": 0, "Galaxy": 1859392, "Lancium": 0, "Texas Pacific Land": 1811074, "Energy Transfer": 1276187, "Kinder Morgan": 1506307,
}
def j(u):
    for k in range(3):
        try:
            r = requests.get(u, headers=H, timeout=60)
            if r.status_code == 200: return r
            print("  status", r.status_code, u); return None
        except Exception as e:
            time.sleep(2)
    return None
def sentences(t, maxn=25, need=NUM):
    t = re.sub(r"\s+", " ", t)
    out = []
    for s in re.split(r"(?<=[.;])\s", t):
        if need.search(s) and re.search(r"ERCOT|Texas|Oncor|CenterPoint|Houston|Rockdale|Corsicana|Helios|Abilene|Sweetwater|Barber Lake|Pecos|Ector|Cedarvale|Odessa|Denton|Alvarado|Harwood|Bruceville|Muskogee|Childress", s) and TOPIC.search(s):
            out.append(s.strip()[:420])
    seen = set(); res = []
    for s in out:
        if s[:80] in seen: continue
        seen.add(s[:80]); res.append(s)
    return res[:maxn]
for name, cik in COS.items():
    if not cik: continue
    r = j(f"https://data.sec.gov/submissions/CIK{cik:010d}.json")
    if r is None: print("##", name, "no submissions"); continue
    d = r.json()["filings"]["recent"]
    rows = [(d["form"][i], d["filingDate"][i], d["accessionNumber"][i], d["primaryDocument"][i]) for i in range(len(d["form"]))]
    pick = [x for x in rows if x[0] in ("10-K", "10-Q")][:2] + [x for x in rows if x[0] == "8-K"][:3]
    print(f"\n######## {name} CIK {cik}: {[(x[0], x[1]) for x in pick]}")
    for form, dt, acc, doc in pick:
        u = f"https://www.sec.gov/Archives/edgar/data/{cik}/{acc.replace('-', '')}/{doc}"
        rr = j(u)
        if rr is None: continue
        txt = re.sub(r"<[^>]+>", " ", rr.text)
        hits = sentences(txt, 18 if form != "8-K" else 8)
        print(f"-- {form} {dt} {u} ({len(txt)} chars): {len(hits)} hits")
        for s in hits: print("   *", s)
        time.sleep(0.3)
# EDGAR full text search (exhibits / press releases)
for q in ['"large load" ERCOT "data center" queue gigawatts', 'ERCOT "crypto" "large load" queue MW', 'Oncor "large load" interconnection queue GW data center']:
    for fr in ("2026-01-01",):
        r = j("https://efts.sec.gov/LATEST/search-index?q=" + requests.utils.quote(q) + f"&dateRange=custom&startdt={fr}&enddt=2026-10-05")
        print("\nFTS", q, None if r is None else r.text[:1500])
print("DONE")
