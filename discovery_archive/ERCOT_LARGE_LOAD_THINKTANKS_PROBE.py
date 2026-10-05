"""One-off probe (5 Oct 2026): think-tank / trade / regulator pages for ERCOT large-load GW by type (Grid Strategies, Brattle, E3, Texas Blockchain Council, PUCT 58317, LBNL, EIA). Log only."""
import re, requests
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
NUM = re.compile(r"\b\d[\d,.]*\s?(GW|MW|gigawatts?|megawatts?)\b", re.I)
TOP = re.compile(r"ERCOT|Texas", re.I)
KEY = re.compile(r"data.?cent|crypto|bitcoin|large.?load|industrial|hydrogen|oil and gas", re.I)
URLS = ["https://gridstrategiesllc.com/reports/", "https://gridstrategiesllc.com/", "https://www.brattle.com/insights-events/publications/",
        "https://www.ethree.com/publications/", "https://www.texasblockchaincouncil.org/", "https://www.texasblockchaincouncil.org/resources",
        "https://interchange.puc.texas.gov/search/filings/?ControlNumber=58317&ItemMatch=Equal&UtilityType=A&ItemNumber=1",
        "https://www.puc.texas.gov/industry/electric/rulemaking/", "https://www.eia.gov/todayinenergy/detail.php?id=65084",
        "https://www.eia.gov/todayinenergy/", "https://www.eia.gov/electricity/data/eia861/", "https://www.comptroller.texas.gov/economy/",
        "https://www.potomaceconomics.com/markets/ercot/", "https://www.potomaceconomics.com/wp-content/uploads/2026/06/2025-State-of-the-Market-Report-for-ERCOT.pdf",
        "https://www.ercot.com/files/docs/2026/05/13/ERCOT-Large-Load-Interconnection-Process-Update.pdf"]
for u in URLS:
    try:
        r = requests.get(u, headers=H, timeout=60)
    except Exception as e:
        print("FAIL", u, type(e).__name__); continue
    print(f"\n== {u} {r.status_code} {len(r.content)} bytes {r.headers.get('content-type','')[:40]}")
    if r.status_code != 200: continue
    if "pdf" in r.headers.get("content-type", ""):
        import fitz
        doc = fitz.open(stream=r.content, filetype="pdf"); txt = " ".join(p.get_text() for p in doc)
    else:
        txt = re.sub(r"<[^>]+>", " ", r.text)
        links = [l for l in re.findall(r'href=["\']([^"\']+)["\']', r.text) if re.search(r"ercot|texas|large.?load|data.?cent|crypto|\.pdf", l, re.I)]
        for l in sorted(set(links))[:25]: print("   link:", l[:200])
    txt = re.sub(r"\s+", " ", txt)
    n = 0
    for s in re.split(r"(?<=[.;])\s", txt):
        if NUM.search(s) and TOP.search(s) and KEY.search(s):
            print("   *", s.strip()[:380]); n += 1
            if n >= 14: break
print("DONE")
