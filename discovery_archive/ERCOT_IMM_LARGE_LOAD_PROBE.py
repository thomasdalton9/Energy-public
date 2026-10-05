"""One-off probe (5 Oct 2026): ERCOT Independent Market Monitor (Potomac Economics) State of the Market 2025 and quarterly reports: crypto / data-centre / large-load MW with context. Log only."""
import re, requests, pymupdf
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
P = "https://www.potomaceconomics.com/wp-content/uploads/2026/"
URLS = [P + "06/2025-State-of-the-Market-Report-for-ERCOT.pdf", P + "07/FY2025-26-Q3-IMM_Quarterly_Market_Report.pdf", P + "05/FY2025-26-Q2-IMM_Quarterly_Market_Report.pdf",
        P + "08/2026-07_Nodal_Monthly_Report.pdf"]
KEY = re.compile(r"crypto|data.?cent|large.?load|bitcoin|flexible load|hydrogen|oil and gas|industrial", re.I)
NUM = re.compile(r"\d[\d,.]*\s?(GW|MW|gigawatt|megawatt|TWh|GWh)", re.I)
for u in URLS:
    try:
        r = requests.get(u, headers=H, timeout=180)
        doc = pymupdf.open(stream=r.content, filetype="pdf")
    except Exception as e:
        print("FAIL", u, type(e).__name__); continue
    print(f"\n==== {u} pages={len(doc)}")
    for i, pg in enumerate(doc):
        t = re.sub(r"\s+", " ", pg.get_text())
        for s in re.split(r"(?<=[.;])\s", t):
            if KEY.search(s) and NUM.search(s) and re.search(r"ERCOT|Texas|queue|mine|load", s, re.I):
                print(f"  p{i+1}: {s.strip()[:520]}")
print("DONE")
