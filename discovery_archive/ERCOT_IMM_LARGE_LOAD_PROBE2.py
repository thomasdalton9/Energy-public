"""One-off probe (5 Oct 2026): full text of the IMM 2025 State of the Market pages on crypto / large loads (context for the 4.6 GW figure). Log only."""
import re, requests, pymupdf
u = "https://www.potomaceconomics.com/wp-content/uploads/2026/06/2025-State-of-the-Market-Report-for-ERCOT.pdf"
doc = pymupdf.open(stream=requests.get(u, headers={"User-Agent": "Mozilla/5.0"}, timeout=180).content, filetype="pdf")
for i in (36, 49, 50, 51, 52, 53):
    print(f"\n---- p{i}\n", re.sub(r"\s+", " ", doc[i - 1].get_text())[:3800])
for i, pg in enumerate(doc):
    t = pg.get_text()
    for m in re.finditer(r"[^.]{0,260}(cryptocurrenc|data cent)[^.]{0,260}\.", t, re.I):
        sn = re.sub(r"\s+", " ", m.group(0))
        if re.search(r"\d", sn) and i + 1 not in (36, 49, 50, 51, 52, 53): print(f"  p{i+1}: {sn[:520]}")
print("DONE")
