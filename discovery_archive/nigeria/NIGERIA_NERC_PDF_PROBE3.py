"""Nigeria: NERC quarterly report generation section (2.1.3 quarterly generation, 2.1.5 generation mix) + report list."""
import re, os, urllib3, requests
import fitz
urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
os.makedirs("/tmp/ng", exist_ok=True)
allf = {}
for pg in range(1, 9):
    u = "https://nerc.gov.ng/resource-category/nerc-reports/" + (f"page/{pg}/" if pg > 1 else "")
    try:
        r = requests.get(u, headers=H, timeout=(8, 40), verify=False)
        if r.status_code != 200:
            continue
        for f in set(re.findall(r'href=["\']([^"\']+\.pdf)', r.text, re.I)):
            allf[f.split("/uploads/")[-1]] = pg
    except Exception as e:
        print("ERR", u, e)
print("REPORTS:", sorted(allf), flush=True)
r = requests.get("https://nerc.gov.ng/wp-content/uploads/2026/09/2026_Q2_Report.pdf", headers=H, timeout=(10, 90), verify=False)
open("/tmp/ng/q2.pdf", "wb").write(r.content)
doc = fitz.open("/tmp/ng/q2.pdf")
for i in range(24, 31):
    t = re.sub(r"[ \t]+", " ", doc[i].get_text())
    t = re.sub(r"\n\s*\n+", "\n", t)
    print(f"\n=== p{i+1} ===\n{t[:2200]}", flush=True)
    try:
        for tb in doc[i].find_tables().tables[:3]:
            if tb.row_count > 3:
                print("TABLE", tb.row_count, "x", tb.col_count)
                for row in tb.extract()[:34]:
                    print("  ", row)
    except Exception as e:
        print("tables err", e)
