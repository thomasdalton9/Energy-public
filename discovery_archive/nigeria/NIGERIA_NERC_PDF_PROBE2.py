"""Nigeria: inspect NERC quarterly report tables (plant generation, generation mix) and list all quarterly reports."""
import re, os, urllib3, requests
import fitz
urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
os.makedirs("/tmp/ng", exist_ok=True)

print("## all NERC report links", flush=True)
allf = {}
for pg in range(1, 9):
    u = "https://nerc.gov.ng/resource-category/nerc-reports/" + (f"page/{pg}/" if pg > 1 else "")
    try:
        r = requests.get(u, headers=H, timeout=(8, 40), verify=False)
    except Exception as e:
        print(u, "ERR", e); continue
    if r.status_code != 200:
        print(u, r.status_code); continue
    for f in sorted(set(re.findall(r'href=["\']([^"\']+\.pdf)', r.text, re.I))):
        allf[f] = pg
for f, pg in allf.items():
    print(pg, f.split("/uploads/")[-1])

r = requests.get("https://nerc.gov.ng/wp-content/uploads/2026/09/2026_Q2_Report.pdf", headers=H, timeout=(10, 90), verify=False)
open("/tmp/ng/q2.pdf", "wb").write(r.content)
doc = fitz.open("/tmp/ng/q2.pdf")
print("pages", len(doc), flush=True)
for i, pg in enumerate(doc):
    t = pg.get_text()
    heads = re.findall(r"(?im)^\s*((?:Table|Figure)\s+[A-Z0-9.]+.*)$", t)
    print(f"p{i+1}: {len(t)} chars; heads: {[h[:70] for h in heads[:4]]}", flush=True)
shown = 0
for i, pg in enumerate(doc):
    t = pg.get_text()
    if re.search(r"(?i)generation mix|energy generated|quarterly generation|Kainji", t) and shown < 5:
        shown += 1
        print(f"\n=== p{i+1} text ===\n" + re.sub(r"[ \t]+", " ", t)[:2500], flush=True)
        try:
            for tb in pg.find_tables().tables[:2]:
                print("TABLE rows:", tb.row_count, "cols:", tb.col_count)
                for row in tb.extract()[:8]:
                    print("  ", row)
        except Exception as e:
            print("find_tables err", e)
