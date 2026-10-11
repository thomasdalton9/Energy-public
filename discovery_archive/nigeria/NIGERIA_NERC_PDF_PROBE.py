"""Nigeria: do NERC's PDFs (quarterly report, operational factsheet) hold generation by plant/fuel as extractable text tables?"""
import re, os, urllib3, requests
urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
URLS = {
 "Q2_2026": "https://nerc.gov.ng/wp-content/uploads/2026/09/2026_Q2_Report.pdf",
 "Fact_Apr2026": "https://nerc.gov.ng/wp-content/uploads/2026/05/Operational-Factsheet-April-2026.pdf.pdf",
 "Fact_Feb2026": "https://nerc.gov.ng/wp-content/uploads/2026/03/Operational-Performance-Factsheet-February-2026.pdf",
 "Harm_2020": "https://nerc.gov.ng/wp-content/uploads/2020/07/Harmonized Report for Friday, June 12 to Thursday, June 18, 2020.pdf",
}
import fitz
os.makedirs("/tmp/ng", exist_ok=True)
for k, u in URLS.items():
    try:
        r = requests.get(u, headers=H, timeout=(10, 90), verify=False)
        print(f"\n##### {k} {r.status_code} {len(r.content)}B {r.headers.get('content-type')}", flush=True)
        if r.status_code != 200 or not r.content.startswith(b"%PDF"):
            continue
        p = f"/tmp/ng/{k}.pdf"
        open(p, "wb").write(r.content)
        doc = fitz.open(p)
        print("pages", len(doc), flush=True)
        hits = []
        for i, pg in enumerate(doc):
            t = pg.get_text()
            if re.search(r"(?i)hydro|gas.fired|energy (generated|sent)|generation (by|per)|GENCO|Kainji|Egbin|Jebba|Shiroro", t):
                hits.append(i)
        print("pages mentioning plants/generation:", hits[:40], flush=True)
        for i in hits[:4]:
            print(f"--- page {i+1} text ---")
            print(re.sub(r"[ \t]+", " ", doc[i].get_text())[:1800], flush=True)
    except Exception as e:
        print(k, "ERR", type(e).__name__, str(e)[:150], flush=True)
