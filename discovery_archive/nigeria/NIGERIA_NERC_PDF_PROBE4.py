"""Nigeria: why 2025 NERC reports fail the plant-table parse and what older (2017-2024/Q2) reports hold on generation."""
import re, sys, urllib3, requests
import fitz
urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
B = "https://nerc.gov.ng/wp-content/uploads/"
import pandas as pd
REP = pd.read_excel("output/Data and Chart Outputs/nigeria_power_generation_quarterly.xlsx", sheet_name="Reports")
def url_of(q):
    return REP.loc[REP.quarter == q, "url"].iloc[0]
def doc(u):
    r = requests.get(u, headers=H, timeout=(10, 120), verify=False)
    return fitz.open(stream=r.content, filetype="pdf")

d = doc(url_of("2025/Q1"))
for i, pg in enumerate(d):
    t = pg.get_text()
    if "Average Hourly Generation" in t and re.search(r"(?m)^Total\s*$", t):
        lines = [l.strip() for l in t.split("\n") if l.strip()]
        print(f"## 2025Q1 p{i+1}: {len(lines)} lines; first 150:")
        print(" | ".join(lines[:150]))
        break

for tag, u in [("2024Q2", url_of("2024/Q2")), ("2023Q2", url_of("2023/Q2")), ("2021Q3", url_of("2021/Q3")), ("2019Q2", url_of("2019/Q2"))]:
    if not u:
        continue
    try:
        d = doc(u)
    except Exception as e:
        print(tag, "ERR", e); continue
    print(f"\n######## {tag}: {len(d)} pages; p1: {re.sub(chr(10), ' ', d[0].get_text())[:160]}")
    n = 0
    for i, pg in enumerate(d):
        t = pg.get_text()
        if re.search(r"(?i)(average hourly generation|energy generated|generation \(GWh\)|MWh/h|Kainji)", t) and n < 3:
            n += 1
            tt = re.sub(r"[ \t]+", " ", t)
            tt = re.sub(r"\n\s*\n+", "\n", tt)
            print(f"--- {tag} p{i+1} ---\n{tt[:1500]}")
