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

for tag in ["2024/Q2", "2023/Q2", "2022/Q2", "2021/Q3", "2020/Q2", "2019/Q2"]:
    d = doc(url_of(tag))
    txt = " ".join(p.get_text() for p in d)
    txt = re.sub(r"\s+", " ", txt)
    print(f"\n## {tag}: {len(d)} pages")
    for pat in [r"[^.]{0,120}(?:total|Total)[^.]{0,40}(?:generation|generated)[^.]{0,120}", r"[^.]{0,80}(?:hydro|Hydro)[^.]{0,30}(?:share|mix|contribut)[^.]{0,120}"]:
        for m in list(re.finditer(pat, txt))[:3]:
            print("   >", m.group(0)[:300])
