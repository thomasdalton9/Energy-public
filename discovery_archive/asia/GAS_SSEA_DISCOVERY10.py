"""
South & Southeast Asia gas, round 10: why the column-position MBS parser rejects rows in the 2022 / 2023 issues -
print the table 3.2 header words and two data rows with their x-positions (MBS_Apr_2023, MBS_Apr_2022, MBS_Sep_2024).
"""
import io
import re

import pdfplumber
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
for f in ("MBS_Apr_2023.pdf", "MBS_Apr_2022.pdf", "MBS_Sep_2024.pdf"):
    c = requests.get("https://www.pbs.gov.pk/wp-content/uploads/2020/07/" + f, headers=H, timeout=(20, 240),
                     verify=False).content
    with pdfplumber.open(io.BytesIO(c)) as pdf:
        for p in pdf.pages[:60]:
            tx = p.extract_text() or ""
            if not re.search(r"3\.2\s+Production of Natural Gas", tx):
                continue
            print(f"\n##### {f} page {p.page_number}", flush=True)
            ws = sorted(p.extract_words(), key=lambda w: (round(w["top"]), w["x0"]))
            shown = 0
            for w in ws:
                if w["top"] < 200 or re.match(r"^(Bhalsyedan|Dakhni|Company|Field)$", w["text"]) or shown:
                    pass
            tops = sorted({round(w["top"]) for w in ws})
            for t in tops[:14]:
                row = [w for w in ws if abs(round(w["top"]) - t) <= 1]
                print(f"top {t}: " + " | ".join(f"{w['text']}@{w['x0']:.0f}-{w['x1']:.0f}" for w in row)[:1500], flush=True)
            break
