"""
Probe 6 (SE Europe capacity). Output: results/se_probe6.txt
 - HROTE monthly incentive-scheme report PDFs: which months exist at the 2023-style URL pattern (2019-2026), what the 2024-26 year pages hold
 - ERE Albania periodic market-data PDF: is there an installed-capacity table
"""
import os, re, sys, io, html
import requests
from urllib.parse import quote
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results"); os.makedirs(OUT, exist_ok=True)
LOG = open(os.path.join(OUT, "se_probe6.txt"), "w")
def p(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); LOG.write(s + "\n"); LOG.flush()
import fitz
MON = ["siječanj", "veljača", "ožujak", "travanj", "svibanj", "lipanj", "srpanj", "kolovoz", "rujan", "listopad", "studeni", "prosinac"]
found = {}
for y in range(2019, 2027):
    ok = []
    for i, m in enumerate(MON):
        for pat in (f"https://files.hrote.hr/files/PDF/OIEIK/{y}/Izvjestaji/Mjesecni_izvjestaj_{quote(m)}_{y}.pdf",):
            try:
                r = requests.get(pat, headers=H, timeout=40)
                if r.status_code == 200 and r.content[:4] == b"%PDF":
                    ok.append(i + 1); found[(y, i + 1)] = r.content
            except Exception as e:
                p("ERR", pat[-60:], type(e).__name__)
    p(y, "months found:", ok)
for k in [(2021, 1), (2022, 6), (max(found) if found else None)]:
    if k in found:
        d = fitz.open(stream=found[k], filetype="pdf"); p("-----", k, "pages", len(d))
        p(re.sub(r"\s+", " ", d[0].get_text())[:1500])
r = requests.get("https://www.hrote.hr/2025-godina", headers=H, timeout=40)
t = r.text
i = t.find("OIEIK")
p("2025-godina size", len(t), "OIEIK refs:", re.findall(r'[^"\']*OIEIK[^"\']*', t)[:10])
p(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t[t.find("<main"):]))[:1500])
r = requests.get("https://www.ere.gov.al/media/files/2026/03/31/Të_dhënat_e_tregut_të_energjisë_12-_Mujor_2025.pdf", headers=H, timeout=90)
p("ERE", r.status_code, len(r.content))
if r.content[:4] == b"%PDF":
    d = fitz.open(stream=r.content, filetype="pdf"); p("pages", len(d))
    for i in range(len(d)):
        tx = d[i].get_text()
        if re.search(r"(?i)(fuqi|kapacitet|MW)", tx) and re.search(r"(?i)(instaluar|install)", tx): p(f"[p{i+1}]", re.sub(r"\s+", " ", tx)[:1500])
