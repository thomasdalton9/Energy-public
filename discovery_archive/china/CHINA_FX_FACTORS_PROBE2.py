"""Manual probe 2 (8 Oct 2026): structure of the Fed H.10 historical China page (the source series behind FRED DEXCHUS) and
the LNG tonne -> MMBtu factor in GIIGNL's annual reports. Results: discovery_archive/results/china_fx/probe2_*.txt"""
import os, re
import requests
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "results", "china_fx")
UA = {"User-Agent": "Mozilla/5.0 (research script)"}
log = []
def say(*a):
    s = " ".join(map(str, a)); print(s, flush=True); log.append(s)

r = requests.get("https://www.federalreserve.gov/releases/h10/hist/dat00_ch.htm", headers=UA, timeout=60)
t = r.text
say("h10", r.status_code, len(t))
i = t.find("<table")
say("tables", len(re.findall("<table", t)), "first table at", i)
say(t[i:i + 3000])
j = t.rfind("</table>")
say("....TAIL....")
say(t[max(0, j - 3000):j + 200])
for m in re.finditer(r"(?i)(noon buying|dollars? per|units of foreign|Yuan|Renminbi|Note)[^<]{0,200}", t[:i + 200000]):
    say("TXT:", m.group(0)[:250])
    if len(log) > 120: break

for name, url in [("giignl2026", "https://giignl-documents.s3.fr-par.scw.cloud/public/giignl-2026-annual-report.pdf"),
                  ("giignl2025", "https://giignl-documents.s3.fr-par.scw.cloud/public/ar-2025-annual-report.pdf")]:
    try:
        r = requests.get(url, headers=UA, timeout=120)
        say(name, r.status_code, len(r.content))
        import fitz
        doc = fitz.open(stream=r.content, filetype="pdf")
        txt = "\n".join(f"=== page {k+1} ===\n" + p.get_text() for k, p in enumerate(doc))
        open(os.path.join(OUT, f"probe2_{name}.txt"), "w").write(txt)
        for m in re.finditer(r"(?i)(mmbtu|million btu|btu)", txt):
            say(f"  {name} @{m.start()}:", txt[max(0, m.start()-200):m.end()+200].replace("\n", " | "))
    except Exception as e:
        say(name, "failed", type(e).__name__, e)
open(os.path.join(OUT, "probe2_log.txt"), "w").write("\n".join(log))
