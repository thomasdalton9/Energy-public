"""Manual probe 3 (8 Oct 2026): UN Energy Statistics Yearbook 2023 Table I (energy unit conversion: kcal, Btu, kJ)."""
import os
import requests
import fitz
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "results", "china_fx")
UA = {"User-Agent": "Mozilla/5.0 (research script)"}
log = []
for n in ("09i", "09", "09iii", "08"):
    url = f"https://unstats.un.org/unsd/energystats/pubs/yearbook/2023/{n}.pdf"
    try:
        r = requests.get(url, headers=UA, timeout=60)
        log.append(f"{url} {r.status_code} {len(r.content)}")
        if r.status_code == 200 and r.content[:4] == b"%PDF":
            doc = fitz.open(stream=r.content, filetype="pdf")
            txt = "\n".join(p.get_text() for p in doc)
            open(os.path.join(OUT, f"probe3_un2023_{n}.txt"), "w").write(txt[:60000])
    except Exception as e:
        log.append(f"{url} {type(e).__name__}")
open(os.path.join(OUT, "probe3_log.txt"), "w").write("\n".join(log)); print("\n".join(log))
