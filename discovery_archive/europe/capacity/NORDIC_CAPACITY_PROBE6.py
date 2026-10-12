"""Probe 6 (12 Oct 2026): Litgrid installed-capacity pages and the 'Gamybos pajegumai' open data set; Elering consumption-and-generation page."""
import re
import requests
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 (energy-data probe)"
def text(r): return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", r.text, flags=re.S|re.I)))
for u in ["https://www.litgrid.eu/en/power-system/electric-power-system/installed-capacity",
          "https://www.litgrid.eu/sistema/elektros-energetikos-sistema/irengtoji-galia",
          "https://www.litgrid.eu/atviru-duomenu-rinkiniai/gamybos-pajegumai-mw-irengti-visuose-esamuose-gamybos-vienetuose-kuriu-instaliuotoji-generavimo-galia-sudaro-1-mw-arba-daugiau-suma-pagal-kiekviena-gamybos-rusi",
          "https://www.elering.ee/en/electricity-consumption-and-generation",
          "https://www.elering.ee/en/production-and-forecast"]:
    print("=====", u[:120])
    try:
        r = S.get(u, timeout=60); print(r.status_code, len(r.text))
        t = text(r); i = max(t.find("apacity"), t.find("galia"), 0)
        print(t[max(0, i-200):i+1800])
        print("LINKS", sorted(set(re.findall(r'(?:href|src|data-[a-z-]+)="([^"]*(?:\.xls|\.csv|\.json|api|download|file)[^"]*)"', r.text)))[:20])
    except Exception as e: print("ERR", type(e).__name__, e)
