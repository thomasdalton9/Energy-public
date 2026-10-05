"""
Probe: can Poland's Yamal/Kondratki (BY>PL) entry flows for Jan-May 2022 be found in a raw operator source? ENTSOG has no Kondratki/Wysokoje rows.
Fetches Gaz-System / EuRoPol Gaz / FGSZ-MEKH landing pages and prints the links that mention transparency / data / flows / statistics.
"""
import re
import requests

PAGES = ["https://www.gaz-system.pl/en/", "https://www.gaz-system.pl/pl/", "https://www.gaz-system.pl/en/transmission-system/transmission-services/",
         "https://en.gaz-system.pl/", "https://swo.gaz-system.pl", "https://europolgaz.com.pl/en", "https://www.europolgaz.com.pl/",
         "https://ebiznes.gaz-system.pl/", "https://www.ure.gov.pl/en/", "https://www.fgsz.hu/en/", "https://mekh.hu/en/"]
H = {"User-Agent": "Mozilla/5.0"}
for u in PAGES:
    try:
        r = requests.get(u, headers=H, timeout=25)
    except Exception as e:  # noqa: BLE001
        print("ERR", type(e).__name__, u, flush=True)
        continue
    print(r.status_code, len(r.text), u, flush=True)
    if r.status_code == 200:
        seen = set()
        for m in re.finditer(r'href="([^"]+)"[^>]*>([^<]{3,80})<', r.text):
            href, txt = m.group(1), m.group(2).strip()
            if re.search(r"transparen|flow|data|statist|dane|przep|kondratki|operational|nomination|allocation", href + txt, re.I) and href not in seen:
                seen.add(href)
                print("   ", txt[:60], "->", href[:140])
                if len(seen) > 25:
                    break
