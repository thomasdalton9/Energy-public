"""
Probe 11 (manual workflow south_korea_probe.yml): Petronet (KNOC) English menu loader - print loadMenu() and try the
page it fetches for crude imports by country and product stocks; save to discovery_archive/results/south_korea/.
"""
import os
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "en-US,en;q=0.9,ko;q=0.8"}
OUT = "discovery_archive/results/south_korea"
os.makedirs(OUT, exist_ok=True)
S = requests.Session()
S.headers.update(H)
B = "https://www.petronet.co.kr"
r = S.get(B + "/v4/eng/main.jsp", timeout=(10, 40))
r.encoding = "utf-8"
t = r.text
srcs = re.findall(r'<script[^>]+src="([^"]+)"', t)
print("scripts:", srcs)
i = t.find("function loadMenu")
print("inline loadMenu:", re.sub(r"\s+", " ", t[i:i + 900]) if i >= 0 else None)
log = []
for s in srcs:
    if s.startswith("/") or s.startswith("http") and "petronet" in s:
        u = s if s.startswith("http") else B + s
        try:
            js = S.get(u, timeout=(10, 40)).text
        except Exception as e:  # noqa: BLE001
            print("ERR", u, e)
            continue
        j = js.find("loadMenu")
        if j >= 0:
            print("JS", u, "loadMenu:", re.sub(r"\s+", " ", js[max(0, j - 100):j + 1500]))
            log.append(u)
open(os.path.join(OUT, "petronet_scripts.txt"), "w").write("\n".join(srcs))
# guesses for the content page
for q in ["/v4/eng/main2.jsp?menu1=EngKDXQ&menu2=EngKDXQ04&menu3=EngKDXQ0800",
          "/v4/eng/sub.jsp?menu1=EngKDXQ&menu2=EngKDXQ04&menu3=EngKDXQ0800",
          "/v4/eng/KDXQ0800.jsp", "/v4/eng/menu.jsp?menu=EngKDXQ0800"]:
    try:
        rr = S.get(B + q, timeout=(10, 30))
        print("GUESS", q, rr.status_code, len(rr.text), re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", rr.text))[:200])
    except Exception as e:  # noqa: BLE001
        print("GUESS ERR", q, type(e).__name__)
