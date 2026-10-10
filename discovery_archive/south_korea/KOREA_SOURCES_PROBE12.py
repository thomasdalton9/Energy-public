"""
Probe 12 (manual workflow south_korea_probe.yml): last gas / oil attempts, compact output, saved to
discovery_archive/results/south_korea/probe12.txt:
  - Petronet English sub.jsp (POST of the menu form) for crude import by country, products stock
  - KESIS monthly energy statistics bulletin pages (attachments)
  - KITA K-stat, data.go.kr search for KOGAS / LNG files
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
lines = []


def say(*a):
    s = " ".join(str(x) for x in a)
    print(s, flush=True)
    lines.append(s)


def txt(html, n=700):
    t = re.sub(r"(?s)<(script|style).*?</\1>", " ", html)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))[:n]


P = "https://www.petronet.co.kr"
for ids in (("EngKDXQ", "EngKDXQ04", "EngKDXQ0800", "04", "04_01", "04_01_01"),
            ("EngKDXQ", "EngKDPQ010", "EngKDPQ0100", "04", "04_02", "04_02_01")):
    try:
        r = S.post(P + "/v4/eng/sub.jsp", data=dict(zip(("fmuId", "smuId", "tmuId", "fmuOrd", "smuOrd", "tmuOrd"), ids)), timeout=(10, 40))
        r.encoding = "utf-8"
        say("PETRONET", ids[2], r.status_code, len(r.text))
        say("  text:", txt(r.text, 500))
        say("  ajax:", re.findall(r"url\s*:\s*['\"]([^'\"]+)['\"]", r.text)[:8], re.findall(r"\$\.(?:post|get|ajax)\(['\"]([^'\"]+)", r.text)[:8])
        say("  iframes/includes:", re.findall(r'<iframe[^>]+src="([^"]+)"', r.text)[:5], re.findall(r"(?:load|include)\(['\"]([^'\"]+)", r.text)[:5])
        say("  tables:", len(re.findall(r"<table", r.text)), " numbers:", re.findall(r"\d{1,3}(?:,\d{3})+", r.text)[:8])
    except Exception as e:  # noqa: BLE001
        say("PETRONET ERR", ids[2], type(e).__name__)

K = "https://www.kesis.net"
for u in ("/menu.es?mid=a10301010000", "/menu.es?mid=a10101000000"):
    try:
        r = S.get(K + u, timeout=(10, 40))
        say("KESIS", u, r.status_code, len(r.text))
        body = txt(r.text, 4000)
        j = body.find("전체메뉴 닫기")
        say("  text:", body[j:j + 700])
        say("  files:", re.findall(r'href="([^"]*(?:boardDownload|download|\.xls|\.pdf|\.hwp)[^"]*)"', r.text, re.I)[:10])
        say("  data:", re.findall(r'(?:ajax|url)\s*:\s*[\'"]([^\'"]+\.(?:do|es|json)[^\'"]*)', r.text)[:8])
    except Exception as e:  # noqa: BLE001
        say("KESIS ERR", u, type(e).__name__)

for u in ("https://stat.kita.net/stat/kts/pum/PumItemList.screen", "https://stat.kita.net/stat/kts/ctr/CtrTotalList.screen",
          "https://www.data.go.kr/tcs/dss/selectDataSetList.do?keyword=%ED%95%9C%EA%B5%AD%EA%B0%80%EC%8A%A4%EA%B3%B5%EC%82%AC+LNG",
          "https://www.data.go.kr/tcs/dss/selectDataSetList.do?keyword=LNG+%EB%8F%84%EC%9E%85"):
    try:
        r = S.get(u, timeout=(10, 40))
        say("GET", u[:110], r.status_code, len(r.text))
        say("  text:", txt(r.text, 500))
        say("  datasets:", re.findall(r'/data/(\d{7,8})/(?:fileData|openapi)\.do', r.text)[:12])
    except Exception as e:  # noqa: BLE001
        say("GET ERR", u[:80], type(e).__name__)
open(os.path.join(OUT, "probe12.txt"), "w", encoding="utf-8").write("\n".join(lines))
