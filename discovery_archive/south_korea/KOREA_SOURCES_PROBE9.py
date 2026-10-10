"""
Probe 9 (manual workflow south_korea_probe4.yml): KOGAS LNG import/sales pages (main text only) and Petronet English menu
anchors (href/onclick) for crude/product import and stock tables.
"""
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8"}
K = "https://www.kogas.or.kr/site/koGas/"
for key in ("1030302000000", "1030304000000", "1040301000000", "1040302000000"):
    r = requests.get(K + key, headers=H, timeout=(10, 40))
    t = re.sub(r"(?s)<(script|style).*?</\1>", " ", r.text)
    t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))
    j = t.rfind("블로그 ")
    j2 = t.find("카카오스토리", 0)
    body = t[j:] if j > 0 else t
    print("=== KOGAS", key, r.status_code, "main:", body[:1300], flush=True)
    tabs = re.findall(r"<table[^>]*>", r.text)
    print("   tables:", len(tabs), "| numbers sample:", re.findall(r"\d{1,3}(?:,\d{3})+", body)[:12], flush=True)

r = requests.get("https://www.petronet.co.kr/v4/eng/main.jsp", headers=H, timeout=(10, 40))
r.encoding = "utf-8"
anchors = re.findall(r"<a\b[^>]*>[^<]*(?:Import|Stock|Export|Demand|Production)[^<]*</a>", r.text)
print("PETRONET eng anchors:", len(anchors))
for a in anchors[:30]:
    print("  ", re.sub(r"\s+", " ", a)[:260])
print("PETRONET js menu fn:", re.findall(r"(?:goMenu|menuMove|fn_\w+|go\w+)\(['\"][^)]{1,80}\)", r.text)[:15])
