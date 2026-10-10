"""
Probe 2 (manual workflow south_korea_probe.yml): find the data endpoints behind EPSIS (KPX) pages - every .do link on
the main page, the .do / .json strings inside the fuel-source generation and SMP pages and their JS files - plus the
KGU, KESIS and Petronet menus.
"""
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8"}
S = requests.Session()
S.headers.update(H)
B = "https://epsis.kpx.or.kr"


def get(u, **kw):
    try:
        r = S.get(u, timeout=(10, 30), **kw)
        print(f"GET {u} -> {r.status_code} {len(r.content)}", flush=True)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"GET {u} ERROR {type(e).__name__} {str(e)[:100]}", flush=True)


r = get(B + "/epsisnew/")
home = r.text
links = sorted(set(re.findall(r'href="(/epsisnew/[^"]+)"', home)))
print("HOME links:", len(links))
for l in links:
    print("  ", l)
# anchors with text
for h, t in re.findall(r'<a[^>]+href="(/epsisnew/[^"]+)"[^>]*>(.*?)</a>', home, re.S):
    t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
    if t:
        print("  A", h, "|", t[:40])

for path in ["/epsisnew/selectEkmaSmpSmpChart.do?menuId=040201", "/epsisnew/selectEkgeEpsEccChart.do?menuId=010300",
             "/epsisnew/selectEkifBftChart.do?menuId=030200"]:
    r = get(B + path)
    if not r or r.status_code != 200:
        continue
    t = r.text
    print("  strings .do:", sorted(set(re.findall(r"[\w/]+\.do[\w?=&]*", t)))[:40])
    print("  scripts:", re.findall(r'<script[^>]+src="([^"]+)"', t)[:20])
    print("  forms:", re.findall(r"<form[^>]*>", t)[:5])
    print("  inputs:", re.findall(r'<input[^>]+name="([^"]+)"', t)[:30])
    print("  selects:", re.findall(r'<select[^>]+(?:name|id)="([^"]+)"', t)[:30])
    m = re.search(r"(?s)(excel|Excel|엑셀).{0,300}", t)
    print("  excel ctx:", re.sub(r"\s+", " ", m.group(0))[:300] if m else None)
    for js in re.findall(r'<script[^>]+src="([^"]+)"', t):
        if "epsis" in js.lower() or "/js/" in js:
            if js.startswith("/"):
                jr = get(B + js)
                if jr and jr.status_code == 200:
                    print("   JS .do:", sorted(set(re.findall(r"['\"]([\w/]+\.do[\w?=&]*)['\"]", jr.text)))[:30])

# sitemap guess
for p in ["/epsisnew/selectEkmaSitemapForm.do", "/epsisnew/sitemap.do", "/epsisnew/selectEkmaSitemap.do"]:
    r = get(B + p)
    if r is not None and r.status_code == 200:
        print(sorted(set(re.findall(r'href="(/epsisnew/[^"]+)"', r.text)))[:120])
        break

for u in ["https://www.kgu.or.kr/center/domestic_lng", "https://www.kgu.or.kr/center/energy_statistics",
          "https://www.kesis.net/main/main.jsp", "https://www.petronet.co.kr/v3/index.jsp"]:
    r = get(u)
    if r is not None and r.status_code == 200:
        t = re.sub(r"(?s)<(script|style).*?</\1>", " ", r.text)
        print("  text:", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))[:700])
        print("  hrefs:", sorted(set(re.findall(r'href="([^"#]+)"', r.text)))[:60])
