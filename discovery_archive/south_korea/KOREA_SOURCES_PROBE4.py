"""
Probe 4 (manual workflow south_korea_probe4.yml): gas / LNG / oil sources - KESIS (energy statistics monthly), Petronet
(KNOC) v4 menu, KOGAS site menu (statistics, public data), KOGAS LNG import pages, Korea Customs trade statistics.
Prints text, board rows and attachment links.
"""
import re
import concurrent.futures as cf
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8"}
URLS = [
    "https://www.kesis.net/board.es?mid=a10503000000&bid=0037",
    "https://www.kesis.net/menu.es?mid=a10100000000",
    "https://www.kesis.net/menu.es?mid=a10201010100",
    "https://www.kesis.net/menu.es?mid=a10202010100",
    "https://www.kesis.net/menu.es?mid=a10300000000",
    "https://www.petronet.co.kr/v4/main.jsp",
    "https://www.petronet.co.kr/v4/eng/main.jsp",
    "https://www.kogas.or.kr/site/koGas/1050101000000",
    "https://www.kogas.or.kr/site/koGas/1010100000000",
    "https://www.kogas.or.kr/site/koGas/1040200000000",
    "https://www.kogas.or.kr/site/koGas/1020301000000",
    "https://www.kogas.or.kr/site/koGas/1051000000000",
    "https://www.kogas.or.kr/site/koGas/1051100000000",
    "https://www.kogas.or.kr/site/koGas/1051200000000",
    "https://www.kogas.or.kr/site/koGas/1050901010000",
    "https://www.kogas.or.kr/site/koGas/1050601000000",
    "https://www.kogas.or.kr/site/koGas/1040703000000",
    "https://www.kogas.or.kr/site/koGas/1040202000000",
    "https://unipass.customs.go.kr/ets/index_eng.do",
    "https://www.customs.go.kr/english/main.do",
]
LK = re.compile(r'<a[^>]+href="([^"#]+)"[^>]*>(.*?)</a>', re.S)


def one(u):
    out = [f"=== {u}"]
    try:
        r = requests.get(u, headers=H, timeout=(10, 25))
        if r.encoding in (None, "ISO-8859-1"):
            r.encoding = r.apparent_encoding
        t = r.text
        out.append(f"  status {r.status_code} bytes {len(r.content)} type {r.headers.get('content-type')}")
        body = re.sub(r"(?s)<(script|style).*?</\1>", " ", t)
        out.append("  text: " + re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", body))[:900])
        kw = re.compile(r"(통계|월보|통계표|수급|LNG|도입|수입|재고|판매|download|file|xls|hwp|pdf|attach|board|statis)", re.I)
        rows = []
        for h, txt in LK.findall(t):
            txt = re.sub(r"<[^>]+>|\s+", " ", txt).strip()
            if kw.search(h + txt) and txt:
                rows.append(f"{txt[:50]} -> {h[:140]}")
        out.append("  links:\n    " + "\n    ".join(dict.fromkeys(rows[:45])))
    except Exception as e:  # noqa: BLE001
        out.append(f"  ERROR {type(e).__name__}: {str(e)[:150]}")
    return "\n".join(out)


with cf.ThreadPoolExecutor(10) as ex:
    for res in ex.map(one, URLS):
        print(res, flush=True)
