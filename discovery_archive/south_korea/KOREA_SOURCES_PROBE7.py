"""
Probe 7 (manual workflow south_korea_probe4.yml): gas / LNG / oil pages, compact output - KOGAS import/sales pages and
boards, KESIS board, Petronet v4 menu. Prints main-content text and attachment links only.
"""
import re
import concurrent.futures as cf
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8"}
K = "https://www.kogas.or.kr"
URLS = [
    K + "/site/koGas/1030302000000", K + "/site/koGas/1030304000000", K + "/site/koGas/1040301000000",
    K + "/site/koGas/1040302000000", K + "/site/koGas/goBoard.do?boardNo=27&Key=1040602000000",
    K + "/site/koGas/goBoard.do?boardNo=30&Key=1040603000000",
    K + "/site/koGas/goBoard.do?boardNo=102&Key=1040601000000",
    "https://www.kesis.net/board.es?mid=a10503000000&bid=0037",
    "https://www.kesis.net/menu.es?mid=a10201010100",
    "https://www.kesis.net/menu.es?mid=a10202010100",
    "https://www.petronet.co.kr/v4/main.jsp",
    "https://www.petronet.co.kr/v4/eng/main.jsp",
]
FILE = re.compile(r'href="([^"]*(?:fileDown|download|\.xls|\.xlsx|\.csv|\.pdf|\.hwp|atch|file)[^"]*)"', re.I)


def one(u):
    out = [f"=== {u}"]
    try:
        r = requests.get(u, headers=H, timeout=(10, 30))
        if r.encoding in (None, "ISO-8859-1"):
            r.encoding = r.apparent_encoding
        t = r.text
        out.append(f"  status {r.status_code} bytes {len(r.content)}")
        body = re.sub(r"(?s)<(script|style).*?</\1>", " ", t)
        body = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", body))
        for marker in ("트위터 블로그", "본문 바로가기", "본문으로"):
            j = body.rfind(marker)
            if j > 0:
                body = body[j:]
                break
        out.append("  main: " + body[:1500])
        fl = list(dict.fromkeys(FILE.findall(t)))[:15]
        out.append("  files: " + " | ".join(x[:120] for x in fl))
        if "petronet" in u:
            items = re.findall(r'<a[^>]+href="([^"#]+)"[^>]*>([^<]{2,40})</a>', t)
            keep = [f"{a.strip()}->{h[:60]}" for h, a in items if re.search(r"(재고|수출입|수급|LNG|가스|도입|원유|수입)", a)]
            out.append("  menu: " + " | ".join(dict.fromkeys(keep[:40])))
        if "kesis" in u:
            items = re.findall(r'<a[^>]+href="([^"#]+)"[^>]*>(.*?)</a>', t, re.S)
            keep = [re.sub(r"<[^>]+>|\s+", " ", a).strip()[:50] + "->" + h[:70] for h, a in items if "board" in h or "menu.es" in h]
            out.append("  links: " + " | ".join(dict.fromkeys(keep[:40])))
    except Exception as e:  # noqa: BLE001
        out.append(f"  ERROR {type(e).__name__}: {str(e)[:150]}")
    return "\n".join(out)


with cf.ThreadPoolExecutor(6) as ex:
    for res in ex.map(one, URLS):
        print(res, flush=True)
