"""
Grid-operator data discovery, round 5: EVN daily-post list (Vietnam). VIETNAM_EVN.py's first run found posts only
from 2025-10-01 and none in January 2026, and stopped at page 33. Print every post link (href + title) on the list
pages around January 2026 and past page 33, to see the older / differently named posts.
"""
import re

import requests

LIST = "https://www.evn.com.vn/vi-VN/news-l/Thong-tin-tom-tat-van-hanh-HTD-Quoc-gia-60-2015"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36"}


def main():
    for page in (20, 21, 22, 23, 24, 31, 32, 33, 34, 35, 36, 40, 50, 80, 120):
        try:
            t = requests.get(f"{LIST}?page={page}", headers=H, timeout=(15, 60)).text
        except Exception as e:  # noqa: BLE001
            print(f"page {page}: {e}", flush=True)
            continue
        links = []
        for m in re.finditer(r'<a[^>]+href="(/d/vi-VN/news/[^"]+)"[^>]*>(.*?)</a>', t, re.S):
            title = re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()
            links.append((m.group(1), title[:90]))
        seen = []
        for h, ti in links:
            if h not in [x for x, _ in seen]:
                seen.append((h, ti))
        print(f"\n==== page {page}: {len(seen)} links, {len(t)} bytes", flush=True)
        for h, ti in seen[:40]:
            print(f"  {h[:140]} | {ti}", flush=True)


if __name__ == "__main__":
    main()
