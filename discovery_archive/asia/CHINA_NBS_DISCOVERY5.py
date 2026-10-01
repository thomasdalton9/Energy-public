"""
Fifth NBS discovery pass (GitHub Actions only). DISCOVERY4's parallel ID
scan tripped stats.gov.cn's anti-bot JS challenge ("Please enable
JavaScript" served with HTTP 200), so most "found" pages were the
challenge, not releases. It did show migrated release IDs run roughly in
date order (Dec 2012 PMI at 1898246, Oct 2021 at ~1901240). This pass
scans the IDs just below the oldest indexed release SEQUENTIALLY and
slowly, detecting the challenge page, to find the Jan-Sep 2021 releases.
"""

import re
import sys
import time

import requests

S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                                "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"})
KEYS = ("规模以上工业增加值", "能源生产", "流通领域重要生产资料", "产能利用率")


def fetch(url):
    for attempt in range(4):
        try:
            r = S.get(url, timeout=(10, 20))
        except requests.RequestException:
            time.sleep(5)
            continue
        r.encoding = "utf-8"
        if r.status_code == 200 and "Please enable JavaScript" in r.text[:2000]:
            print(f"   challenge on {url}, backing off", flush=True)
            time.sleep(30 * (attempt + 1))
            continue
        return r
    return None


def main():
    lo, hi = int(sys.argv[1]) if len(sys.argv) > 1 else 1900800, int(sys.argv[2]) if len(sys.argv) > 2 else 1901240
    n = 0
    for i in range(hi, lo, -1):
        url = f"https://www.stats.gov.cn/sj/zxfb/202302/t20230203_{i}.html"
        r = fetch(url)
        time.sleep(0.6)
        if r is None:
            print(f"  {i}: giving up (challenge/timeouts)")
            break
        if r.status_code != 200:
            continue
        n += 1
        m = re.search(r"<title>(.*?)</title>", r.text, re.S)
        t = m.group(1).strip() if m else "?"
        if any(k in t for k in KEYS) or n % 25 == 0:
            print(f"  {i} {t[:70]}", flush=True)
    print(f"  {n} pages in range")


if __name__ == "__main__":
    main()
