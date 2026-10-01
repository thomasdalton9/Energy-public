"""
Sixth NBS discovery pass (GitHub Actions only): the China Statistical
Yearbook's energy chapter (annual energy production / consumption /
balance tables) - is it pullable as data (xls) rather than images?
DISCOVERY3 got the 2023-2025 yearbook menus (left.htm, ~60 KB) but
found no "能源"/"C09" links in them, so this pass prints the menu's raw
structure around the energy chapter and tries the usual table URL
patterns (html/C09-01.xls/.xlsx/.jpg/.htm, English E09-01).
"""

import re
import time

import requests

S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                                "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"})


def get(url):
    time.sleep(1.0)
    try:
        r = S.get(url, timeout=(10, 30))
    except requests.RequestException as e:
        print(f"  {url}: {type(e).__name__}")
        return None
    ct = r.headers.get("content-type", "")
    chall = "Please enable JavaScript" in r.text[:3000] if "html" in ct else False
    print(f"  {url}: {r.status_code} {len(r.content)}B {ct}{' CHALLENGE' if chall else ''}")
    return None if chall else r


def main():
    for y in (2025, 2024):
        base = f"https://www.stats.gov.cn/sj/ndsj/{y}/"
        r = get(base + "left.htm")
        if r is not None and r.status_code == 200:
            txt = r.content.decode("utf-8", "replace")
            for enc in ("gbk", "gb18030"):
                if "能源" not in txt:
                    txt = r.content.decode(enc, "replace")
            i = txt.find("能源")
            print(f"  '能源' at {i}; snippet:\n{txt[max(0, i - 1500):i + 2500]}")
            print("  hrefs with 09:", sorted(set(re.findall(r'(?:href|src)=["\']([^"\']*0?9-?\d+[^"\']*)["\']', txt)))[:40])
        for stem in ("C09-01", "C09-02", "C09-1", "E09-01", "c09-01"):
            for ext in ("xls", "xlsx", "jpg", "htm", "html"):
                get(f"{base}html/{stem}.{ext}")


if __name__ == "__main__":
    main()
