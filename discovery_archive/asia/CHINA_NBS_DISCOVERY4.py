"""
Fourth NBS discovery pass (GitHub Actions only). DISCOVERY3 showed the
Chinese /sj/zxfb/ index is capped at ~1000 items (67 pages), reaching
back only to Oct 2021. Older releases were migrated to the new site on
2023-02-03 with sequential IDs (/sj/zxfb/202302/t20230203_19012xx.html
for Oct 2021) - this pass scans IDs below the oldest indexed one to see
whether Jan-Sep 2021 (and earlier) releases still exist unlisted.
Also inspects the statistical yearbook's left-hand table menu (energy
chapter) for annual energy tables.
"""

import collections
import concurrent.futures as cf
import re

import requests

S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                                "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"})
S.mount("https://", requests.adapters.HTTPAdapter(pool_connections=16, pool_maxsize=16))
KEYS = ("规模以上工业增加值", "能源生产", "流通领域重要生产资料", "产能利用率", "工业生产者出厂价格")


def title_of(i, date):
    url = f"https://www.stats.gov.cn/sj/zxfb/202302/t{date}_{i}.html"
    try:
        r = S.get(url, timeout=(10, 20))
    except requests.RequestException:
        return i, url, None
    if r.status_code != 200:
        return i, url, None
    r.encoding = "utf-8"
    m = re.search(r"<title>(.*?)</title>", r.text, re.S)
    return i, url, (m.group(1).strip() if m else "?")


def scan(lo, hi, date):
    found = []
    with cf.ThreadPoolExecutor(12) as ex:
        for i, url, t in ex.map(lambda i: title_of(i, date), range(lo, hi)):
            if t:
                found.append((i, url, t))
    print(f"  date {date} ids {lo}-{hi}: {len(found)} pages exist")
    years = collections.Counter((re.findall(r"((?:19|20)\d\d)年", t) or ["?"])[0] for _, _, t in found)
    print(f"  title years: {dict(sorted(years.items()))}")
    for i, url, t in found:
        if any(k in t for k in KEYS):
            print(f"   {i} {t[:70]} | {url}")
    if found:
        print(f"  lowest id found: {found[0][0]} {found[0][2][:60]}")
    return found


def main():
    print("===== ID scan of migrated releases (20230203)")
    scan(1897500, 1901240, "20230203")
    print("\n===== ID scan (20230202)")
    scan(1896500, 1897600, "20230202")

    print("\n===== yearbook menu")
    for y in (2025,):
        r = S.get(f"https://www.stats.gov.cn/sj/ndsj/{y}/left.htm", timeout=30)
        for enc in ("utf-8", "gbk"):
            try:
                txt = r.content.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        print(f"  decoded as {enc}; head: {re.sub(chr(10), ' ', txt[:600])}")
        hits = [m for m in re.finditer(r"<a[^>]*href=\"([^\"]+)\"[^>]*>(.*?)</a>", txt, re.S)]
        print(f"  {len(hits)} links; sample: {[h.group(1) for h in hits[:5]]}")
        for m in hits:
            t = re.sub(r"<[^>]+>|\s+", "", m.group(2))
            if "能源" in t or "电力" in t or "煤" in t:
                print(f"     {t[:60]} | {m.group(1)}")


if __name__ == "__main__":
    main()
