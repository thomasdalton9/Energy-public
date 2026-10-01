"""
Third NBS discovery pass (GitHub Actions only). DISCOVERY2 found the
English press-release index stops at Apr 2024 and the Chinese /sj/zxfb/
index at ~2022/23, and data.stats.gov.cn is IP-blocked (403 UrlACL).
This pass looks for deeper history:
  1. older Chinese release indexes (信息公开 /xxgk/sjfb/zxfb2020/, the
     pre-2023 /tjsj/zxfb/), with year coverage of the energy-relevant
     release types;
  2. what the oldest pages of /sj/zxfb/ actually hold (migrated items);
  3. the statistical yearbook (/sj/ndsj/) energy chapter, for annual
     energy production/consumption tables;
  4. raw text of the Chinese monthly energy production release.
"""

import collections
import re
import time

import requests

S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                                "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"})
KEYS = {"IPO": "规模以上工业增加值", "ENERGY": "能源生产情况", "PRICES": "流通领域重要生产资料市场价格",
        "CAPU": "产能利用率", "PPI": "工业生产者出厂价格"}


def get(url):
    for _ in range(3):
        try:
            r = S.get(url, timeout=(10, 30))
            r.encoding = "utf-8"
            return r
        except requests.RequestException as e:
            print(f"   ERR {url}: {type(e).__name__}")
            time.sleep(5)
    return None


def links(html, base):
    out = []
    for m in re.finditer(r"<a\b([^>]*)>(.*?)</a>", html, re.S):
        attrs, inner = m.group(1), m.group(2)
        h = re.search(r'href="([^"]*)"', attrs)
        t = re.search(r'title="([^"]*)"', attrs)
        title = (t.group(1) if t else re.sub(r"<[^>]+>", "", inner)).strip()
        if h and title:
            href = h.group(1)
            if href.startswith("./"):
                href = base + href[2:]
            elif href.startswith("../"):
                href = base.rsplit("/", 2)[0] + "/" + href[3:]
            out.append((title, href))
    return out


def crawl(base, max_pages, show_tail=False):
    seen = {}
    fails = 0
    pages = []
    for i in range(max_pages):
        url = base if i == 0 else f"{base}index_{i}.html"
        r = get(url)
        if r is None or r.status_code != 200:
            print(f"  page {i}: {None if r is None else r.status_code}")
            fails += 1
            if fails >= 3:
                break
            continue
        fails = 0
        pl = [(t, h) for t, h in links(r.text, base) if re.search(r"/t\d{8}_", h)]
        pages.append(pl)
        for t, h in pl:
            seen.setdefault(h, t)
    print(f"  {base}: {len(pages)} pages, {len(seen)} release links")
    if show_tail:
        for pl in pages[-3:]:
            for t, h in pl[:20]:
                print(f"     {t[:60]} | {h}")
    for k, kw in KEYS.items():
        hits = sorted((re.search(r"/t(\d{8})_", h).group(1), t) for h, t in seen.items() if kw in t)
        years = collections.Counter(re.search(r"(20\d\d)年", t).group(1) if re.search(r"(20\d\d)年", t) else "?"
                                    for _, t in hits)
        print(f"  {k}: {len(hits)} by title-year {dict(sorted(years.items()))}")
        for d, t in hits[-3:]:
            print(f"     oldest-in-list: {d} {t[:60]}")
        for d, t in hits[:2]:
            print(f"     newest: {d} {t[:60]}")
    return seen


def main():
    print("===== /sj/zxfb/ tail pages")
    crawl("https://www.stats.gov.cn/sj/zxfb/", 70, show_tail=True)
    for base in ("https://www.stats.gov.cn/xxgk/sjfb/zxfb2020/", "https://www.stats.gov.cn/tjsj/zxfb/",
                 "https://www.stats.gov.cn/xxgk/sjfb/", "https://www.stats.gov.cn/sj/sjjd/"):
        print(f"\n===== {base}")
        crawl(base, 200, show_tail=True)

    print("\n===== /sj/ and /sj/ndsj/ links")
    for u in ("https://www.stats.gov.cn/sj/", "https://www.stats.gov.cn/sj/ndsj/"):
        r = get(u)
        if r is not None and r.status_code == 200:
            for t, h in links(r.text, u)[:150]:
                if any(k in t for k in ("年鉴", "月", "数据", "能源", "20")):
                    print(f"   {t[:50]} | {h}")
    for y in (2025, 2024, 2023):
        for u in (f"https://www.stats.gov.cn/sj/ndsj/{y}/indexch.htm", f"https://www.stats.gov.cn/sj/ndsj/{y}/left.htm"):
            r = get(u)
            print(f"  yearbook {u}: {None if r is None else (r.status_code, len(r.content))}")
            if r is not None and r.status_code == 200:
                for t, h in links(r.text, u.rsplit("/", 1)[0] + "/"):
                    if "能源" in t or re.search(r"html/[CE]09", h):
                        print(f"     {t[:60]} | {h}")
    print("\n===== Chinese energy production release text")
    r = get("https://www.stats.gov.cn/sj/zxfb/202609/t20260915_1965307.html")
    for u in ("https://www.stats.gov.cn/sj/zxfb/202606/t20260616_1963948.html",):
        r = get(u)
        if r is not None:
            txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
            i = txt.find("一、")
            print(txt[i:i + 3500])
            print("  IMGS:", re.findall(r'<img[^>]+src="([^"]+)"', r.text)[:12])
            print("  ATT:", re.findall(r'href="([^"]+\.(?:xls|xlsx|pdf))"', r.text)[:12])


if __name__ == "__main__":
    main()
