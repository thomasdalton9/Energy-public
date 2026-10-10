"""
NBS audit probe 1 (GitHub Actions only). 1) Crawl the whole /sj/zxfb/ release list and save every
(title, url) so release types / missing months can be counted. 2) Try the older /tjsj/zxfb/ archive.
3) Dump the full text + table rows of one sample release per type.
Results: discovery_archive/results/china_nbs_audit/
"""
import os
import re
import sys

import requests

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "asia"))
import china_nbs_common as nbs  # noqa: E402

OUT = os.path.join(HERE, "..", "results", "china_nbs_audit")
os.makedirs(OUT, exist_ok=True)


def save(name, text):
    with open(os.path.join(OUT, name), "w", encoding="utf-8") as f:
        f.write(text)


def page_text(html):
    h = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", html, flags=re.S | re.I)
    h = re.sub(r"</(p|div|tr|h\d|li)>", "\n", h, flags=re.I)
    h = re.sub(r"</t[dh]>", " | ", h, flags=re.I)
    h = re.sub(r"<[^>]+>", "", h).replace("&nbsp;", " ")
    h = re.sub(r"[ \t　]+", " ", h)
    return re.sub(r"\n\s*\n+", "\n", h)


# 1) full list
items = []
for i in range(nbs.MAX_INDEX_PAGES):
    url = nbs.CN_INDEX if i == 0 else f"{nbs.CN_INDEX}index_{i}.html"
    try:
        html = nbs.fetch(url)
    except Exception as e:  # noqa: BLE001
        nbs.log(f"page {i}: {type(e).__name__}")
        break
    if html is None:
        nbs.log(f"page {i}: 404 end")
        break
    for t, h in nbs._links(html, url.rsplit("/", 1)[0] + "/"):
        items.append((t, h))
seen, uniq = set(), []
for t, h in items:
    if h not in seen:
        seen.add(h)
        uniq.append((t, h))
save("release_list.tsv", "\n".join(f"{t}\t{h}" for t, h in uniq))
print(f"list: {len(uniq)} releases")

# 2) older archive
for u in ("https://www.stats.gov.cn/tjsj/zxfb/", "https://www.stats.gov.cn/tjsj/zxfb/index_1.html",
          "https://www.stats.gov.cn/sj/sjjd/", "https://www.stats.gov.cn/sj/zxfb/index_66.html"):
    try:
        html = nbs.fetch(u)
        ls = nbs._links(html or "", u.rsplit("/", 1)[0] + "/") if html else []
        print(u, "->", "404" if html is None else f"{len(html)} bytes, {len(ls)} links, first: {ls[:2]}")
    except Exception as e:  # noqa: BLE001
        print(u, "->", type(e).__name__)

# 3) sample releases
PATS = {
    "industrial_latest": r"规模以上工业增加值",
    "energy_production": r"能源生产情况",
    "cpi": r"居民消费价格",
    "ppi": r"工业生产者出厂价格",
    "capacity": r"工业产能利用率",
    "profits": r"工业企业利润",
    "national_economy": r"国民经济",
    "fai": r"固定资产投资",
    "retail": r"社会消费品零售总额",
    "xun_prices": r"生产资料市场价格",
    "price_index_70": r"70个大中城市",
}
want = {}
for t, h in uniq:
    for k, rx in PATS.items():
        if re.search(rx, t):
            want.setdefault(k, []).append((t, h))
for k, lst in want.items():
    print(k, len(lst), "newest:", lst[0][0], "oldest:", lst[-1][0])
pick = []
for k, lst in want.items():
    pick.append((k + "_newest", lst[0]))
    if k in ("xun_prices", "ppi", "industrial_latest", "cpi") and len(lst) > 1:
        pick.append((k + "_oldest", lst[-1]))
for k, lst in want.items():
    if k == "xun_prices":
        for t, h in lst:
            if re.search(r"2025年12月下旬|2026年1月上旬|2025年12月中旬", t):
                pick.append(("xun_" + re.sub(r"\W", "", t)[:14], (t, h)))
    if k == "industrial_latest":
        for t, h in lst:
            if "1—2月" in t or "1-2月" in t:
                pick.append(("industrial_janfeb_" + t[:5], (t, h)))
                break
for name, (t, h) in pick:
    try:
        html = nbs.fetch(h) or ""
    except Exception as e:  # noqa: BLE001
        print(name, "fetch failed", type(e).__name__)
        continue
    txt = f"TITLE {t}\nURL {h}\n\n" + page_text(html)
    save(f"sample_{name}.txt", txt[:90000])
    print(name, t, len(txt))
