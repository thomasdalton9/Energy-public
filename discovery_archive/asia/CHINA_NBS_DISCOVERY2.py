"""
Second NBS discovery pass (run in GitHub Actions - stats.gov.cn is
blocked from the dev sandbox). Goal: find EVERYTHING energy/industry-
relevant NBS publishes that we can pull, and how far back it goes.

Checks:
  1. English press-release index (www.stats.gov.cn/english/PressRelease/)
     crawled deep: every distinct release type + earliest edition found.
  2. Table layout of "Industrial Production Operation" (English) at
     sample dates across 2019-2026 - the generation-by-source rows are
     missing before Sep 2025 in the current pull, see why.
  3. Chinese latest-release index (www.stats.gov.cn/sj/zxfb/): every
     distinct release type + earliest edition; table dumps of the
     energy production, industrial value-added (product output table)
     and 10-day producer-goods market price releases.
  4. data.stats.gov.cn query engine, re-probed with a browser-like
     session (cookie + referer) in case the 403 UrlACL was header-driven.
"""

import collections
import json
import re
import sys
import time

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}
S = requests.Session()
S.headers.update(HEADERS)
EN_BASE = "https://www.stats.gov.cn/english/PressRelease/"
CN_BASE = "https://www.stats.gov.cn/sj/zxfb/"


def get(url, **kw):
    for a in range(3):
        try:
            r = S.get(url, timeout=(10, 30), **kw)
            r.encoding = r.apparent_encoding if "charset" not in r.headers.get("content-type", "") else r.encoding
            return r
        except requests.RequestException as e:
            print(f"   ERR {url}: {type(e).__name__}", flush=True)
            time.sleep(5)
    return None


def links(html, base):
    out = []
    for m in re.finditer(r"<a\b([^>]*)>(.*?)</a>", html, re.S):
        attrs, inner = m.group(1), m.group(2)
        h = re.search(r'href="([^"]*)"', attrs)
        t = re.search(r'title="([^"]*)"', attrs)
        title = t.group(1) if t else re.sub(r"<[^>]+>", "", inner).strip()
        if not h or not title:
            continue
        href = h.group(1)
        if href.startswith("./"):
            href = base + href[2:]
        if "/t20" not in href and "/t19" not in href:
            continue
        out.append((title.strip(), href))
    return out


def crawl(base, max_pages, label):
    seen = {}
    fails = 0
    last_page = 0
    for i in range(max_pages):
        url = base if i == 0 else f"{base}index_{i}.html"
        r = get(url)
        if r is None or r.status_code != 200:
            fails += 1
            print(f"  {label} page {i}: {None if r is None else r.status_code}", flush=True)
            if fails >= 4:
                break
            continue
        fails = 0
        last_page = i
        new = 0
        for t, h in links(r.text, base):
            if h not in seen:
                seen[h] = t
                new += 1
        if i % 20 == 0:
            print(f"  {label} page {i}: {new} new links, sample: {list(seen.values())[-1][:60]}", flush=True)
    print(f"  {label}: crawled to page {last_page}, {len(seen)} links", flush=True)
    return seen


def url_date(h):
    m = re.search(r"/t(\d{8})_", h)
    return m.group(1) if m else "?"


def summarize(seen, norm):
    groups = collections.defaultdict(list)
    for h, t in seen.items():
        groups[norm(t)].append((url_date(h), t, h))
    for k in sorted(groups, key=lambda k: -len(groups[k])):
        v = sorted(groups[k])
        print(f"  [{len(v):3d}] {k[:90]}  | {v[0][0]} .. {v[-1][0]} | earliest: {v[0][2]}")
    return groups


def table_rows(html):
    rows = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S | re.I):
        cells = [re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", c)).replace("&nbsp;", " ").strip()
                 for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.S | re.I)]
        cells = [c for c in cells if c]
        if cells:
            rows.append(cells)
    return rows


DUMPED = set()


def dump(url, max_rows=120, text_chars=0):
    if url in DUMPED:
        return
    DUMPED.add(url)
    r = get(url)
    print(f"\n----- {url} status={None if r is None else r.status_code}", flush=True)
    if r is None or r.status_code != 200:
        return
    rows = table_rows(r.text)
    print(f"  {len(rows)} table rows")
    for row in rows[:max_rows]:
        print("   | " + " | ".join(row)[:200])
    if text_chars:
        txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
        i = max(0, txt.find("In "))
        print("  TEXT: " + txt[i:i + text_chars])
    # attachments (xls/pdf) linked from the release
    att = re.findall(r'href="([^"]+\.(?:xls|xlsx|pdf|doc|docx|csv))"', r.text, re.I)
    if att:
        print("  ATTACHMENTS:", att[:10])


def pick(groups, key, years):
    out = []
    v = sorted(groups.get(key, []))
    for y in years:
        cand = [x for x in v if x[0].startswith(y)]
        if cand:
            out.append(cand[len(cand) // 2][2])
    return out


def en_norm(t):
    return re.sub(r"\s+", " ", re.sub(r"(January|February|March|April|May|June|July|August|September|October|"
                                      r"November|December|First|Second|Third|Fourth|Quarter|Half|\d{4}|\d+)", "#", t)).strip()


def cn_norm(t):
    return re.sub(r"[\d一二三四五六七八九十〇上下中]+(年|月份|月|季度|旬|日)", "#", t)


def probe_easyquery():
    print("\n===== data.stats.gov.cn easyquery re-probe", flush=True)
    for root in ("https://data.stats.gov.cn/", "https://data.stats.gov.cn/english/"):
        r = get(root + "easyquery.htm?cn=A01")
        print(f"  landing {root}: {None if r is None else (r.status_code, len(r.content), r.text[:200])}")
        params = {"m": "QueryData", "dbcode": "hgyd", "rowcode": "zb", "colcode": "sj", "wds": "[]",
                  "dfwds": json.dumps([{"wdcode": "zb", "valuecode": "A03010101"}]), "k1": str(int(time.time() * 1000))}
        r = get(root + "easyquery.htm", params=params, headers={"Referer": root + "easyquery.htm?cn=A01",
                                                               "X-Requested-With": "XMLHttpRequest"})
        print(f"  query {root}: {None if r is None else (r.status_code, r.text[:400])}")
    for u in ("https://data.stats.gov.cn/dg/", "https://data.stats.gov.cn/search.htm?s=原煤",
              "https://www.stats.gov.cn/sj/", "https://www.stats.gov.cn/sj/ndsj/"):
        r = get(u)
        print(f"  {u}: {None if r is None else (r.status_code, len(r.content))}")


def main():
    print("===== ENGLISH press-release index", flush=True)
    en = crawl(EN_BASE, 300, "EN")
    eg = summarize(en, en_norm)
    ipo = next((k for k in eg if "Industrial Production Operation" in k), None)
    epk = next((k for k in eg if k.startswith("Energy Production")), None)
    years = ["2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025"]
    if ipo:
        print(f"\n===== EN table samples: {ipo}")
        for u in pick(eg, ipo, years):
            dump(u, 80)
    if epk:
        print(f"\n===== EN energy production samples: {epk}")
        for u in pick(eg, epk, ["2019", "2021", "2023", "2025"]):
            dump(u, 40, 1500)

    print("\n===== CHINESE zxfb index", flush=True)
    cn = crawl(CN_BASE, 400, "CN")
    cg = summarize(cn, cn_norm)
    want = ["能源生产", "规模以上工业", "工业生产", "生产资料市场价格", "工业生产者", "产能利用", "能源"]
    for w in want:
        for k in [k for k in cg if w in k][:3]:
            print(f"\n===== CN samples: {k}")
            for u in pick(cg, k, ["2019", "2021", "2023", "2025", "2026"])[:4]:
                dump(u, 150)
    probe_easyquery()


if __name__ == "__main__":
    main()
