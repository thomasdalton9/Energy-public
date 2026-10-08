"""
Shared helpers for the two National Energy Administration (NEA, nea.gov.cn) monthly pulls:
  asia/CHINA_NEA_CAPACITY.py     (全国电力(工业)统计数据: installed and added capacity by type)
  asia/CHINA_NEA_CONSUMPTION.py  (全社会用电量: electricity consumption by sector)

How the NEA releases are found (discovery_archive/nea/NEA_HUNT*.py, results in discovery_archive/results/nea/): the press list
https://www.nea.gov.cn/xwfb/index.htm is script-rendered, but its page names the list's data source
(data="datasource:<id>") and the script loads ./ds_<id>.json - one JSON file holding the WHOLE list (title, publishTime,
publishUrl) - so there is no 2-year hole behind the paged list. The department column /sjzz/ghs/ (综合司) carries the same
releases, with another ds_<id>.json; both are read and merged. The ids are re-read from the column page on every run.
"""
import re
import sys
import time

import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36", "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8"}
SESSION = requests.Session()
SESSION.headers.update(HEADERS)
SECTIONS = ["/xwfb/", "/sjzz/ghs/"]
SITE = "https://www.nea.gov.cn"
CAP_RE = re.compile(r"全国电力(?:工业)?统计数据")
CONS_RE = re.compile(r"全社会用电量")
_last = [0.0]


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def fetch(url, attempts=4, binary=False):
    """GET with retries; None on 404."""
    err = None
    for attempt in range(attempts):
        time.sleep(max(0.0, 0.6 - (time.time() - _last[0])))
        try:
            r = SESSION.get(url, timeout=(10, 30))
            _last[0] = time.time()
        except requests.RequestException as e:
            err = e
            time.sleep(6 * (attempt + 1))
            continue
        if r.status_code == 404:
            return None
        if r.status_code != 200:
            err = requests.HTTPError(f"{r.status_code} {url}")
            time.sleep(6 * (attempt + 1))
            continue
        if binary:
            return r.content
        r.encoding = "utf-8"
        return r.text
    raise err


def clean(s):
    return re.sub(r"<[^>]+>", "", s or "").strip()


def page_text(html):
    """Page text, one line per paragraph / table row part, table cells separated by ' | '."""
    html = re.sub(r"(?s)<(script|style).*?</\1>", " ", html)
    html = re.sub(r"\s+", " ", html)
    html = re.sub(r"</(p|div|tr|li|h\d)>", "\n", html)
    html = re.sub(r"</t[dh]>", " | ", html)
    t = re.sub(r"<[^>]+>", "", html)
    t = re.sub(r"&nbsp;|&emsp;|　", " ", t)
    t = t.replace("&gt;", ">").replace("&lt;", "<").replace("&amp;", "&")
    return re.sub(r"\n\s*\n+", "\n", re.sub(r"[ \t]+", " ", t))


def article_key(url):
    m = re.search(r"/([0-9a-f]{32})/c\.html", url) or re.search(r"/(c_\d+)\.htm", url)
    return m.group(1) if m else url


def releases():
    """All capacity ('cap') and consumption ('cons') releases on the NEA site: [{kind, date, title, url}], oldest first."""
    from urllib.parse import urljoin
    out = {}
    for sec in SECTIONS:
        html = fetch(f"{SITE}{sec}index.htm")
        if not html:
            continue
        for ds in sorted(set(re.findall(r"datasource:([0-9a-f]{32})", html))):
            txt = fetch(f"{SITE}{sec}ds_{ds}.json")
            if not txt:
                continue
            try:
                import json
                items = json.loads(txt).get("datasource") or []
            except ValueError:
                continue
            for i in items:
                title = clean(i.get("title") or i.get("showTitle"))
                kind = "cap" if CAP_RE.search(title) else ("cons" if CONS_RE.search(title) else None)
                if not kind or not i.get("publishUrl"):
                    continue
                url = urljoin(f"{SITE}{sec}", i["publishUrl"])
                if url.startswith("http://www.nea.gov.cn/"):
                    url = "https://" + url[len("http://"):]
                key = (kind, article_key(url))
                if key not in out or str(i["publishTime"]) < out[key]["_t"]:   # keep the earliest stamp of a duplicate
                    out[key] = {"kind": kind, "date": str(i["publishTime"])[:10], "title": title, "url": url,
                                "_t": str(i["publishTime"])}
    rows = sorted(out.values(), key=lambda r: (r["date"], r["kind"]))
    for r in rows:
        r.pop("_t")
    return rows


def body_of(text):
    """The part of the page text after the 'source' line (drops navigation)."""
    i = text.find("来源")
    return text[i:] if i >= 0 else text
