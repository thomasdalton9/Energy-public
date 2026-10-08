"""
Probe (manual workflow china_sources_probe.yml): which ADDITIONAL official China energy sources answer from GitHub
Actions, beyond the NBS releases already pulled (asia/CHINA_NBS_*.py). Prints status, size, <title> and the links on
each page that look like energy tables (natural gas, imports, capacity, electricity consumption). Nothing is saved
or used by a pull: a source only becomes a pull once a later probe has opened and parsed its tables.
"""
import re
import sys
import time

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8"}
URLS = [
    # NBS annual (yearbook / energy balance), data portal
    ("NBS yearbook index", "https://www.stats.gov.cn/sj/ndsj/"),
    ("NBS yearbook 2025 index", "https://www.stats.gov.cn/sj/ndsj/2025/indexch.htm"),
    ("NBS yearbook 2024 index", "https://www.stats.gov.cn/sj/ndsj/2024/indexch.htm"),
    ("NBS energy releases list (CN)", "https://www.stats.gov.cn/sj/zxfb/"),
    ("NBS sjjd interpretations", "https://www.stats.gov.cn/sj/sjjd/"),
    # Customs (GACC)
    ("GACC statistics CN", "http://www.customs.gov.cn/customs/302249/zfxxgk/2799825/302274/302277/index.html"),
    ("GACC statistics CN https", "https://www.customs.gov.cn/customs/302249/zfxxgk/2799825/302274/302277/index.html"),
    ("GACC statistics EN", "http://english.customs.gov.cn/statics/report/monthly.html"),
    ("GACC EN home", "http://english.customs.gov.cn/"),
    ("GACC query", "http://stats.customs.gov.cn/"),
    ("GACC CN home", "https://www.customs.gov.cn/"),
    # NEA
    ("NEA home", "https://www.nea.gov.cn/"),
    ("NEA statistics list", "https://www.nea.gov.cn/sjzz/"),
    ("NEA stats news", "https://www.nea.gov.cn/sjzz/tjsj/index.htm"),
    ("NEA stats news 2", "https://www.nea.gov.cn/sjzz/index.htm"),
    ("NEA EN", "https://www.nea.gov.cn/english/"),
    # CEC
    ("CEC home", "https://www.cec.org.cn/"),
    ("CEC stats", "https://cec.org.cn/template/default/tjsj.html"),
    ("CEC EN", "https://english.cec.org.cn/"),
    # NDRC
    ("NDRC home", "https://www.ndrc.gov.cn/"),
    ("NDRC news", "https://www.ndrc.gov.cn/xxgk/jd/jd/"),
    ("NDRC fgsj", "https://www.ndrc.gov.cn/fggz/jjyxtj/"),
    # Gov portal / other
    ("gov.cn statistics", "https://www.gov.cn/lianbo/"),
    ("CNPC ETRI", "https://etri.cnpc.com.cn/"),
    ("China Coal Transport (CCTD)", "https://www.cctd.com.cn/"),
    ("CNOOC", "https://www.cnooc.com.cn/"),
    ("PBoC/SAFE n/a - IEA China", "https://www.iea.org/countries/china"),
    ("Ember China monthly csv", "https://ember-energy.org/data/china-electricity-data/"),
]
KEY = re.compile(r"天然气|进口|电力工业|装机|用电量|能源|原油|煤炭|液化|natural gas|import|capacity|electricity|energy|LNG|coal", re.I)


def fetch(label, url):
    t0 = time.time()
    try:
        r = requests.get(url, headers=H, timeout=(10, 30), allow_redirects=True)
    except requests.RequestException as e:
        print(f"[{label}] {url}\n   ERROR {type(e).__name__}: {str(e)[:150]}", flush=True)
        return
    r.encoding = r.apparent_encoding if r.encoding in (None, "ISO-8859-1") else r.encoding
    txt = r.text
    title = re.search(r"<title[^>]*>(.*?)</title>", txt, re.S | re.I)
    print(f"[{label}] {url}\n   status={r.status_code} bytes={len(r.content)} final={r.url} "
          f"{time.time() - t0:.1f}s title={title.group(1).strip()[:90] if title else None}", flush=True)
    if r.status_code == 200 and "Please enable JavaScript" in txt[:3000]:
        print("   anti-bot challenge page")
    links = re.findall(r'<a[^>]+href="([^"]+)"[^>]*>([^<]{4,80})</a>', txt)
    shown = 0
    for href, text in links:
        if KEY.search(text) and shown < 12:
            print(f"      {text.strip()[:70]} -> {href[:110]}")
            shown += 1


for label, url in URLS:
    fetch(label, url)
    time.sleep(1)
