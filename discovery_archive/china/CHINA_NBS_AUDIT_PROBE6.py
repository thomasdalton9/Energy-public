"""
NBS audit probe 6 (GitHub Actions only): find the five published-but-unlisted 10-day price releases (2021 2月中旬, 2022 2月上旬,
2023 1月下旬, 2024 2月中旬, 2025 1月下旬) - candidate URLs from the Wayback CDX index of the month folders around their
publication dates, then each unknown candidate is fetched from stats.gov.cn and its title printed.
Output: discovery_archive/results/china_nbs_audit/probe6_candidates.tsv
"""
import os
import re
import sys

import requests

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "asia"))
import china_nbs_common as nbs  # noqa: E402

RES = os.path.join(HERE, "..", "results", "china_nbs_audit")
known = {l.split("\t")[1].strip() for l in open(os.path.join(RES, "release_list.tsv"), encoding="utf-8") if "\t" in l}
DIRS = ["202102", "202103", "202202", "202203", "202302", "202303", "202402", "202403", "202501", "202502"]
out = open(os.path.join(RES, "probe6_candidates.tsv"), "w", encoding="utf-8")
for d in DIRS:
    url = ("http://web.archive.org/cdx/search/cdx?url=www.stats.gov.cn/sj/zxfb/" + d +
           "/&matchType=prefix&output=txt&fl=original&collapse=urlkey&limit=3000")
    try:
        r = requests.get(url, timeout=60)
        urls = sorted({u.split("?")[0].replace("http://", "https://").replace(":80", "") for u in r.text.split() if u.endswith(".html")})
        print(d, "CDX status", r.status_code, len(urls), "urls", flush=True)
    except Exception as e:  # noqa: BLE001
        print(d, "CDX failed", type(e).__name__, flush=True)
        continue
    new = [u for u in urls if u not in known and re.search(r"/t\d{8}_\d+\.html$", u)]
    print("  not on the release list:", len(new), flush=True)
    for u in new:
        if d == "202302" and "t20230203_" in u:
            continue   # migrated archive: covered by the ID scan
        try:
            html = nbs.fetch(u) or ""
        except Exception as e:  # noqa: BLE001
            out.write(f"{u}\tERROR {type(e).__name__}\n")
            continue
        m = re.search(r"<title>(.*?)</title>", html, re.S)
        t = re.sub(r"\s+", "", m.group(1)) if m else "?"
        out.write(f"{u}\t{t}\n")
        out.flush()
        if re.search(r"流通领域|重要生产资料", t):
            print("  FOUND", u, t, flush=True)
out.close()
