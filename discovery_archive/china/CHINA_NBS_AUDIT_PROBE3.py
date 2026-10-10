"""
NBS audit probe 3 (GitHub Actions only): sequential ID scan of the migrated archive
(/sj/zxfb/202302/t20230203_<id>.html) to find releases older than the release list reaches (Jan 2021 and earlier).
usage: CHINA_NBS_AUDIT_PROBE3.py LO HI  -> discovery_archive/results/china_nbs_audit/idscan_<LO>.tsv (id, title, url) for every
page that answers 200 with a title (all titles kept; filtered later).
"""
import os
import re
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "asia"))
import china_nbs_common as nbs  # noqa: E402

lo, hi = int(sys.argv[1]), int(sys.argv[2])
RES = os.path.join(HERE, "..", "results", "china_nbs_audit")
nbs.PAUSE_SECONDS = 0.5
n = 0
with open(os.path.join(RES, f"idscan_{lo}.tsv"), "w", encoding="utf-8") as f:
    for i in range(hi, lo - 1, -1):
        url = f"https://www.stats.gov.cn/sj/zxfb/202302/t20230203_{i}.html"
        try:
            html = nbs.fetch(url, attempts=3)
        except Exception as e:  # noqa: BLE001
            print(i, "giving up", type(e).__name__, flush=True)
            f.write(f"{i}\tERROR {type(e).__name__}\t{url}\n")
            break
        if html is None:
            continue
        m = re.search(r"<title>(.*?)</title>", html, re.S)
        t = re.sub(r"\s+", "", m.group(1)) if m else "?"
        n += 1
        f.write(f"{i}\t{t}\t{url}\n")
        f.flush()
print(n, "pages in range", lo, hi)
