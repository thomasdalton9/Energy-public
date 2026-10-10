"""
NBS audit probe 5 (GitHub Actions only): ID scan of the migrated archive for 2021 releases (ids 1900961-1901260, titles of
CPI / profits / FAI / retail / energy production that the list does not reach) and the gaps in the 10-day price series
(2021 2月中旬, 2022 2月上旬, 2023 1月下旬).  usage: CHINA_NBS_AUDIT_PROBE5.py LO HI
Output: discovery_archive/results/china_nbs_audit/idscan_<LO>.tsv
"""
import os
import re
import sys

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
            f.write(f"{i}\tERROR {type(e).__name__}\t{url}\n")
            break
        if html is None:
            continue
        m = re.search(r"<title>(.*?)</title>", html, re.S)
        n += 1
        f.write(f"{i}\t{re.sub(chr(92) + 's+', '', m.group(1)) if m else '?'}\t{url}\n")
        f.flush()
print(n, "pages in range", lo, hi)
