"""
Seventh NBS discovery pass (GitHub Actions only): print the product rows
of a 2021 industrial production release and the industry rows of 2021-22
capacity-utilisation releases, to check label changes (solar cells are
missing from the pulled output before 2023; three quarters lack the
industry-total capacity utilisation).
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "asia"))
import china_nbs_common as nbs  # noqa: E402

OLD = "https://www.stats.gov.cn/sj/zxfb/202302/t20230203_{}.html"
for i in (1901218, 1901025, 1901247, 1901053):
    url = OLD.format(i)
    html = nbs.fetch(url) or ""
    print(f"\n===== {url}")
    for cells in nbs.table_rows(html)[:120]:
        print("   | " + " | ".join(cells)[:160])
