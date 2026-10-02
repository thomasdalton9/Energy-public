"""
Round 2: petroleum.gov.gy/data-chart/<chart>/ pages (gas-produced, gas-injected-flared-and-used, oil-production)
are ~0.5 MB each - the series look embedded in the page. Find how: chart library config (Chart.js / Highcharts /
ApexCharts / Google Charts / Visualizer plugin), wpDataTables / TablePress tables, CSV / JSON links, admin-ajax.
Prints the structure and the first and last data points of every series found.
"""
import json
import re

import requests

S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                                "Chrome/124.0 Safari/537.36"})
for slug in ["gas-produced", "gas-injected-flared-and-used", "oil-production"]:
    url = f"https://petroleum.gov.gy/data-chart/{slug}/"
    t = S.get(url, timeout=90).text
    print(f"\n===== {slug}: {len(t):,} chars", flush=True)
    for kw in ["visualizer", "wpdatatable", "tablepress", "highcharts", "Highcharts", "Chart(", "new Chart", "apexcharts",
               "google.visualization", "plotly", "echarts", "amcharts", "admin-ajax", ".csv", ".json", "datasets",
               "series", "MMSCF", "mmscf", "Reinject", "reinject", "Flared", "flared", "Fuel"]:
        n = t.count(kw)
        if n:
            i = t.find(kw)
            print(f"  '{kw}' x{n}: ...{t[max(0, i - 150):i + 250]!r}", flush=True)
    # the largest script blocks
    scripts = re.findall(r"<script[^>]*>(.*?)</script>", t, re.S)
    big = sorted(scripts, key=len, reverse=True)[:3]
    for s in big:
        print(f"  big script {len(s):,} chars: head {s[:400]!r}", flush=True)
        print(f"     tail {s[-300:]!r}", flush=True)
    # JSON-ish arrays of numbers / dates
    for m in list(re.finditer(r"\[(\s*\"?\d{4}-\d{2}-\d{2}[^\]]{0,200})", t))[:3]:
        print(f"  date array: {m.group(0)[:200]!r}", flush=True)
    tables = re.findall(r"<table.*?</table>", t, re.S)
    print(f"  {len(tables)} html tables; first rows: "
          f"{[re.sub(r'<[^>]+>', '|', r)[:200] for r in re.findall(r'<tr.*?</tr>', tables[0], re.S)[:4]] if tables else None}",
          flush=True)
