"""
Round 5: OilNOW (2024) quotes the Ministry: production data is on petroleum.gov.gy/data-visualization, "updated daily"
with a one-month lag, covering gas flared / used / reinjected. The /data-chart/ pages stop at Sep-2023 - is
/data-visualization a separate, current dashboard? Print: status, redirects, embedded data arrays with their date
ranges, iframes (Power BI / Tableau / Looker / ArcGIS / Google Sheets), script sources, ajax endpoints; then follow
each iframe and same-site data link one level down.
"""
import datetime as dt
import json
import re
from urllib.parse import urljoin

import requests

S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                                "Chrome/124.0 Safari/537.36"})
B = "https://petroleum.gov.gy"
seen = set()


def arrays(t, label):
    for m in re.finditer(r"(?:var|let|const)\s+(\w+)\s*=\s*'(\[.*?\])'\s*;", t, re.S):
        try:
            rows = json.loads(m.group(2))
        except Exception:   # noqa: BLE001
            continue
        ms = [x.get("Data") for x in rows if isinstance(x, dict) and isinstance(x.get("Data"), (int, float))]
        rng = (f"{dt.datetime.utcfromtimestamp(min(ms) / 1000):%Y-%m-%d} to "
               f"{dt.datetime.utcfromtimestamp(max(ms) / 1000):%Y-%m-%d}") if ms else "no Data field"
        keys = list(rows[-1]) if rows and isinstance(rows[-1], dict) else []
        print(f"    array {m.group(1)}: {len(rows)} rows, {rng}; keys {keys[:10]}", flush=True)


def look(url, depth=0):
    if url in seen or depth > 1:
        return
    seen.add(url)
    try:
        r = S.get(url, timeout=90, allow_redirects=True)
    except Exception as e:   # noqa: BLE001
        print(f"\n== {url}: {e}", flush=True)
        return
    t = r.text
    print(f"\n== {url} -> {r.url} [{r.status_code}] {len(t):,} chars, {r.headers.get('content-type')}", flush=True)
    title = re.search(r"<title>(.*?)</title>", t, re.S)
    desc = re.search(r'<meta name="description" content="([^"]+)"', t)
    print(f"    title {title.group(1).strip()[:120] if title else None!r}; description {desc.group(1)[:200] if desc else None!r}",
          flush=True)
    arrays(t, url)
    frames = re.findall(r"<iframe[^>]+src=[\"']([^\"']+)", t)
    print(f"    iframes: {frames}", flush=True)
    for kw in ["powerbi", "app.powerbi.com", "tableau", "lookerstudio", "datastudio", "arcgis", "docs.google.com",
               "admin-ajax", "wp-json", "fetch(", "$.ajax", "amcharts", "Highcharts", "chart.js", "wpdatatable",
               "tablepress", "2024", "2025", "2026", ".csv", ".xlsx"]:
        n = t.count(kw)
        if n:
            i = t.find(kw)
            snip = re.sub(r"\s+", " ", t[max(0, i - 120):i + 200])
            print(f"    '{kw}' x{n}: ...{snip!r}", flush=True)
    print("    script src:", sorted(set(re.findall(r"<script[^>]+src=[\"']([^\"']+)", t)))[:40], flush=True)
    links = sorted(set(urljoin(r.url, h) for h in re.findall(r"href=[\"']([^\"'#]+)", t)))
    data_links = [x for x in links if re.search(r"data|chart|visual|production|dashboard|statistic", x, re.I)
                  and "petroleum.gov.gy" in x]
    print(f"    data links ({len(data_links)}): {data_links[:60]}", flush=True)
    for f in frames:
        look(urljoin(r.url, f), depth + 1)
    for x in data_links:
        if re.search(r"visual|dashboard|production-data|statistic", x, re.I):
            look(x, depth + 1)


for path in ["/data-visualization", "/data-visualization/", "/data-visualisation/", "/data-centre/", "/production-data/",
             "/dashboard/"]:
    look(B + path)
