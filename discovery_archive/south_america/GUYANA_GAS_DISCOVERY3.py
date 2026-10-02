"""
Round 3: the data-chart pages' embedded `var _data = '[...]'` arrays stop at 30-Sep-2023 although the page header says
the period runs to Jul-2026. Find where the later days come from: every _data-like variable (first/last date, count),
every <script src>, inline fetch/ajax/admin-ajax/wp-json calls, date-range inputs or year selectors, nonces, and any
links to CSV/XLSX downloads or other data-chart pages.
"""
import json
import re

import requests

S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                                "Chrome/124.0 Safari/537.36"})


def show(label, t, kw, width=300, limit=6):
    for i in [m.start() for m in re.finditer(re.escape(kw), t)][:limit]:
        print(f"  {label} '{kw}' @{i}: {t[max(0, i - 150):i + width]!r}", flush=True)


for slug in ["gas-produced", "gas-injected-flared-and-used"]:
    url = f"https://petroleum.gov.gy/data-chart/{slug}/"
    r = S.get(url, timeout=90)
    t = r.text
    print(f"\n===== {slug}: {len(t):,} chars, status {r.status_code}", flush=True)
    for m in re.finditer(r"(?:var|let|const)\s+(\w+)\s*=\s*'(\[.*?\])'\s*;", t, re.S):
        try:
            rows = json.loads(m.group(2))
        except Exception as e:   # noqa: BLE001
            print(f"  var {m.group(1)}: not JSON ({e})", flush=True)
            continue
        keys = list(rows[0].keys()) if rows else []
        dates = [x.get("Date") for x in rows if isinstance(x, dict)]
        print(f"  var {m.group(1)} @{m.start()}: {len(rows)} rows, keys {keys}, first {dates[:1]}, last {dates[-1:]}",
              flush=True)
    for m in re.finditer(r"(?:var|let|const)\s+(\w+)\s*=\s*([\[{\"'])", t):
        print(f"  var decl {m.group(1)} ({m.group(2)}) @{m.start()}", flush=True)
    print("  script src:", sorted(set(re.findall(r"<script[^>]+src=[\"']([^\"']+)", t))), flush=True)
    for kw in ["admin-ajax", "wp-json", "fetch(", "$.ajax", "$.post", "$.get", "XMLHttpRequest", "nonce", "action:",
               "action=", "date_from", "dateFrom", "start_date", "end_date", "datepicker", "daterange", "<select",
               "<input", "_data", "2023", "Sep-2023", "Jul-2026", "July 2026", ".csv", ".xlsx", "download"]:
        show(slug, t, kw, limit=4)
    links = sorted(set(re.findall(r"href=[\"'](https?://petroleum\.gov\.gy/[^\"'#]+)", t)))
    print("  data links:", [x for x in links if re.search(r"data|chart|report|download|\.csv|\.xls", x, re.I)], flush=True)
    # inline scripts that reference _data: print them whole (trimmed) so the chart/filter logic is visible
    for s in re.findall(r"<script[^>]*>(.*?)</script>", t, re.S):
        if "_data" in s and len(s) < 400_000:
            body = re.sub(r"'\[.*?\]'", "'[...]'", s, flags=re.S)
            print(f"  --- script using _data ({len(s):,} chars, arrays elided):\n{body[:4000]}", flush=True)

# the listing page of all data charts, and the site's REST API for the data-chart post type
for url in ["https://petroleum.gov.gy/data-centre/", "https://petroleum.gov.gy/data-chart/",
            "https://petroleum.gov.gy/wp-json/wp/v2/types", "https://petroleum.gov.gy/wp-json/"]:
    try:
        r = S.get(url, timeout=60)
        print(f"\n===== {url}: {r.status_code}, {len(r.text):,} chars", flush=True)
        if "wp-json" in url:
            print(r.text[:3000], flush=True)
        else:
            print("  links:", sorted(set(re.findall(r"href=[\"'](https?://petroleum\.gov\.gy/data[^\"'#]*)", r.text)))[:80],
                  flush=True)
    except Exception as e:   # noqa: BLE001
        print(f"{url}: {e}", flush=True)
