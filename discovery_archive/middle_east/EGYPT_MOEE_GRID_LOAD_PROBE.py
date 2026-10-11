"""
Probe (Oct 2026): is the Egyptian Ministry of Electricity (MOEE) home-page 'Grid Loads' panel current?
Prints every span whose id starts ctl00_content_h_ (date, peak, min, temperature, expected peak) on the
English and Arabic home pages, plus the HTTP Last-Modified/Date headers.

Workflow (manual): discovery_archive/workflows/egypt_moee_grid_load_probe.yml.
"""
import re

import requests

for url in ["http://www.moee.gov.eg/english_new/home.aspx", "http://www.moee.gov.eg/arabic/home.aspx"]:
    try:
        r = requests.get(url, timeout=60, headers={"User-Agent": "Mozilla/5.0"})
    except Exception as e:  # noqa: BLE001
        print(url, "failed:", e)
        continue
    print(f"\n=== {url}: HTTP {r.status_code}, Date {r.headers.get('Date')}, Last-Modified {r.headers.get('Last-Modified')}")
    r.encoding = r.apparent_encoding
    for sid, txt in re.findall(r'<span[^>]*id="(ctl00_content_h_[^"]*)"[^>]*>(.*?)</span>', r.text, re.S):
        print(f"  {sid}: {' '.join(re.sub('<[^>]+>', ' ', txt).split())!r}")
    i = r.text.find("h_max_load")
    print("  context:", " ".join(re.sub("<[^>]+>", " ", r.text[max(0, i - 2500):i + 1500]).split())[:1500])
