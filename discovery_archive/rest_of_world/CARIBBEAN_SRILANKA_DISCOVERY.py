"""
First discovery pass for two brand-new countries (web search only so
far, nothing confirmed live):

  Trinidad & Tobago: data.gov.tt / finance.gov.tt publish a "Natural Gas
  Production and Utilisation" CSV (Ministry of Energy and Energy
  Industries) - annual, 2016-2022, by sector (power generation,
  petrochemicals, LNG, other). Confirming the real CSV downloads and
  its exact columns.

  Sri Lanka: PUCSL (Public Utilities Commission) publishes daily
  generation-mix PDF reports, but a web search also surfaced a
  "Dispatch Data Dashboard" at gendata.pucsl.gov.lk/home - if that's a
  real structured dashboard (not another PDF wall), it would be far
  easier to pull from than scraping PDFs. Fetching its HTML to look for
  an underlying API/JSON the page itself calls.

Usage: python3 CARIBBEAN_SRILANKA_DISCOVERY.py
"""
import io
import sys

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 60)


def try_get(label, url, **kwargs):
    print(f"\n{'=' * 70}\n{label}: {url}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kwargs)
        print(f"  status={r.status_code} bytes={len(r.content)} content-type={r.headers.get('content-type')}",
              file=sys.stderr)
        return r
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}", file=sys.stderr)
        return None


def trinidad_gas_csv():
    for url in [
        "https://www.finance.gov.tt/wp-content/uploads/2024/12/Appx-8-ROTE-2022-Natural-Gas.csv",
        "https://data.gov.tt/dataset/energy-sector-of-trinidad-and-tobago/resource/ccf96726-c247-436d-a6dc-de73c15f50c9",
    ]:
        r = try_get("Trinidad gas CSV candidate", url)
        if r is not None and r.status_code == 200 and "text" in (r.headers.get("content-type") or ""):
            import pandas as pd
            try:
                df = pd.read_csv(io.BytesIO(r.content))
                print(f"  parsed OK: shape={df.shape}", file=sys.stderr)
                print(f"  columns: {list(df.columns)}", file=sys.stderr)
                print(df.head(20).to_string(), file=sys.stderr)
            except Exception as e:
                print(f"  CSV parse failed: {type(e).__name__}: {e}", file=sys.stderr)
                print(r.text[:1000], file=sys.stderr)


def sri_lanka_pucsl_dashboard():
    r = try_get("PUCSL gendata dashboard", "https://gendata.pucsl.gov.lk/home")
    if r is not None and r.status_code == 200:
        text = r.text
        print(f"  page length: {len(text)} chars", file=sys.stderr)
        # look for embedded API calls / JSON / script src references
        import re
        for pattern, label in [
            (r'(https?://[^\s"\']+api[^\s"\']*)', "api-looking URLs"),
            (r'(https?://[^\s"\']+\.json[^\s"\']*)', "json URLs"),
            (r'src="([^"]+\.js)"', "JS bundles"),
            (r'(fetch\([^)]+\))', "fetch() calls in inline script"),
        ]:
            matches = sorted(set(re.findall(pattern, text)))[:15]
            if matches:
                print(f"  {label}: {matches}", file=sys.stderr)
        print(f"  first 1500 chars of page:\n{text[:1500]}", file=sys.stderr)


if __name__ == "__main__":
    print("=== Trinidad & Tobago gas CSV ===", file=sys.stderr)
    try:
        trinidad_gas_csv()
    except Exception as e:
        print(f"  FAILED: {type(e).__name__}: {e}", file=sys.stderr)

    print("\n=== Sri Lanka PUCSL dispatch dashboard ===", file=sys.stderr)
    try:
        sri_lanka_pucsl_dashboard()
    except Exception as e:
        print(f"  FAILED: {type(e).__name__}: {e}", file=sys.stderr)
