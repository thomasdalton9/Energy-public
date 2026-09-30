"""
Discovery pass for a token-free route to AEMO NEM generation-by-fuel-
type history back to 2020, daily, per state - an alternative to
australia_power_mix.py's OpenElectricity API (which needs a paid-
signup-free-but-still-a-registered token the user hasn't set up).

AEMO's own NEMWEB (nemweb.com.au) is a plain public directory listing,
no auth - the underlying raw data OpenElectricity/OpenNEM themselves
are built on. Two tables are what's needed:
  - DUDETAILSUMMARY: maps each generating unit (DUID) to its NEM region
    and fuel/technology type, with effective-date ranges (units change
    fuel/owner rarely, but it happens).
  - DISPATCH_UNIT_SCADA: actual 5-minute dispatched MW per DUID - the
    real generation numbers. This is the big one (every unit, every
    5 minutes) - checking here whether pulling ~6 years of it is even
    practically sized/reachable before committing to building a full
    parser.

A public PyPI package, NEMOSIS (github.com/UNSW-CEEM/NEMOSIS), already
implements exactly this NEMWEB download/cache/parse logic for both
tables (handles the "Current" vs "Archive" vs monthly "MMSDM_Historical"
zip layouts transparently) - checking whether it installs and works
cleanly here rather than reimplementing NEMWEB's zip layout by hand.

This script just probes reachability/feasibility and reports what's
actually there - no assumptions carried in from outside this run.
"""

import sys

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 30)


def try_get(label, url, **kwargs):
    print(f"\n{'=' * 70}\n{label}: {url}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kwargs)
        print(f"  status={r.status_code}  bytes={len(r.content)}  content-type={r.headers.get('content-type')}")
        return r
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}")
        return None


def try_head(label, url):
    print(f"\n{'=' * 70}\n{label} (HEAD): {url}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.head(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        length = r.headers.get("content-length")
        mb = f"{int(length) / 1e6:.1f} MB" if length else "unknown size"
        print(f"  status={r.status_code}  size={mb}")
        return r
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}")
        return None


def probe_nemweb_reachability():
    for url in [
        "https://nemweb.com.au/",
        "https://www.nemweb.com.au/Reports/Current/",
        "https://www.nemweb.com.au/Reports/Archive/",
    ]:
        try_get("nemweb.com.au", url)


def probe_mmsdm_historical_archive():
    # MMSDM_HistoricalDataSets - AEMO's monthly zipped full-database
    # snapshots, the source NEMOSIS itself pulls from for historical
    # (non-"current") months. Checking the directory listing for 2020
    # and a recent month, and the size of one actual DISPATCH_UNIT_SCADA
    # monthly zip via HEAD (without downloading it) to gauge whether a
    # 6-year x 12-month backfill is remotely practical in one CI run.
    for url in [
        "https://www.nemweb.com.au/Reports/Archive/MMSDM_HistoricalDataSets/",
        "https://www.nemweb.com.au/Reports/Archive/MMSDM_HistoricalDataSets/2020/",
        "https://www.nemweb.com.au/Reports/Archive/MMSDM_HistoricalDataSets/2020/MMSDM_2020_01/",
    ]:
        r = try_get("MMSDM historical archive listing", url)
        if r is not None and r.status_code == 200:
            print(r.text[:2500])

    # guessed exact monthly DISPATCH_UNIT_SCADA zip filename pattern
    # (AEMO's own documented naming convention for MMSDM data model
    # report files) - HEAD-only, to see the size without downloading
    for url in [
        "https://www.nemweb.com.au/Reports/Archive/MMSDM_HistoricalDataSets/2020/MMSDM_2020_01/MMSDM_Historical_Data_SQLLoader/DATA/PUBLIC_DVD_DISPATCH_UNIT_SCADA_202001010000.zip",
    ]:
        try_head("guessed DISPATCH_UNIT_SCADA monthly zip", url)


def probe_nemosis_install():
    import subprocess

    result = subprocess.run(
        [sys.executable, "-m", "pip", "install", "--quiet", "nemosis"],
        capture_output=True, text=True, timeout=120,
    )
    print(f"pip install nemosis: returncode={result.returncode}")
    if result.returncode != 0:
        print(f"  stdout: {result.stdout[-2000:]}")
        print(f"  stderr: {result.stderr[-2000:]}")
        return

    import tempfile
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            from nemosis import static_table_xl, dynamic_data_compiler

            print("Trying static_table (DUDETAILSUMMARY - unit-to-region/fuel-type mapping)...")
            df = static_table_xl.static_table_xl_cache(
                "2024/01/01 00:00:00", "2024/01/02 00:00:00",
                "Generators and Scheduled Loads", tmpdir,
            )
            print(f"  columns: {list(df.columns)}")
            print(df.head(10).to_string())
        except Exception as e:
            print(f"  static table probe FAILED: {type(e).__name__}: {e}")

        try:
            from nemosis import dynamic_data_compiler

            print("\nTrying dynamic_data_compiler for DISPATCH_UNIT_SCADA, one hour, one day...")
            df = dynamic_data_compiler(
                "2024/06/01 00:00:00", "2024/06/01 01:00:00",
                "DISPATCH_UNIT_SCADA", tmpdir,
            )
            print(f"  rows={len(df)}  columns={list(df.columns)}")
            print(df.head(10).to_string())
        except Exception as e:
            print(f"  dynamic_data_compiler probe FAILED: {type(e).__name__}: {e}")


def main():
    probe_nemweb_reachability()
    probe_mmsdm_historical_archive()
    probe_nemosis_install()


if __name__ == "__main__":
    main()
