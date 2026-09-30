"""
One-off discovery probe for EIA's Drilling Productivity Report (DPR) -
a direct public .xlsx download (no API key, unlike EIA-930/STEO's
api.eia.gov), published monthly, with rig counts AND oil/gas production
per major basin. Being a plain file download rather than Baker Hughes'
Akamai-gated site, this should actually work from GitHub Actions (Baker
Hughes' rig count script needs to run locally instead - see its
docstring).

Not yet confirmed: the exact sheet names/layout inside the workbook -
this probe downloads it and prints structure (sheet names, header rows,
a look at any sheet whose name suggests rig count or basin production)
before EIA_RIG_AND_BASIN_GAS_DAILY.py is written against it.
"""

import sys

import openpyxl
import requests

URL = "https://www.eia.gov/petroleum/drilling/xls/dpr-data.xlsx"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 60)


def main():
    print(f"Fetching {URL} ...", file=sys.stderr)
    r = requests.get(URL, headers=HEADERS, timeout=TIMEOUT)
    print(f"  status {r.status_code}, {len(r.content)} bytes, content-type "
          f"{r.headers.get('content-type')}", file=sys.stderr)
    r.raise_for_status()

    with open("/tmp/dpr-data.xlsx", "wb") as f:
        f.write(r.content)

    wb = openpyxl.load_workbook("/tmp/dpr-data.xlsx", data_only=True)
    print(f"\nSheet names ({len(wb.sheetnames)}): {wb.sheetnames}")

    for name in wb.sheetnames:
        ws = wb[name]
        print(f"\n=== Sheet '{name}': {ws.max_row} rows x {ws.max_column} cols ===")
        for row in ws.iter_rows(min_row=1, max_row=min(8, ws.max_row), max_col=min(10, ws.max_column)):
            print([cell.value for cell in row])


if __name__ == "__main__":
    main()
