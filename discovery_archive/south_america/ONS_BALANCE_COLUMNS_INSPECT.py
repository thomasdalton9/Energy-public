"""
"ONS Brazil.py" only extracts val_gerhidraulica/val_gertermica/
val_gereolica/val_gersolar (generation by source) from ONS's
BALANCO_ENERGIA_SUBSISTEMA parquet - but that dataset is literally named
"Energy BALANCE by Subsystem", so it likely also has load/demand and
inter-subsystem exchange columns nobody's looked at yet. Dumping the
full column list and a few sample rows from the current year's file
before building a real national+regional balance script.
"""
import io
import sys

import pandas as pd
import requests

URL = "https://ons-aws-prod-opendata.s3.amazonaws.com/dataset/balanco_energia_subsistema_ho/BALANCO_ENERGIA_SUBSISTEMA_{year}.parquet"
HEADERS = {"User-Agent": "gas-demand-scripts/1.0"}


def main():
    year = 2026
    r = requests.get(URL.format(year=year), headers=HEADERS, timeout=(10, 120))
    print(f"status={r.status_code} bytes={len(r.content)}", file=sys.stderr)
    r.raise_for_status()
    df = pd.read_parquet(io.BytesIO(r.content))
    print(f"shape={df.shape}", file=sys.stderr)
    print(f"columns: {list(df.columns)}", file=sys.stderr)
    print(df.head(10).to_string(), file=sys.stderr)
    print(f"\nunique id_subsistema: {sorted(df['id_subsistema'].unique()) if 'id_subsistema' in df.columns else 'N/A'}",
          file=sys.stderr)


if __name__ == "__main__":
    main()
