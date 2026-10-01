"""
Inspect the real ANEEL SIGA generation-enterprise registry CSV (found
via CAPACITY_DISCOVERY.py's CKAN package_search - confirmed real,
public, no key) before building a production script against it.
"""
import io
import sys

import pandas as pd
import requests

URL = ("https://dadosabertos.aneel.gov.br/dataset/6d90b77c-c5f5-4d81-bdec-7bc619494bb9/"
       "resource/11ec447d-698d-4ab8-977f-b424d5deee6a/download/siga-empreendimentos-geracao.csv")
HEADERS = {"User-Agent": "gas-demand-scripts/1.0"}


def main():
    r = requests.get(URL, headers=HEADERS, timeout=(10, 120))
    r.raise_for_status()
    print(f"status={r.status_code} bytes={len(r.content)}", file=sys.stderr)
    # ANEEL CSVs are typically ';'-delimited, latin-1, comma-decimal
    for sep, enc in [(";", "latin-1"), (",", "utf-8"), (";", "utf-8")]:
        try:
            df = pd.read_csv(io.BytesIO(r.content), sep=sep, encoding=enc, nrows=5000, low_memory=False)
            if df.shape[1] > 3:
                print(f"\nParsed OK with sep={sep!r} encoding={enc!r}: shape={df.shape}", file=sys.stderr)
                print(f"columns: {list(df.columns)}", file=sys.stderr)
                print(df.head(5).to_string(), file=sys.stderr)
                break
        except Exception as e:
            print(f"sep={sep!r} encoding={enc!r} failed: {type(e).__name__}: {e}", file=sys.stderr)


if __name__ == "__main__":
    main()
