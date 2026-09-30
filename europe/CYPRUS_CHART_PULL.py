"""
One-off pull of the most recent 365 days of Cyprus daily generation mix,
reusing cyprus_generation_mix.py's own fetch/parse logic verbatim. This is
NOT the full historical pull (2016-10-26 to today - see
cyprus_generation_mix.py for that, which takes far longer to run); it's a
fast ~365-request sample covering one full seasonal cycle, for charting.

Prints the resulting CSV to stdout between marker lines so it can be
retrieved straight from the workflow's job log rather than needing an
artifact download.
"""

import sys
from datetime import date, timedelta

sys.path.insert(0, "europe")
import cyprus_generation_mix as cgm

FROM_DATE = date.today() - timedelta(days=365)
TO_DATE = date.today()


def main():
    df, capacity = cgm.build_daily_dataframe(FROM_DATE, TO_DATE)
    print(f"capacity: {capacity}", file=sys.stderr)
    print(f"{len(df)} days fetched", file=sys.stderr)
    print("===CSV_START===")
    print(df.to_csv())
    print("===CSV_END===")


if __name__ == "__main__":
    main()
