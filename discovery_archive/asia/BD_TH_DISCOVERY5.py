"""
Check-up probe (runs via bd_th_discovery.yml, which takes the newest BD_TH_DISCOVERY*.py): India CEA daily coal stock
report layout. INDIA_NPP.py read 'actual stock' as 0 for most days before late Aug 2026, so print the header rows,
column labels and the grand-total row of dailyCoal1 files from 2024, 2025 and Sep 2026.
"""
import io
import os
import sys
from datetime import date

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "asia"))
import INDIA_NPP as N  # noqa: E402


def main():
    for d in (date(2024, 7, 15), date(2025, 6, 15), date(2026, 8, 20), date(2026, 9, 20)):
        c = N.fetch(f"{N.NPP}/fuel/{d:%d-%m-%Y}/dailyCoal1-{d:%Y-%m-%d}.xls")
        print(f"\n######## {d}: {'none' if c is None else len(c)} bytes", flush=True)
        if c is None:
            continue
        df = N.frame(c)
        with pd.option_context("display.width", 250, "display.max_columns", 40, "display.max_colwidth", 22):
            print(df.head(14).to_string())
            lab = df.apply(lambda r: " ".join(N.text(v) for v in r.iloc[:8]).upper(), axis=1)
            print("GRAND rows:\n" + df[lab.str.contains("GRAND TOTAL|ALL INDIA")].to_string())
        try:
            print("PARSED:", N.coal_day(d))
        except Exception as e:  # noqa: BLE001
            print("PARSE ERROR", repr(e))


if __name__ == "__main__":
    main()
