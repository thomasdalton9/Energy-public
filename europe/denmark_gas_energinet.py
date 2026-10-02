"""
Denmark daily gas flows from Energinet Gas TSO (Energi Data Service dataset "Gasflow" - commercial gas amounts,
no key). kWh per gas day -> GWh/d: sources positive (biogas, North Sea, Tyra), sinks negative (Danish demand,
Germany, Sweden, Poland), storage +/- (injection negative convention as published).
"""
import argparse
import os
import sys
from datetime import date

import pandas as pd

import europe_common as ec
import xlsx_notes

URL = "https://api.energidataservice.dk/dataset/Gasflow"
COLS = {"KWhFromBiogas": "Biogas", "KWhFromNorthSea": "North Sea", "kWhFromTyra": "Tyra",
        "KWhToOrFromStorage": "Storage (net, + = withdrawal)", "KWhToOrFromGermany": "Germany (net)",
        "KWhToSweden": "Sweden", "KWhToPoland": "Poland (Baltic Pipe)", "KWhToDenmark": "Danish demand"}


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=os.path.join(ec.DATA_DIR, "denmark_gas_flows.xlsx"))
    args = ap.parse_args()
    existing = ec.load_archive(args.out)
    start = date(2021, 1, 1) if existing.empty else max(date(2021, 1, 1), max(existing.index) - pd.Timedelta(days=21))
    j = ec.get_json(URL, {"start": f"{start.isoformat()}T00:00", "limit": 0, "sort": "GasDay asc"})
    d = pd.DataFrame(j.get("records", []))
    if d.empty:
        raise RuntimeError("Energinet Gasflow returned no rows")
    print("columns:", list(d.columns), file=sys.stderr)
    d["date"] = pd.to_datetime(d["GasDay"]).dt.date
    new = d.set_index("date")[[c for c in COLS if c in d]].apply(pd.to_numeric, errors="coerce") / 1e6
    new = new.rename(columns=COLS).round(3)
    new.index.name = "date"
    combined = new if existing.empty else pd.concat([existing, new])
    combined = combined[~combined.index.duplicated(keep="last")].sort_index()
    notes = ["UNITS", "GWh per gas day (Energinet publishes kWh). Sources positive; exits negative.", "", "COLUMNS",
             ", ".join(combined.columns), "", "COVERAGE", "Daily from 2021-01-01; last 21 days re-pulled each run.", "",
             "SOURCE", "Energinet Gas TSO, Energi Data Service 'Gasflow': https://www.energidataservice.dk/tso-gas/Gasflow"]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Daily": combined}, notes, ec.NOTES_TITLES)
    print(f"Saved {args.out} ({len(combined)} days, {combined.index.min()} to {combined.index.max()})")


if __name__ == "__main__":
    main()
