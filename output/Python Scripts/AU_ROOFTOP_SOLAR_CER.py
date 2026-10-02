"""
Australia small-scale (rooftop) solar and home batteries from the Clean Energy Regulator's postcode data
(public, monthly by postcode; rolled up to states here):

  https://cer.gov.au/document/sgu-solar-installations-2011-to-present-and-totals         installs (CSV)
  https://cer.gov.au/document/sres-postcode-data-capacity-2011-to-present-and-totals     capacity, kW (xlsx)
  https://cer.gov.au/document/sgu-battery-installations-2011-to-present-and-totals       battery installs (CSV)

Writes au_rooftop_solar.xlsx: "Solar installs" (systems per month by state), "Solar capacity" (MW added per month
by state) and "Battery installs" (per month by state, from Jul 2025 in this file).

Each file holds the full history, so a run downloads it only when its Last-Modified date has changed (recorded on
the Units sheet). The latest ~12 months keep rising for a year (installers have 12 months to create
certificates), so recent months are understated until then.

Usage: python3 AU_ROOFTOP_SOLAR_CER.py [--out "output/Data and Chart Outputs/au_rooftop_solar.xlsx"]
"""
import argparse
import io
import os
import re
import sys

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

CER = "https://cer.gov.au/document/"
FILES = {"Solar installs": "sgu-solar-installations-2011-to-present-and-totals",
         "Solar capacity": "sres-postcode-data-capacity-2011-to-present-and-totals",
         "Battery installs": "sgu-battery-installations-2011-to-present-and-totals"}
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "au_rooftop_solar.xlsx")
STATES = ["NSW", "VIC", "QLD", "SA", "WA", "TAS", "ACT", "NT"]


def state_of(pc):
    """Australian postcode -> state/territory."""
    try:
        p = int(float(pc))
    except (TypeError, ValueError):
        return None
    if 200 <= p <= 299 or 2600 <= p <= 2618 or 2900 <= p <= 2920:
        return "ACT"
    if 800 <= p <= 999:
        return "NT"
    if 1000 <= p <= 2999:
        return "NSW"
    if 3000 <= p <= 3999 or 8000 <= p <= 8999:
        return "VIC"
    if 4000 <= p <= 4999 or 9000 <= p <= 9999:
        return "QLD"
    if 5000 <= p <= 5999:
        return "SA"
    if 6000 <= p <= 6999:
        return "WA"
    if 7000 <= p <= 7999:
        return "TAS"
    return None


def by_state(d, pc_col, scale=1.0):
    """Wide postcode x 'Mon YYYY - ...' table -> month x state."""
    months = {c: pd.to_datetime(re.match(r"([A-Z][a-z]{2} \d{4})", str(c)).group(1), format="%b %Y")
              for c in d.columns if re.match(r"[A-Z][a-z]{2} \d{4}", str(c))}
    d = d.assign(state=d[pc_col].map(state_of)).dropna(subset=["state"])
    v = d[list(months)].apply(pd.to_numeric, errors="coerce").groupby(d["state"]).sum(min_count=1).T
    v.index = [months[c] for c in v.index]
    v = v.reindex(columns=[s for s in STATES if s in v.columns]) * scale
    v["Total"] = v.sum(axis=1, min_count=1)
    v.index.name = "Month"
    return v.sort_index()


def read(name, content):
    if content[:2] == b"PK":   # capacity workbook: sheet SGU-Solar, title rows above the header
        raw = pd.read_excel(io.BytesIO(content), sheet_name="SGU-Solar", header=None)
        h = next(i for i in range(len(raw)) if raw.iloc[i].astype(str).str.contains("Postcode", case=False).any())
        d = pd.read_excel(io.BytesIO(content), sheet_name="SGU-Solar", header=h)
        return by_state(d, next(c for c in d.columns if "postcode" in str(c).lower()), scale=1 / 1000.0)   # kW -> MW
    d = pd.read_csv(io.BytesIO(content), low_memory=False)
    return by_state(d, next(c for c in d.columns if "postcode" in str(c).lower()))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    try:
        units = " ".join(str(x) for x in pd.read_excel(args.out, sheet_name=0, header=None).iloc[:, 0].dropna())
    except (FileNotFoundError, ValueError, KeyError, OSError):
        units = ""
    sheets, stamps = {}, {}
    for name, doc in FILES.items():
        url = CER + doc
        head = requests.head(url, headers=HEADERS, timeout=(10, 60), allow_redirects=True)
        stamp = head.headers.get("Last-Modified", "")
        stamps[name] = stamp
        if stamp and f"{name} file Last-Modified: {stamp}" in units:
            try:   # unchanged: keep the saved sheet, nothing downloaded
                old = pd.read_excel(args.out, sheet_name=name, index_col=0)
                old.index = pd.to_datetime(old.index)
                sheets[name] = old
                print(f"{name}: unchanged since {stamp}", flush=True)
                continue
            except (ValueError, KeyError):
                pass
        r = requests.get(url, headers=HEADERS, timeout=(10, 180))
        r.raise_for_status()
        stamps[name] = r.headers.get("Last-Modified", stamp)
        sheets[name] = read(name, r.content)
        t = sheets[name]
        print(f"{name}: {len(t)} months {t.index.min():%Y-%m}..{t.index.max():%Y-%m}; last: "
              f"{t.iloc[-1].round(1).to_dict()}", flush=True)
    s = sheets.get("Solar installs", pd.DataFrame())
    notes = [
        "UNITS",
        "Solar installs: small-scale solar systems (under 100 kW, mostly rooftop) installed per month, by state "
        "(postcodes mapped to states). Solar capacity: MW of those systems installed per month (CER 'deemed' "
        "capacity, kW / 1,000). Battery installs: home batteries installed with small-scale systems per month.",
        "Recent months are understated: installers have 12 months to create certificates, so the latest year keeps "
        "rising in later releases.",
        "",
        "COVERAGE",
        (f"Australia by state, monthly, {s.index.min():%b %Y} to {s.index.max():%b %Y}. " if not s.empty else "") +
        "Battery installs only from Jul 2025 in this file.",
        "",
        "SOURCE",
        "Clean Energy Regulator, small-scale installation postcode data: " + ", ".join(CER + d for d in FILES.values()),
        "https://cer.gov.au/markets/reports-and-data/small-scale-installation-postcode-data",
        *[f"{k} file Last-Modified: {v or 'not given'}" for k, v in stamps.items()],
    ]
    out = {k: v.round(2) for k, v in sheets.items()}
    for v in out.values():
        v.index.name = "Month"
    xlsx_notes.write_workbook(args.out, out, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
