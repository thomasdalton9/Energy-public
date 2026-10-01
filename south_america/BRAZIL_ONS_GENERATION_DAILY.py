"""
Brazil daily power generation by fuel from ONS open data, in the standard
grid-operator layout shared by every country (scheduled in GitHub Actions;
"ONS Brazil.py" stays the owner's local hydro/thermal/wind/solar script).

Source: ONS "Geracao por usina em base horaria" (dataset geracao_usina_2_ho)
- every plant's hourly generation (MWmed, i.e. MWh in that hour) tagged
with its plant type and fuel. Yearly parquet files up to 2021, monthly
files from 2022:
  https://ons-aws-prod-opendata.s3.amazonaws.com/dataset/geracao_usina_2_ho/GERACAO_USINA-2_{YYYY}.parquet
  https://ons-aws-prod-opendata.s3.amazonaws.com/dataset/geracao_usina_2_ho/GERACAO_USINA-2_{YYYY}_{MM}.parquet

Each day's hourly national totals per ONS fuel label are summed to MWh per
day ('By ONS fuel' tab), then mapped onto the standard fuels ('Daily').

Incremental: the existing workbook is read back; only months with missing
days, plus the latest REFRESH_MONTHS months (ONS re-publishes recent
months), are downloaded. --full rebuilds from START.

Output (sheet 'Daily'): date, Hydro_MWh, Gas_MWh, Wind_MWh, Solar_MWh,
Coal_MWh, Nuclear_MWh, Oil_MWh, Bioenergy_MWh, Other_MWh, Total_MWh.

Usage: python3 south_america/BRAZIL_ONS_GENERATION_DAILY.py [--out PATH] [--full]
"""

import argparse
import io
import os
import sys
from datetime import date, timedelta

import pandas as pd
import requests

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)  # for xlsx_notes

import xlsx_notes  # noqa: E402

START = date(2021, 1, 1)
REFRESH_MONTHS = 2  # this month and last month are always re-downloaded
BASE = "https://ons-aws-prod-opendata.s3.amazonaws.com/dataset/geracao_usina_2_ho/"
MONTH_FILE = BASE + "GERACAO_USINA-2_{y}_{m:02d}.parquet"
YEAR_FILE = BASE + "GERACAO_USINA-2_{y}.parquet"
DEFAULT_OUT = os.path.join(REPO, "output", "Data and Chart Outputs", "brazil_power_generation_daily.xlsx")
READ_COLS = ["din_instante", "nom_tipousina", "nom_tipocombustivel", "val_geracao"]

FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]
# ONS fuel label (nom_tipocombustivel) -> standard fuel
FUEL_MAP = {
    "Hidráulica": "Hydro",
    "Gás": "Gas", "Multi-Combustível Gás/Diesel": "Gas",
    "Eólica": "Wind",
    "Fotovoltaica": "Solar",
    "Carvão": "Coal",
    "Nuclear": "Nuclear",
    "Óleo Combustível": "Oil", "Óleo Diesel": "Oil", "Multi-Combustível Diesel/Óleo": "Oil",
    "Biomassa": "Bioenergy", "Outras Multi-Combustível": "Bioenergy",
    "Resíduos Industriais": "Other",
}

NOTES = [
    "UNITS",
    "MWh per day (energy generated that day). ONS reports each plant's hourly output in MWmed (average MW over "
    "the hour = MWh); a day is the sum of its 24 hourly national totals. Total_MWh = sum of the fuel columns.",
    "",
    "TABS",
    "'Daily': standard layout shared by every country's grid-operator workbook. "
    "'By ONS fuel': the same days in MWh per ONS fuel label (nom_tipocombustivel), before mapping.",
    "",
    "CATEGORY MAPPING (ONS nom_tipocombustivel -> column)",
    "Hydro = Hidráulica. Gas = Gás (+ Multi-Combustível Gás/Diesel if it appears). Wind = Eólica. "
    "Solar = Fotovoltaica (utility-scale, hybrid wind-solar parks, and from ONS's MMGD series the estimated "
    "distributed micro/mini generation). Coal = Carvão. Nuclear = Nuclear (Angra 1 and 2). "
    "Oil = Óleo Combustível, Óleo Diesel, Multi-Combustível Diesel/Óleo. "
    "Bioenergy = Biomassa + Outras Multi-Combustível (ONS's 'Pequenas Usinas (Tipo III)' small-thermal "
    "aggregates per state - mostly sugar-cane bagasse cogeneration in SP/MG/GO/MS, strongly seasonal with the "
    "Apr-Nov harvest). Other = Resíduos Industriais (steel-mill gases at Ternium/Atlantico and Pecem, pulp-mill "
    "black liquor at Suzano) and any new ONS label not yet mapped.",
    "",
    "COVERAGE",
    "Plants in ONS's operating model (Tipo I, II-A/B/C, III, plant clusters) plus ONS's estimates for small "
    "plants. Distributed solar (MMGD, 'Pequenas Usinas (MMGD)') only enters the ONS series in May-2023, so "
    "Solar steps up then (about +2 TWh/month). Self-generation that never reaches the grid (most bagasse burnt "
    "for mills' own use, captive gas at industry/oil fields) is not included, so Bioenergy and Gas sit well "
    "below Ember/EPE. Hydro, Wind, Coal and Nuclear match Ember's monthly figures to within ~1% (Ember uses "
    "ONS for these); Ember's 'Other Fossil' = this workbook's Oil + Other.",
    "",
    "SOURCE",
    "ONS open data, dataset geracao_usina_2_ho (Geracao por usina em base horaria): "
    "https://ons-aws-prod-opendata.s3.amazonaws.com/dataset/geracao_usina_2_ho/GERACAO_USINA-2_YYYY_MM.parquet "
    "(monthly from 2022; GERACAO_USINA-2_YYYY.parquet for 2021). History from 2021-01-01. The latest day is "
    "dropped while ONS is still filling it in.",
]


def fetch(url):
    """Parquet at url as a DataFrame, or None when ONS hasn't published it."""
    r = requests.get(url, timeout=600)
    if r.status_code in (403, 404):
        return None
    r.raise_for_status()
    return pd.read_parquet(io.BytesIO(r.content), columns=READ_COLS)


def daily_by_label(raw):
    """Hourly plant rows -> MWh per day per ONS fuel label (complete days only)."""
    df = pd.DataFrame({
        "time": pd.to_datetime(raw["din_instante"]),
        "label": raw["nom_tipocombustivel"].astype(str).str.strip(),
        "mw": pd.to_numeric(raw["val_geracao"].astype(str), errors="coerce"),
    })
    hourly = df.pivot_table(index="time", columns="label", values="mw", aggfunc="sum")
    day = hourly.index.normalize()
    n = pd.Series(1, index=hourly.index).groupby(day).sum()
    step_h = pd.Series(hourly.index).diff().dt.total_seconds().div(3600).median()
    step_h = 1.0 if pd.isna(step_h) or step_h <= 0 else step_h
    per_day = round(24 / step_h)
    out = hourly.groupby(day).mean() * 24  # mean MW x 24 h = MWh (equals the hourly sum on full days)
    rows = df.groupby(df["time"].dt.normalize()).size()
    complete = n >= per_day
    # the newest day(s) can have every hour but not every plant yet - drop recent trailing days with clearly
    # fewer plant rows than usual (published months are final, so this only looks at the last two weeks)
    full = rows[complete].tail(30).median() if complete.any() else 0
    recent = pd.Timestamp(date.today() - timedelta(days=14))
    for d in reversed(list(out.index)):
        if d < recent or (complete.get(d, False) and rows.get(d, 0) >= 0.9 * full):
            break
        complete[d] = False
    dropped = [d for d in out.index if not complete.get(d, False)]
    if dropped:
        print(f"    dropped incomplete day(s): {', '.join(f'{d:%Y-%m-%d}' for d in dropped)}")
    out = out[complete.reindex(out.index).fillna(False).values]
    out.index = out.index.date
    return out


def months_to_fetch(have_dates, today, full):
    last = date(today.year, today.month, 1)
    months = pd.period_range(START, last, freq="M")
    if full or not have_dates:
        return list(months)
    have = set(have_dates)
    recent = set(pd.period_range(pd.Period(last, "M") - (REFRESH_MONTHS - 1), last, freq="M"))
    need = []
    for p in months:
        days = pd.date_range(p.start_time, min(p.end_time, pd.Timestamp(today - timedelta(days=1))), freq="D").date
        if p in recent or any(d not in have for d in days):
            need.append(p)
    return need


def load_existing(path):
    try:
        by_label = pd.read_excel(path, sheet_name="By ONS fuel", index_col=0)
    except (FileNotFoundError, ValueError) as e:
        print(f"No existing workbook to extend ({type(e).__name__}) - full history")
        return pd.DataFrame()
    by_label.index = pd.to_datetime(by_label.index).date
    return by_label


def to_standard(by_label):
    unknown = [c for c in by_label.columns if c not in FUEL_MAP]
    if unknown:
        print(f"WARNING: ONS fuel labels not mapped, counted as Other: {unknown}")
    out = pd.DataFrame(index=by_label.index)
    for f in FUELS:
        cols = [c for c in by_label.columns if FUEL_MAP.get(c, "Other") == f]
        out[f"{f}_MWh"] = by_label[cols].sum(axis=1, min_count=1) if cols else 0.0
    out = out.fillna(0.0)
    out["Total_MWh"] = out.sum(axis=1)
    out.index = pd.to_datetime(out.index)
    out.index.name = "date"
    return out.round(1)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=DEFAULT_OUT, help="workbook to write (and extend)")
    ap.add_argument("--full", action="store_true", help="ignore the existing workbook and rebuild from 2021")
    args = ap.parse_args()

    existing = pd.DataFrame() if args.full else load_existing(args.out)
    today = date.today()
    need = months_to_fetch(list(existing.index), today, args.full)
    print(f"Existing days: {len(existing):,}; months to fetch: {[str(p) for p in need]}", flush=True)

    fresh, years_done = [], {}
    for p in need:
        y, m = p.year, p.month
        if y in years_done:
            continue
        print(f"  {p}: ", end="", flush=True)
        raw = fetch(MONTH_FILE.format(y=y, m=m))
        if raw is None:
            print("no monthly file - trying the yearly file", flush=True)
            raw = fetch(YEAR_FILE.format(y=y))
            if raw is None:
                print(f"    {y}: not published (yet)")
                continue
            years_done[y] = True
        print(f"{len(raw):,} rows", flush=True)
        fresh.append(daily_by_label(raw))

    if fresh:
        new = pd.concat(fresh)
        new = new[~new.index.duplicated(keep="last")]
        new = new[[d >= START for d in new.index]]
        # fetched days replace what was there (a fuel missing from a fetched day is really absent)
        by_label = pd.concat([existing.drop(index=new.index, errors="ignore"), new]) if not existing.empty else new
    else:
        by_label = existing
    if by_label.empty:
        sys.exit("No ONS data")
    by_label = by_label.sort_index()
    by_label.index = pd.to_datetime(by_label.index)
    by_label.index.name = "date"
    by_label = by_label.round(1)

    daily = to_standard(by_label)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "By ONS fuel": by_label}, NOTES,
                              {"UNITS", "TABS", "CATEGORY MAPPING (ONS nom_tipocombustivel -> column)", "COVERAGE",
                               "SOURCE"})
    gaps = pd.date_range(daily.index.min(), daily.index.max()).difference(daily.index)
    print(f"\nSaved {args.out}: {len(daily):,} days, {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}; "
          f"{len(gaps)} missing day(s){': ' + ', '.join(f'{d:%Y-%m-%d}' for d in gaps[:20]) if len(gaps) else ''}")
    print("\nLatest days (MWh):")
    print(daily.tail(5).to_string())
    print("\nLatest months (GWh):")
    print((daily.resample("MS").sum() / 1000).round(0).tail(4).to_string())


if __name__ == "__main__":
    main()
