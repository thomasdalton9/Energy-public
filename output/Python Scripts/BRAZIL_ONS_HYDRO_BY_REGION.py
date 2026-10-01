print("STARTING", flush=True)

import argparse
import datetime as dt
import io
import sys
from pathlib import Path

import pandas as pd
import requests

START_YEAR = 2000
END_YEAR = dt.date.today().year

# Default (no arguments): the owner's desktop pipeline - writes the four
# ear_*.csv files to the current directory, as before.
# --out <xlsx>: scheduled pull (.github/workflows/brazil_hydro_reservoirs.yml)
# - writes one workbook (sheet "Daily": storage % and stored energy per
# subsystem plus the SIN national total) instead of the CSVs. Incremental:
# the workbook is its own archive, so only years ONS may still revise
# (this year, and last year until April) are downloaded again.
ap = argparse.ArgumentParser(description="ONS EAR reservoir storage by subsystem")
ap.add_argument("--out", help="write a workbook here instead of the ear_*.csv files")
ARGS = ap.parse_args()
OUT_KEEP = None   # --out: rows already in the workbook that are kept as they are

BASE = (
    "https://ons-aws-prod-opendata.s3.amazonaws.com/dataset/"
    "ear_subsistema_di/EAR_DIARIO_SUBSISTEMA_"
)

# Cache: each year's file is kept in ons_cache/ (shared with "ONS
# Brazil.py"). Past years are downloaded once; this year's - and last
# year's until April, while ONS may still revise it - every run.
CACHE_DIR = Path(__file__).resolve().parent / "ons_cache"


def settled(year):
    today = dt.date.today()
    return year < today.year - 1 or (year == today.year - 1 and today.month > 3)


def parse(content, ext):
    if ext == "csv":
        return pd.read_csv(io.BytesIO(content), sep=";", encoding="utf-8")
    return pd.read_excel(io.BytesIO(content))


def get_year(year):
    """
    Cached copy if the year is settled; otherwise download -
    CSV first, falling back to XLSX - and cache it.
    """

    CACHE_DIR.mkdir(exist_ok=True)
    for ext in ("csv", "xlsx"):
        cached = CACHE_DIR / f"EAR_DIARIO_SUBSISTEMA_{year}.{ext}"
        if cached.exists() and settled(year):
            df = parse(cached.read_bytes(), ext)
            print(f"  {year}: {len(df):,} rows (cached)", flush=True)
            return df

    for ext in ("csv", "xlsx"):
        url = f"{BASE}{year}.{ext}"
        cached = CACHE_DIR / f"EAR_DIARIO_SUBSISTEMA_{year}.{ext}"

        try:
            r = requests.get(url, timeout=60)
            r.raise_for_status()
            df = parse(r.content, ext)
            cached.write_bytes(r.content)

            print(
                f"  {year}: {len(df):,} rows ({ext})",
                flush=True
            )

            return df

        except Exception as e:
            if cached.exists():  # offline or ONS hiccup - use the last good copy
                df = parse(cached.read_bytes(), ext)
                print(f"  {year}: download failed ({type(e).__name__}) - using cached copy", flush=True)
                return df
            if ext == "xlsx":
                print(
                    f"  {year}: FAILED ({type(e).__name__})",
                    flush=True
                )

    return None


if ARGS.out:
    try:
        OUT_KEEP = pd.read_excel(ARGS.out, sheet_name="Daily", index_col=0)
        OUT_KEEP.index = pd.to_datetime(OUT_KEEP.index, errors="coerce")
        OUT_KEEP = OUT_KEEP[OUT_KEEP.index.notna()].sort_index()
        first_unsettled = next(y for y in range(START_YEAR, END_YEAR + 1) if not settled(y))
        if len(OUT_KEEP):
            START_YEAR = max(START_YEAR, min(first_unsettled, OUT_KEEP.index.max().year))
            OUT_KEEP = OUT_KEEP[OUT_KEEP.index < pd.Timestamp(START_YEAR, 1, 1)]
            print(f"{ARGS.out}: keeping {len(OUT_KEEP):,} days before {START_YEAR}; "
                  f"downloading {START_YEAR}-{END_YEAR}", flush=True)
    except (FileNotFoundError, ValueError, KeyError) as e:
        print(f"{ARGS.out}: no usable archive ({type(e).__name__}) - full history from {START_YEAR}", flush=True)
        OUT_KEEP = None


# ----------------------------------------------------------
# DOWNLOAD ALL YEARS
# ----------------------------------------------------------

frames = []

for y in range(START_YEAR, END_YEAR + 1):

    df = get_year(y)

    if df is not None and len(df):
        frames.append(df)

if not frames:
    raise RuntimeError("No data downloaded")

raw = pd.concat(
    frames,
    ignore_index=True
)

print(
    f"\ncolumns: {list(raw.columns)}\n",
    flush=True
)


# ----------------------------------------------------------
# COLUMN DETECTION
# ----------------------------------------------------------

def pick(*keywords, exclude=()):

    for c in raw.columns:

        lc = c.lower()

        if (
            all(k in lc for k in keywords)
            and not any(x in lc for x in exclude)
        ):
            return c

    raise KeyError(
        f"Could not find column: {keywords}"
    )


date_col = pick("data")

sub_col = pick(
    "subsistema",
    exclude=("id",)
)

energy_col = pick(
    "verif",
    exclude=("percent",)
)

pct_col = pick("percent")

print(
    f"using -> "
    f"{date_col} | "
    f"{sub_col} | "
    f"{energy_col} | "
    f"{pct_col}",
    flush=True
)


# ----------------------------------------------------------
# CLEAN DATAFRAME
# ----------------------------------------------------------

df = raw[
    [
        date_col,
        sub_col,
        energy_col,
        pct_col,
    ]
].copy()

df.columns = [
    "Date",
    "Subsystem",
    "StoredEnergy_MWmes",
    "StoragePct",
]

df["Date"] = pd.to_datetime(
    df["Date"],
    errors="coerce"
)

df["StoredEnergy_MWmes"] = pd.to_numeric(
    df["StoredEnergy_MWmes"],
    errors="coerce"
)

df["StoragePct"] = pd.to_numeric(
    df["StoragePct"],
    errors="coerce"
)

df = df.dropna(
    subset=[
        "Date",
        "StoredEnergy_MWmes",
        "StoragePct",
    ]
)

print(
    f"\nloaded {len(df):,} rows\n",
    flush=True
)


# ----------------------------------------------------------
# DAILY STORAGE %
# ----------------------------------------------------------

daily_pct = (
    df.pivot_table(
        index="Date",
        columns="Subsystem",
        values="StoragePct",
        aggfunc="last",
    )
    .sort_index()
    .reset_index()
)

# ----------------------------------------------------------
# DAILY STORED ENERGY
# ----------------------------------------------------------

daily_energy = (
    df.pivot_table(
        index="Date",
        columns="Subsystem",
        values="StoredEnergy_MWmes",
        aggfunc="last",
    )
    .sort_index()
    .reset_index()
)

# ----------------------------------------------------------
# --out: SCHEDULED WORKBOOK (instead of the CSVs below)
# ----------------------------------------------------------

# ONS subsystem names -> short codes, in the workbook's column order.
SUBSYSTEM_CODES = {
    "SUDESTE": "SE_CO", "SUDESTE/CENTRO-OESTE": "SE_CO", "SE": "SE_CO", "SE/CO": "SE_CO",
    "SUL": "S", "S": "S",
    "NORDESTE": "NE", "NE": "NE",
    "NORTE": "N", "N": "N",
}
CODES = ["SE_CO", "S", "NE", "N"]


def write_out_workbook(path):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    import xlsx_notes

    # Maximum storable energy per subsystem (ear_max_subsistema, MWmes)
    # - published in the same files. If ONS ever drops it, fall back to
    # the capacity implied by stored energy / storage %.
    try:
        max_col = pick("max", exclude=("percent",))
    except KeyError:
        max_col = None
    print(f"capacity column: {max_col or '(none - implied from energy / %)'}", flush=True)

    d = raw[[date_col, sub_col, energy_col, pct_col] + ([max_col] if max_col else [])].copy()
    d.columns = ["Date", "Subsystem", "Energy", "Pct"] + (["Max"] if max_col else [])
    d["Date"] = pd.to_datetime(d["Date"], errors="coerce")
    for c in d.columns[2:]:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    if not max_col:
        d["Max"] = d["Energy"] / (d["Pct"].where(d["Pct"] > 0) / 100)
    names = d["Subsystem"].astype(str).str.strip().str.upper()
    d["Code"] = names.map(SUBSYSTEM_CODES)
    unknown = sorted(set(names[d["Code"].isna()]))
    if unknown:
        print(f"WARNING: unmapped subsystem names ignored: {unknown}", flush=True)
    d = d.dropna(subset=["Date", "Code"])

    wide = {}
    for value, suffix in (("Pct", "pct"), ("Energy", "MWmes"), ("Max", "max_MWmes")):
        wide[suffix] = d.pivot_table(index="Date", columns="Code", values=value, aggfunc="last").reindex(columns=CODES)
    # SIN national % = total stored energy / total maximum storable
    # energy across the four subsystems (energy-weighted, the way ONS
    # computes its own SIN EAR %). Only on days all four report.
    complete = wide["MWmes"].notna().all(axis=1) & wide["max_MWmes"].notna().all(axis=1)
    sin_energy = wide["MWmes"].sum(axis=1).where(complete)
    sin_max = wide["max_MWmes"].sum(axis=1).where(complete)

    out = pd.DataFrame(index=wide["pct"].index)
    for c in CODES:
        out[f"{c}_pct"] = wide["pct"][c]
    out["SIN_pct"] = sin_energy / sin_max * 100
    for c in CODES:
        out[f"{c}_MWmes"] = wide["MWmes"][c]
    out["SIN_MWmes"] = sin_energy
    out["SIN_max_MWmes"] = sin_max
    out = out.dropna(how="all")

    if OUT_KEEP is not None and len(OUT_KEEP):
        out = pd.concat([OUT_KEEP.reindex(columns=out.columns), out[out.index > OUT_KEEP.index.max()]])
    out = out[~out.index.duplicated(keep="last")].sort_index().round(4)
    out.index = pd.DatetimeIndex(out.index).date
    out.index.name = "date"

    notes = [
        "UNITS",
        "*_pct: reservoir storage, % of maximum storable energy (EAR % - energia armazenada).",
        "*_MWmes: stored energy, MW-month (MWmes = average MW sustained for one month; x 0.73 = GWh approx).",
        "SIN_max_MWmes: maximum storable energy of the whole interconnected system (sum of the subsystems' EARmax).",
        "",
        "SUBSYSTEMS",
        "SE_CO = Sudeste/Centro-Oeste (about 70% of Brazil's storage capacity), S = Sul, NE = Nordeste, N = Norte.",
        "SIN = Sistema Interligado Nacional (the national total).",
        "",
        "SIN NATIONAL %",
        "ONS's subsystem dataset carries no national row, so SIN_pct is energy-weighted:",
        "sum of the four subsystems' stored energy (ear_verif_subsistema_mwmes) / sum of their maximum storable",
        "energy (ear_max_subsistema) x 100 - the same definition ONS uses for its SIN EAR %. Left blank on any day",
        "a subsystem is missing.",
        "",
        "SOURCE",
        "ONS open data (dados.ons.org.br), dataset ear_subsistema_di (EAR diario por subsistema), yearly files",
        BASE + "<year>.csv",
        f"Daily, from {out.index.min()} to {out.index.max()}. Updated daily by .github/workflows/brazil_hydro_reservoirs.yml;",
        "only the current year (and last year until April, while ONS may revise it) is re-downloaded each run.",
    ]
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    xlsx_notes.write_workbook(path, {"Daily": out}, notes, {"UNITS", "SUBSYSTEMS", "SIN NATIONAL %", "SOURCE"})
    print(f"\nWrote {path}: {len(out):,} days, {out.index.min()} to {out.index.max()}", flush=True)
    print(out.tail(5).to_string(float_format="%.1f"), flush=True)


if ARGS.out:
    write_out_workbook(ARGS.out)
    sys.exit(0)


# ----------------------------------------------------------
# MONTHLY STORAGE %
# ----------------------------------------------------------

monthly_pct = (
    daily_pct
    .set_index("Date")
    .resample("ME")
    .last()
    .reset_index()
    .rename(
        columns={
            "Date": "MonthEnd"
        }
    )
)

# ----------------------------------------------------------
# MONTHLY STORED ENERGY
# ----------------------------------------------------------

monthly_energy = (
    daily_energy
    .set_index("Date")
    .resample("ME")
    .last()
    .reset_index()
    .rename(
        columns={
            "Date": "MonthEnd"
        }
    )
)

# ----------------------------------------------------------
# WRITE FILES
# ----------------------------------------------------------

daily_pct.to_csv(
    "ear_daily_pct.csv",
    index=False
)

monthly_pct.to_csv(
    "ear_monthly_pct.csv",
    index=False
)

daily_energy.to_csv(
    "ear_daily_energy.csv",
    index=False
)

monthly_energy.to_csv(
    "ear_monthly_energy.csv",
    index=False
)

print("\nLATEST MONTHLY STORAGE %")
print(
    monthly_pct.tail(12).to_string(
        index=False,
        float_format="%.1f"
    )
)

print("\nLATEST MONTHLY STORED ENERGY")
print(
    monthly_energy.tail(12).to_string(
        index=False,
        float_format="%.0f"
    )
)

print(
    "\nSaved:\n"
    "  ear_daily_pct.csv\n"
    "  ear_monthly_pct.csv\n"
    "  ear_daily_energy.csv\n"
    "  ear_monthly_energy.csv\n",
    flush=True
)