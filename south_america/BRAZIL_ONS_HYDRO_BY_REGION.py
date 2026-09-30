print("STARTING", flush=True)

import datetime as dt
import io
from pathlib import Path

import pandas as pd
import requests

START_YEAR = 2000
END_YEAR = dt.date.today().year

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