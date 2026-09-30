"""
South America dashboard: every South America chart in one workbook,
South_America_Dashboard.xlsx - a Summary tab first, then one tab per
chart with its data beside it. Native Excel charts written with
openpyxl, so it runs anywhere (desktop, GitHub) and Excel needn't be
open or even installed.

Reads (doesn't refetch) the outputs of the pulls MASTER_SOUTH_AMERICA.py
runs first. A missing input skips its chart (named in the Summary tab)
rather than stopping the rest:
  Brazil:
  - "ONS Brazil.py" -> brazil_generation_daily.xlsx, "Data" tab:
    national daily generation by source (MW, daily mean).
  - BRAZIL_ONS_HYDRO_BY_REGION.py -> ear_daily_pct.csv / ear_daily_energy.csv:
    reservoir storage % and stored energy (MWmes) per ONS subsystem.
  Colombia:
  - COLOMBIA_XM_HYDRO.py -> co_hydro_daily_pct.csv: national reservoir
    storage % (XM publishes the national figure directly).
  - COLOMBIA_XM_GENERATION.py -> colombia_generation_daily.xlsx, "Data"
    tab: daily mean MW by fuel (hydro, gas, coal, liquids, wind, solar...).
  Chile:
  - CHILE_CEN_HYDRO.py -> chile_hydro_cota_history.csv: one row per
    reservoir per day it has run (CEN's API only gives the latest
    reading - needs CEN_USER_KEY).
  Uruguay:
  - URUGUAY_ADME.py -> uruguay_generation.xlsx, "Daily" tab: daily mean
    MW by source (ADME SCADA; Salto Grande = Uruguay's half).
  Ecuador:
  - ECUADOR_CENACE.py -> ecuador_generation.xlsx, "Daily" tab: MWh and GW
    by source from CENACE (last complete day each run - no archive).
  Argentina:
  - argentina_generation_mix.py -> argentina_generation_mix_daily.xlsx,
    "Data" tab: daily MWh by source (CAMMESA dispatch programme).

Charts:
  - Brazil Generation: stacked-area GW generation mix (last 3 years).
  - Brazil Thermal vs Range: thermal generation (7-day average, GW)
    against its 5-year seasonal range - the gas-burn / LNG-pull signal:
    thermals run hard when reservoirs are low.
  - Brazil Hydro Generation: hydro converted to Bcf/d gas equivalent -
    what a gas fleet would burn to replace it (same assumption as
    europe/Ember_EUUK_Wind_Solar.py).
  - Brazil Hydro Storage: national reservoir storage %, capacity-
    weighted across the four ONS subsystems (see
    read_hydro_storage_national()), against its 5-year range.
  - Colombia Hydro Storage: same 5-year-range treatment.
  - Argentina Generation: stacked-area GW mix (thermal = gas-dominated
    steam, gas turbine, combined cycle and diesel units together).
  - Uruguay Generation: stacked-area GW by source (hydro, wind, solar,
    thermal, biomass), plus Uruguay hydro vs its 5-year range.
  - Argentina Hydro vs Range: hydro generation against the range since
    CAMMESA's generator data starts (Oct-2024). Argentina has no national reservoir
    feed; ~60% of its hydro is run-of-river (Yacyreta, Salto Grande).
  - Uruguay Reservoir: Rincon del Bonete lake level (m) against its range
    (ADME publishes levels, not stored volume; history from 2020).
  - Salto Grande Inflows: daily inflow to the dam Uruguay and Argentina
    share (7-day average, m3/s) against its 5-year range.
  - Ecuador Generation: stacked-area GW - hydro, oil/diesel thermal, gas,
    renewables. CENACE publishes no reservoir levels on that page.
  - Chile Hydro Levels: one line per reservoir in cota (metres) - no
    %-of-capacity or history is published, so it grows run by run.
  - Colombia Generation: stacked-area GW by fuel, gas split out
    (COLOMBIA_XM_GENERATION.py - XM runs about two weeks behind).
"""

import sys
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import AreaChart, LineChart, Reference
from openpyxl.chart.axis import DateAxis
from openpyxl.styles import Alignment, Font, PatternFill

BASE_DIRECTORY = Path(__file__).resolve().parent

GENERATION_WORKBOOK = BASE_DIRECTORY / "brazil_generation_daily.xlsx"
GENERATION_SHEET = "Data"

HYDRO_CSV = BASE_DIRECTORY / "ear_daily_pct.csv"
HYDRO_ENERGY_CSV = BASE_DIRECTORY / "ear_daily_energy.csv"

COLOMBIA_HYDRO_PCT_CSV = BASE_DIRECTORY / "co_hydro_daily_pct.csv"

CHILE_HYDRO_CSV = BASE_DIRECTORY / "chile_hydro_cota_history.csv"

ARGENTINA_WORKBOOK = BASE_DIRECTORY / "argentina_generation_mix_daily.xlsx"

COLOMBIA_GENERATION_WORKBOOK = BASE_DIRECTORY / "colombia_generation_daily.xlsx"

URUGUAY_WORKBOOK = BASE_DIRECTORY / "uruguay_generation.xlsx"

ECUADOR_WORKBOOK = BASE_DIRECTORY / "ecuador_generation.xlsx"

OUTPUT_WORKBOOK = BASE_DIRECTORY / "South_America_Dashboard.xlsx"

# Generation charts show the last few years only; the seasonal-range
# charts use this many full years before the current one.
CHART_HISTORY_YEARS = 3
HYDRO_SEASONAL_HISTORICAL_YEARS = 5
ROLLING_DAYS = 7

# ============================================================
# GAS EQUIVALENT (Bcf/d) - same assumption/conversion as
# europe/Ember_EUUK_Wind_Solar.py's wind/solar Bcf/d charts, for
# consistency across the whole set of dashboards.
# ============================================================

GAS_PLANT_EFFICIENCY = 0.50
NATURAL_GAS_BTU_PER_SCF = 1037
MMBTU_PER_GWH = 3412.142
BCF_PER_GWH_D = MMBTU_PER_GWH / (NATURAL_GAS_BTU_PER_SCF / 1000) / 1_000_000


def gw_to_gas_equivalent_bcfd(generation_gw):
    """Convert an average daily generation rate (GW) into the Bcf/d of
    natural gas a gas-fired power fleet would have burned to generate
    the same amount, at GAS_PLANT_EFFICIENCY."""
    if pd.isna(generation_gw):
        return np.nan
    gas_thermal_gw = generation_gw / GAS_PLANT_EFFICIENCY
    gas_thermal_gwh_per_day = gas_thermal_gw * 24
    return gas_thermal_gwh_per_day * BCF_PER_GWH_D


# ============================================================
# READ SOURCE DATA
# ============================================================

def read_generation():
    if not GENERATION_WORKBOOK.exists():
        raise FileNotFoundError(
            f"Generation workbook not found: {GENERATION_WORKBOOK}\n"
            "Run 'ONS Brazil.py' first (MASTER_SOUTH_AMERICA.py does this automatically)."
        )
    df = pd.read_excel(GENERATION_WORKBOOK, sheet_name=GENERATION_SHEET, index_col=0)
    df.index = pd.to_datetime(df.index)
    df = df.sort_index()
    # MW daily mean -> GW, easier to read on a chart axis.
    return (df / 1000).rename(columns=lambda c: c.replace("_", " ").title())


def read_hydro_generation_gas_equivalent():
    """Hydro's own generation column, converted to Bcf/d gas
    equivalent - a "what would it take to replace hydro" figure, not
    part of the Generation Mix stack (different units)."""
    if not GENERATION_WORKBOOK.exists():
        raise FileNotFoundError(
            f"Generation workbook not found: {GENERATION_WORKBOOK}\n"
            "Run 'ONS Brazil.py' first (MASTER_SOUTH_AMERICA.py does this automatically)."
        )
    df = pd.read_excel(GENERATION_WORKBOOK, sheet_name=GENERATION_SHEET, index_col=0)
    df.index = pd.to_datetime(df.index)
    df = df.sort_index()
    hydro_gw = df["hydro"] / 1000  # MW -> GW
    bcfd = hydro_gw.apply(gw_to_gas_equivalent_bcfd)
    return bcfd.rename("Hydro Gas Equivalent").to_frame()


def read_hydro():
    if not HYDRO_CSV.exists():
        raise FileNotFoundError(
            f"Hydro storage CSV not found: {HYDRO_CSV}\n"
            "Run BRAZIL_ONS_HYDRO_BY_REGION.py first (MASTER_SOUTH_AMERICA.py does this automatically)."
        )
    df = pd.read_csv(HYDRO_CSV)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"]).set_index("Date").sort_index()
    return df


def read_hydro_energy():
    """Stored energy (MWmes) per ONS subsystem - BRAZIL_ONS_HYDRO_BY_REGION.py's
    other daily output, same pivot shape as read_hydro()'s percentage
    table. Used only to derive each subsystem's implied capacity for
    read_hydro_storage_national() below - not charted directly."""
    if not HYDRO_ENERGY_CSV.exists():
        raise FileNotFoundError(
            f"Hydro stored-energy CSV not found: {HYDRO_ENERGY_CSV}\n"
            "Run BRAZIL_ONS_HYDRO_BY_REGION.py first (MASTER_SOUTH_AMERICA.py does this automatically)."
        )
    df = pd.read_csv(HYDRO_ENERGY_CSV)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"]).set_index("Date").sort_index()
    return df


def read_hydro_storage_national():
    """A single national reservoir-storage-% series, properly
    capacity-weighted: no subsystem capacity figure is pulled
    directly, but it's implied by the two series that ARE already
    pulled - capacity = StoredEnergy / (StoragePct / 100), per
    subsystem per day. Energy and capacity are both additive across
    subsystems (percentages aren't), so summing each separately before
    dividing gives a properly weighted national %, rather than an
    unweighted mean across regions of very different reservoir sizes."""
    pct_df = read_hydro()
    energy_df = read_hydro_energy()

    pct_df, energy_df = pct_df.align(energy_df, join="inner", axis=None)

    # A subsystem reporting 0% (or missing) that day can't imply a
    # capacity - exclude it from that day's national total rather than
    # dividing by zero/NaN. Its energy is excluded for that same day
    # too, not just its capacity - otherwise the numerator would still
    # carry that subsystem's energy while the denominator drops its
    # capacity, understating the true % (energy with no capacity behind
    # it inflates the ratio).
    safe_pct = pct_df.where(pct_df > 0)
    capacity_df = energy_df / (safe_pct / 100)
    safe_energy = energy_df.where(safe_pct.notna())

    national_energy = safe_energy.sum(axis=1, min_count=1)
    national_capacity = capacity_df.sum(axis=1, min_count=1)
    national_pct = (national_energy / national_capacity) * 100
    return national_pct.rename("NationalPct")


def read_colombia_hydro_storage_national():
    """Colombia's national reservoir storage % - unlike Brazil, no
    capacity-weighting derivation needed here: XM (Colombia's grid
    operator) already publishes VolUti (stored energy) and CapUti
    (capacity) at the national ("Sistema") level directly, and
    COLOMBIA_XM_HYDRO.py computes StoragePct = VolUti/CapUti*100 at
    pull time. Same output shape as read_hydro_storage_national()
    (a Series named "NationalPct"), so build_seasonal_range() works
    unchanged for either country."""
    if not COLOMBIA_HYDRO_PCT_CSV.exists():
        raise FileNotFoundError(
            f"Colombia hydro storage CSV not found: {COLOMBIA_HYDRO_PCT_CSV}\n"
            "Run COLOMBIA_XM_HYDRO.py first (MASTER_SOUTH_AMERICA.py does this automatically)."
        )
    df = pd.read_csv(COLOMBIA_HYDRO_PCT_CSV)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"]).set_index("Date").sort_index()
    return df["StoragePct"].rename("NationalPct")


def read_chile_hydro():
    """One column per reservoir, cota (metres) - not a %, see
    CHILE_CEN_HYDRO.py's docstring for why. History only goes back as
    far as that script has been run (no bulk backfill available), so
    this can be a short table, especially early on."""
    if not CHILE_HYDRO_CSV.exists():
        raise FileNotFoundError(
            f"Chile hydro cota history CSV not found: {CHILE_HYDRO_CSV}\n"
            "Run CHILE_CEN_HYDRO.py first (MASTER_SOUTH_AMERICA.py does this automatically)."
        )
    df = pd.read_csv(CHILE_HYDRO_CSV)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    wide = df.pivot_table(index="Date", columns="Reservoir", values="Cota", aggfunc="last")
    return wide.sort_index()


def clip_to_recent_years(df, years):
    if df.empty:
        return df
    cutoff = df.index.max() - pd.DateOffset(years=years)
    return df[df.index >= cutoff]


def build_seasonal_range(daily_series, historical_years=HYDRO_SEASONAL_HISTORICAL_YEARS):
    """Build a 366-row (Jan1..Dec31) seasonal profile from a daily
    Series: Base (historical min), Band (historical max - min),
    Average, Prior (year), Current (year) - one row per calendar day-
    of-year, same structural pattern as this repo's AGSI seasonal-
    range gas storage chart (europe/AGSI_SCRAPE.py), but keyed on
    plain calendar day-of-year rather than a gas-year offset."""
    frame = daily_series.rename("value").to_frame()
    frame["year"] = frame.index.year
    frame["day_of_year"] = frame.index.dayofyear

    latest_date = daily_series.index.max()
    current_year = latest_date.year
    prior_year = current_year - 1
    historical_start = current_year - historical_years
    historical_years_present = sorted(
        y for y in frame["year"].unique() if historical_start <= y < current_year
    )

    historical = frame[frame["year"].isin(historical_years_present)]
    seasonal = historical.groupby("day_of_year")["value"].agg(["min", "max", "mean"])
    seasonal.columns = ["Min", "Max", "Average"]

    current = (
        frame[frame["year"] == current_year]
        .groupby("day_of_year")["value"]
        .last()
        .rename("Current")
    )
    prior = (
        frame[frame["year"] == prior_year]
        .groupby("day_of_year")["value"]
        .last()
        .rename("Prior")
    )

    day_table = pd.DataFrame(index=pd.RangeIndex(1, 367, name="DayOfYear"))
    day_table["ChartDate"] = pd.Timestamp("2001-01-01") + pd.to_timedelta(
        day_table.index - 1, unit="D"
    )

    combined = day_table.join(seasonal).join(current).join(prior)
    combined["Base"] = combined["Min"]
    combined["Band"] = combined["Max"] - combined["Min"]

    result = combined[["ChartDate", "Base", "Band", "Average", "Prior", "Current"]].copy()
    result = result.set_index("ChartDate")
    result.index.name = "Date"
    return result




# ============================================================
# MORE SOURCES
# ============================================================

def read_brazil_thermal_gw():
    """National thermal generation, GW, 7-day average - smooths the
    weekday/weekend swing so the seasonal range reads cleanly."""
    return read_generation()["Thermal"].rolling(ROLLING_DAYS, min_periods=ROLLING_DAYS).mean().dropna().rename("NationalGW")


def read_argentina_generation():
    """Daily MWh by source -> average GW (MWh / 24 / 1000)."""
    if not ARGENTINA_WORKBOOK.exists():
        raise FileNotFoundError(f"{ARGENTINA_WORKBOOK.name} not found - run argentina_generation_mix.py first.")
    df = pd.read_excel(ARGENTINA_WORKBOOK, sheet_name="Data", index_col=0)
    df.index = pd.to_datetime(df.index)
    df = df.dropna(how="all")  # blank rows = days CAMMESA published no generator data
    df = df[df.sum(axis=1) > 0].sort_index() / 24 / 1000
    return df.rename(columns=lambda c: c.replace("_", " ").title())


def read_colombia_generation():
    """Daily mean MW by fuel -> GW."""
    if not COLOMBIA_GENERATION_WORKBOOK.exists():
        raise FileNotFoundError(f"{COLOMBIA_GENERATION_WORKBOOK.name} not found - run COLOMBIA_XM_GENERATION.py first.")
    df = pd.read_excel(COLOMBIA_GENERATION_WORKBOOK, sheet_name="Data", index_col=0)
    df.index = pd.to_datetime(df.index)
    return (df.sort_index() / 1000).rename(columns=lambda c: c.replace("_", " ").title())


def read_uruguay_generation():
    """Daily mean MW by source -> GW (generation sources only)."""
    if not URUGUAY_WORKBOOK.exists():
        raise FileNotFoundError(f"{URUGUAY_WORKBOOK.name} not found - run URUGUAY_ADME.py first.")
    df = pd.read_excel(URUGUAY_WORKBOOK, sheet_name="Daily", index_col=0)
    df.index = pd.to_datetime(df.index)
    df = df[[c for c in ("hydro", "wind", "solar", "thermal", "biomass") if c in df.columns]]
    return (df.sort_index() / 1000).rename(columns=lambda c: c.title())


def read_ecuador_generation():
    """Daily mean GW by source (hydro, oil/diesel thermal, gas, renewables)."""
    if not ECUADOR_WORKBOOK.exists():
        raise FileNotFoundError(f"{ECUADOR_WORKBOOK.name} not found - run ECUADOR_CENACE.py first.")
    df = pd.read_excel(ECUADOR_WORKBOOK, sheet_name="Daily", index_col=0)
    df.index = pd.to_datetime(df.index)
    cols = {"hydro_gw": "Hydro", "thermal_oil_gw": "Oil/diesel", "thermal_gas_gw": "Gas", "renewable_gw": "Renewables"}
    df = df[[c for c in cols if c in df.columns]].rename(columns=cols)
    return df.sort_index()


def read_bonete_level():
    """Rincon del Bonete lake level (m above sea level), daily - Uruguay's
    main hydro storage."""
    if not URUGUAY_WORKBOOK.exists():
        raise FileNotFoundError(f"{URUGUAY_WORKBOOK.name} not found - run URUGUAY_ADME.py first.")
    try:
        df = pd.read_excel(URUGUAY_WORKBOOK, sheet_name="Reservoir levels", index_col=0)
    except ValueError:
        raise FileNotFoundError("no 'Reservoir levels' tab yet - run URUGUAY_ADME.py again.")
    df.index = pd.to_datetime(df.index)
    return df["Bonete level_m"].sort_index().dropna().rename("NationalPct")


def read_salto_grande_inflow():
    """Salto Grande inflow (Aportes), m3/s, 7-day average - the Uruguay
    river feeding the dam Uruguay and Argentina share."""
    if not URUGUAY_WORKBOOK.exists():
        raise FileNotFoundError(f"{URUGUAY_WORKBOOK.name} not found - run URUGUAY_ADME.py first.")
    try:
        df = pd.read_excel(URUGUAY_WORKBOOK, sheet_name="Salto Grande flows", index_col=0)
    except ValueError:
        raise FileNotFoundError("no 'Salto Grande flows' tab yet - run URUGUAY_ADME.py again.")
    df.index = pd.to_datetime(df.index)
    col = next(c for c in df.columns if c.lower().startswith("aportes"))
    return df[col].sort_index().rolling(ROLLING_DAYS, min_periods=ROLLING_DAYS).mean().dropna().rename("NationalPct")


# ============================================================
# WORKBOOK
# ============================================================

# Fixed colour per source across every chart (Brazil and Argentina alike)
SOURCE_COLOURS = {"Hydro": "2A78D6", "Thermal": "EB6834", "Wind": "1BAF7A", "Solar": "EDA100",
                  "Nuclear": "4A3AA7", "Biomass": "008300", "Gas": "EB6834", "Coal": "52514E",
                  "Liquids": "E34948", "Unclassified": "C3C2B7", "Oil/diesel": "E34948", "Renewables": "1BAF7A"}
FALLBACK_COLOURS = ["E87BA4", "E34948", "7F7F7F", "9A9A94"]
LINE_BLACK, LINE_GREY, LINE_ORANGE = "0B0B0B", "808080", "ED7D31"
RANGE_FILL = "9BBED7"
HEADER_FILL = PatternFill("solid", fgColor="DCE6F0")


def write_table(ws, df, date_format="dd-mmm-yyyy"):
    """Data table from A1 (dates in column A); returns the last row."""
    ws.cell(1, 1, df.index.name or "Date").font = Font(bold=True)
    for j, col in enumerate(df.columns, 2):
        ws.cell(1, j, str(col)).font = Font(bold=True)
    for i, (idx, row) in enumerate(df.iterrows(), 2):
        ws.cell(i, 1, pd.Timestamp(idx).to_pydatetime()).number_format = date_format
        for j, v in enumerate(row, 2):
            if not pd.isna(v):
                ws.cell(i, j, round(float(v), 3))
    ws.column_dimensions["A"].width = 12
    ws.freeze_panes = "B2"
    return len(df) + 1


def date_axis(chart, number_format, major_unit, unit):
    chart.x_axis = DateAxis(crossAx=100)
    chart.x_axis.number_format = number_format
    chart.x_axis.majorTimeUnit = unit
    chart.x_axis.majorUnit = major_unit
    chart.x_axis.title = None


def finish(chart, title, y_title, anchor_col, ws, y_min=0):
    chart.title = title
    chart.y_axis.title = y_title
    if y_min is not None:
        chart.y_axis.scaling.min = y_min
    chart.y_axis.majorGridlines.spPr = None
    chart.legend.position = "b"
    chart.width, chart.height = 26, 13
    ws.add_chart(chart, f"{anchor_col}2")


def anchor_after(df):
    from openpyxl.utils import get_column_letter
    return get_column_letter(len(df.columns) + 3)


def stacked_generation_sheet(wb, name, df, title, unit="GW"):
    ws = wb.create_sheet(name)
    last = write_table(ws, df)
    chart = AreaChart()
    chart.grouping = "stacked"
    chart.add_data(Reference(ws, min_col=2, max_col=1 + len(df.columns), min_row=1, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
    spare = iter(FALLBACK_COLOURS)
    for series, col in zip(chart.series, df.columns):
        series.graphicalProperties.solidFill = SOURCE_COLOURS.get(col) or next(spare)
        series.graphicalProperties.line.noFill = True
    date_axis(chart, "mmm-yy", 3, "months")
    finish(chart, title, unit, anchor_after(df), ws)


def line_sheet(wb, name, df, title, y_title, colours=None, y_min=0, date_fmt="mmm-yy", major=(3, "months")):
    ws = wb.create_sheet(name)
    last = write_table(ws, df)
    chart = LineChart()
    chart.add_data(Reference(ws, min_col=2, max_col=1 + len(df.columns), min_row=1, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
    for i, series in enumerate(chart.series):
        series.smooth = False
        series.graphicalProperties.line.width = 22860
        if colours:
            series.graphicalProperties.line.solidFill = colours[i % len(colours)]
    date_axis(chart, date_fmt, *major)
    finish(chart, title, y_title, anchor_after(df), ws, y_min=y_min)


def range_sheet(wb, name, df, title, y_title, y_min=0):
    """Base (invisible) + Band stacked areas give the 5-year min-max
    range; Average, Prior and Current years are lines on top. Columns
    must be build_seasonal_range()'s: Base, Band, Average, Prior, Current."""
    ws = wb.create_sheet(name)
    last = write_table(ws, df, date_format="dd-mmm")
    area = AreaChart()
    area.grouping = "stacked"
    area.add_data(Reference(ws, min_col=2, max_col=3, min_row=1, max_row=last), titles_from_data=True)
    area.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
    base, band = area.series
    base.graphicalProperties.noFill = True
    base.graphicalProperties.line.noFill = True
    band.graphicalProperties.solidFill = RANGE_FILL
    band.graphicalProperties.line.noFill = True
    from openpyxl.chart.series import SeriesLabel
    band.tx = SeriesLabel(v=f"{HYDRO_SEASONAL_HISTORICAL_YEARS}-year range")
    base.tx = SeriesLabel(v=" ")
    line = LineChart()
    line.add_data(Reference(ws, min_col=4, max_col=6, min_row=1, max_row=last), titles_from_data=True)
    labels = [f"{HYDRO_SEASONAL_HISTORICAL_YEARS}-year average", "Prior year", "Current year"]
    styles = [(LINE_GREY, 15875, "dash"), (LINE_ORANGE, 15875, "dash"), (LINE_BLACK, 28575, None)]
    for series, label, (colour, width, dash) in zip(line.series, labels, styles):
        series.tx = SeriesLabel(v=label)
        series.smooth = False
        series.graphicalProperties.line.solidFill = colour
        series.graphicalProperties.line.width = width
        if dash:
            series.graphicalProperties.line.dashStyle = dash
    line.y_axis.axId = area.y_axis.axId
    area += line
    date_axis(area, "mmm", 1, "months")
    finish(area, title, y_title, anchor_after(df), ws, y_min=y_min)


# ============================================================
# SUMMARY
# ============================================================

def summary_row(label, unit, series):
    """Latest value, and how it compares with the same day last year and
    the 5-year average for that day of year."""
    s = series.dropna()
    if s.empty:
        return [label, unit, None, None, None, None]
    last_date = s.index.max()
    latest = s.iloc[-1]
    ly = s[s.index <= last_date - pd.DateOffset(years=1)]
    last_year = ly.iloc[-1] if len(ly) and (last_date - pd.DateOffset(years=1) - ly.index.max()).days <= 7 else None
    hist = s[(s.index.year >= last_date.year - HYDRO_SEASONAL_HISTORICAL_YEARS) & (s.index.year < last_date.year)]
    avg = hist[hist.index.dayofyear == last_date.dayofyear].mean() if len(hist) else None
    return [label, unit, last_date.date(), round(latest, 2),
            None if last_year is None else round(latest - last_year, 2),
            None if avg is None or pd.isna(avg) else round(latest - avg, 2)]


def write_summary(ws, rows, built, skipped):
    ws["A1"] = "South America power & hydro - summary"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Built {date.today():%d-%b-%Y}. Each chart is on its own tab, data beside it."
    headers = ["Metric", "Unit", "Latest day", "Latest", "vs same day last year", f"vs {HYDRO_SEASONAL_HISTORICAL_YEARS}-yr avg for the day"]
    for j, h in enumerate(headers, 1):
        c = ws.cell(4, j, h)
        c.font = Font(bold=True)
        c.fill = HEADER_FILL
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for i, row in enumerate(rows, 5):
        for j, v in enumerate(row, 1):
            c = ws.cell(i, j, v)
            if j == 3 and v is not None:
                c.number_format = "dd-mmm-yyyy"
            if j >= 5 and isinstance(v, (int, float)):
                c.number_format = "+0.00;-0.00;0.00"
    r = len(rows) + 6
    ws.cell(r, 1, "Charts").font = Font(bold=True)
    for k, name in enumerate(built, 1):
        ws.cell(r + k, 1, name)
    if skipped:
        r += len(built) + 2
        ws.cell(r, 1, "Skipped (input missing - run its pull script)").font = Font(bold=True)
        for k, (name, why) in enumerate(skipped, 1):
            ws.cell(r + k, 1, name)
            ws.cell(r + k, 2, why)
    for col, width in zip("ABCDEF", (44, 10, 13, 10, 16, 18)):
        ws.column_dimensions[col].width = width


# ============================================================
# MAIN
# ============================================================

def main():
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    built, skipped, rows = [], [], []

    def attempt(name, fn):
        try:
            fn()
            built.append(name)
            print(f"  {name}: built", file=sys.stderr, flush=True)
        except FileNotFoundError as e:
            why = str(e).splitlines()[0].replace(str(BASE_DIRECTORY) + "/", "").replace(str(BASE_DIRECTORY) + "\\", "")
            skipped.append((name, why))
            print(f"  {name}: SKIPPED - {why}", file=sys.stderr, flush=True)

    def brazil_generation():
        stacked_generation_sheet(wb, "Brazil Generation", clip_to_recent_years(read_generation(), CHART_HISTORY_YEARS),
                                 "Brazil generation mix (GW)")
        gen = read_generation()
        for col in ("Hydro", "Thermal", "Wind", "Solar"):
            rows.append(summary_row(f"Brazil {col.lower()} generation ({ROLLING_DAYS}-day avg)", "GW",
                                    gen[col].rolling(ROLLING_DAYS, min_periods=ROLLING_DAYS).mean()))

    def brazil_thermal():
        range_sheet(wb, "Brazil Thermal vs Range", build_seasonal_range(read_brazil_thermal_gw()),
                    f"Brazil thermal generation ({ROLLING_DAYS}-day avg) vs {HYDRO_SEASONAL_HISTORICAL_YEARS}-year range",
                    "GW")

    def brazil_hydro_gen():
        line_sheet(wb, "Brazil Hydro Generation", clip_to_recent_years(read_hydro_generation_gas_equivalent(), CHART_HISTORY_YEARS),
                   "Brazil hydro generation, gas equivalent (Bcf/d)", "Bcf/d", colours=[LINE_BLACK])

    def brazil_storage():
        national = read_hydro_storage_national()
        range_sheet(wb, "Brazil Hydro Storage", build_seasonal_range(national),
                    f"Brazil reservoir storage vs {HYDRO_SEASONAL_HISTORICAL_YEARS}-year range", "% of capacity")
        rows.insert(0, summary_row("Brazil reservoir storage (national, capacity-weighted)", "%", national))

    def colombia_storage():
        national = read_colombia_hydro_storage_national()
        range_sheet(wb, "Colombia Hydro Storage", build_seasonal_range(national),
                    f"Colombia reservoir storage vs {HYDRO_SEASONAL_HISTORICAL_YEARS}-year range", "% of capacity")
        rows.append(summary_row("Colombia reservoir storage (national)", "%", national))

    def argentina_generation():
        gen = read_argentina_generation()
        stacked_generation_sheet(wb, "Argentina Generation", clip_to_recent_years(gen, CHART_HISTORY_YEARS),
                                 "Argentina generation mix (GW, dispatch programme)")
        range_sheet(wb, "Argentina Hydro vs Range",
                    build_seasonal_range(gen["Hydro"].rolling(ROLLING_DAYS, min_periods=ROLLING_DAYS).mean().dropna()),
                    f"Argentina hydro generation ({ROLLING_DAYS}-day avg) vs range since Oct-2024", "GW")
        rows.append(summary_row(f"Argentina hydro generation ({ROLLING_DAYS}-day avg)", "GW",
                                gen["Hydro"].rolling(ROLLING_DAYS, min_periods=ROLLING_DAYS).mean()))
        rows.append(summary_row(f"Argentina thermal generation ({ROLLING_DAYS}-day avg)", "GW",
                                gen["Thermal"].rolling(ROLLING_DAYS, min_periods=ROLLING_DAYS).mean()))

    def colombia_generation():
        gen = read_colombia_generation()
        stacked_generation_sheet(wb, "Colombia Generation", clip_to_recent_years(gen, CHART_HISTORY_YEARS),
                                 "Colombia generation by fuel (GW)")
        for col in ("Hydro", "Gas", "Coal"):
            if col in gen:
                rows.append(summary_row(f"Colombia {col.lower()} generation ({ROLLING_DAYS}-day avg)", "GW",
                                        gen[col].rolling(ROLLING_DAYS, min_periods=ROLLING_DAYS).mean()))

    def uruguay_generation():
        gen = read_uruguay_generation()
        stacked_generation_sheet(wb, "Uruguay Generation", clip_to_recent_years(gen, CHART_HISTORY_YEARS),
                                 "Uruguay generation by source (GW)")
        range_sheet(wb, "Uruguay Hydro vs Range",
                    build_seasonal_range(gen["Hydro"].rolling(ROLLING_DAYS, min_periods=ROLLING_DAYS).mean().dropna()),
                    f"Uruguay hydro generation ({ROLLING_DAYS}-day avg) vs {HYDRO_SEASONAL_HISTORICAL_YEARS}-year range", "GW")
        for col in ("Hydro", "Thermal"):
            rows.append(summary_row(f"Uruguay {col.lower()} generation ({ROLLING_DAYS}-day avg)", "GW",
                                    gen[col].rolling(ROLLING_DAYS, min_periods=ROLLING_DAYS).mean()))

    def uruguay_reservoir():
        level = read_bonete_level()
        range_sheet(wb, "Uruguay Reservoir", build_seasonal_range(level),
                    f"Rincon del Bonete lake level vs {HYDRO_SEASONAL_HISTORICAL_YEARS}-year range (Uruguay's main storage)",
                    "metres above sea level", y_min=None)
        rows.append(summary_row("Uruguay: Rincon del Bonete lake level", "m", level))

    def ecuador_generation():
        gen = read_ecuador_generation()
        stacked_generation_sheet(wb, "Ecuador Generation", clip_to_recent_years(gen, CHART_HISTORY_YEARS),
                                 "Ecuador generation by source (GW) - history builds daily from the first run")
        for col in ("Hydro", "Oil/diesel", "Gas"):
            if col in gen:
                s = gen[col].dropna()
                rows.append(["Ecuador " + col.lower() + " generation (latest day)", "GW",
                             s.index.max().date() if len(s) else None, round(float(s.iloc[-1]), 2) if len(s) else None,
                             None, None])

    def salto_grande():
        inflow = read_salto_grande_inflow()
        range_sheet(wb, "Salto Grande Inflows", build_seasonal_range(inflow),
                    f"Salto Grande inflows ({ROLLING_DAYS}-day avg) vs {HYDRO_SEASONAL_HISTORICAL_YEARS}-year range",
                    "m3/s")
        rows.append(summary_row(f"Salto Grande inflow, Uruguay river ({ROLLING_DAYS}-day avg)", "m3/s", inflow))

    def chile_levels():
        df = read_chile_hydro()
        line_sheet(wb, "Chile Hydro Levels", df, "Chile reservoir levels (cota, metres)", "metres", y_min=None,
                   date_fmt="dd-mmm", major=(7, "days"))

    print("Building South America dashboard...", file=sys.stderr, flush=True)
    for name, fn in [("Brazil Generation", brazil_generation), ("Brazil Thermal vs Range", brazil_thermal),
                     ("Brazil Hydro Generation", brazil_hydro_gen), ("Brazil Hydro Storage", brazil_storage),
                     ("Colombia Hydro Storage", colombia_storage), ("Colombia Generation", colombia_generation),
                     ("Argentina Generation", argentina_generation), ("Uruguay Generation", uruguay_generation),
                     ("Uruguay Reservoir", uruguay_reservoir), ("Salto Grande Inflows", salto_grande),
                     ("Ecuador Generation", ecuador_generation),
                     ("Chile Hydro Levels", chile_levels)]:
        attempt(name, fn)

    write_summary(summary, rows, built, skipped)
    try:
        wb.save(OUTPUT_WORKBOOK)
    except PermissionError as exc:
        raise RuntimeError(f"Can't write {OUTPUT_WORKBOOK} - it's probably open in Excel. Close it and rerun.") from exc
    print(f"Saved {OUTPUT_WORKBOOK} ({len(built)} charts, {len(skipped)} skipped)", file=sys.stderr, flush=True)
    if not built:
        sys.exit(1)


if __name__ == "__main__":
    main()
