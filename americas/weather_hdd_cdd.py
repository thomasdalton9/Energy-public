"""
Pull trailing-30-day daily temperature for three gas-demand weather
proxies - USA, NW Europe, and JKTC (Japan/Korea/Taiwan/China) - derive
Heating and Cooling Degree Days for each, and compare them against
each region's own 5-year average for the same calendar dates. Adds a
"Weather Chart" tab (one native chart per region: HDD and CDD, actual
vs 5-year average) to the SAME US_Gas_Storage_Dashboard.xlsx that the
other scripts in this folder build - not its own standalone file -
via win32com, the same way append_us_gas_balance.py and
Exports_USMexicoChart.py add their own tabs to that workbook without
touching anyone else's. (It also covers NW Europe/JKTC, not just the
US, but lands in the US workbook so MASTER_USA.py's run always
includes it - see MASTER_USA.py's SCRIPTS list.)

Each region is a population-weighted average across a handful of
major demand-centre cities (weighted by each city's approximate metro-
area population - see BASKETS - not a true population-weighted index
across every city/town in the region, just these proxies weighted
against each other):
    USA: New York, Chicago, Houston, Los Angeles, Atlanta.
    NW Europe: London, Paris, Amsterdam, Brussels, Frankfurt.
    JKTC: Tokyo, Seoul, Taipei, Shanghai.
Degree days use a single base temperature of 18C (~65F, the standard
HDD/CDD reference) for all three baskets, so they stay directly
comparable to each other.

The 5-year average for a given calendar date is the mean of that same
region's HDD/CDD on that exact month/day in each of the preceding 5
years (e.g. 16 Sep's normal = the average of 16 Sep 2021-2025) - a
simple year-over-year normal, not a smoothed climatology.

Data sources: both free, no API key required, same provider (Open-
Meteo) but two different endpoints/models:
  - Historical (trailing window + 5-year normal): ERA5 reanalysis,
    https://archive-api.open-meteo.com/v1/archive
  - Forecast (FORECAST_DAYS ahead of today): blended NWP models
    (ECMWF/GFS/ICON), https://api.open-meteo.com/v1/forecast - picks
    up immediately where the archive window ends (TO_DATE+1), so the
    combined series is continuous with no gap or overlap. Not
    verified live (Open-Meteo is unreachable from the sandbox this
    repo is normally edited in) - if the forecast endpoint rejects a
    start_date that far in the past (it's ARCHIVE_LAG_DAYS days
    behind today), first thing to check on a real run.

If the forecast fetch fails for any reason, the whole forecast step is
skipped (not per-city - a mix of forecast/no-forecast cities would
silently distort the equally-weighted basket average) and the chart
still gets its historical window, same as before this was added.

Requires pandas, requests, openpyxl, and pywin32 (Windows + Excel).
Close US_Gas_Storage_Dashboard.xlsx before running.
"""

import sys
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import requests
from openpyxl.utils import get_column_letter

# pythoncom/win32com are only needed for the Excel chart step, and
# pywin32's first-run COM type-cache generation can stall for a long
# time - importing them lazily (inside update_weather_chart, below)
# means the data-fetch phase below can run and print progress
# regardless, and if something does hang, it's obviously in the Excel
# step rather than looking like a silent hang from the first line.

BASE_DIRECTORY = Path(__file__).resolve().parent
ARCHIVE_URL = "https://archive-api.open-meteo.com/v1/archive"
FORECAST_URL = "https://api.open-meteo.com/v1/forecast"

# Open-Meteo's ERA5 archive typically lags a few days behind today, so
# the trailing window ends a few days back rather than at "today".
ARCHIVE_LAG_DAYS = 5
WINDOW_DAYS = 30
TO_DATE = date.today() - timedelta(days=ARCHIVE_LAG_DAYS)
FROM_DATE = TO_DATE - timedelta(days=WINDOW_DAYS - 1)

# How many days beyond TO_DATE to pull from Open-Meteo's forecast
# endpoint - picks up immediately where the archive window ends, so
# together they cover FROM_DATE through FORECAST_TO_DATE with no gap.
# 16 is Open-Meteo's own forecast horizon ceiling for the blended NWP
# models this uses - the last day or two can come back thin/missing
# for some cities as models disagree furthest out; that's fine here,
# reindex(current_dates) already leaves those as NaN rather than
# erroring.
FORECAST_DAYS = 16
FORECAST_FROM_DATE = TO_DATE + timedelta(days=1)
FORECAST_TO_DATE = TO_DATE + timedelta(days=FORECAST_DAYS)

# How many prior years' same calendar dates the "5-year average" is
# built from. Each city fetch covers this far back once, in one call,
# rather than one call per historical year.
HISTORY_YEARS = 5

DEGREE_DAY_BASE_C = 18.0

OUT_FILE = "US_Gas_Storage_Dashboard.xlsx"
WORKBOOK_PATH = BASE_DIRECTORY / OUT_FILE
CHART_SHEET_NAME = "Weather Chart"

XL_LINE = 4
XL_CATEGORY = 1
XL_VALUE = 2
XL_PRIMARY = 1
XL_TIME_SCALE = 3
XL_DAYS = 0
XL_MARKER_NONE = -4142
XL_LEGEND_BOTTOM = -4107
XL_CALCULATION_AUTOMATIC = -4105
MSO_TRUE = -1
MSO_FALSE = 0


def excel_rgb(red, green, blue):
    return red + green * 256 + blue * 65536


WHITE = excel_rgb(255, 255, 255)
# HDD cold blue, CDD warm orange - opposite hues so the two never blur
# together; each 5-year average takes its line's colour, dashed.
HDD_COLOUR = excel_rgb(42, 120, 214)
CDD_COLOUR = excel_rgb(235, 104, 52)
TODAY_LINE_COLOUR = excel_rgb(90, 90, 90)

# Proxy baskets, weighted by each city's approximate metro-area
# population (millions - the unit doesn't matter, only the ratios
# between cities in the same basket do). Figures are rounded, widely-
# cited metro/urban-area estimates (not a specific census date) - close
# enough for weighting a handful of demand-centre proxies against each
# other, not meant as precise official population statistics. Adjust
# freely if better figures are wanted.
BASKETS = {
    "USA": [
        ("New York", 40.71, -74.01, 19.5),
        ("Chicago", 41.85, -87.65, 9.5),
        ("Houston", 29.76, -95.37, 7.3),
        ("Los Angeles", 34.05, -118.24, 13.0),
        ("Atlanta", 33.75, -84.39, 6.3),
    ],
    "NW Europe": [
        ("London", 51.51, -0.13, 9.6),
        ("Paris", 48.85, 2.35, 12.3),
        ("Amsterdam", 52.37, 4.90, 2.5),
        ("Brussels", 50.85, 4.35, 2.1),
        ("Frankfurt", 50.11, 8.68, 5.8),
    ],
    "JKTC": [
        ("Tokyo", 35.68, 139.65, 37.0),
        ("Seoul", 37.57, 126.98, 25.6),
        ("Taipei", 25.03, 121.56, 7.0),
        ("Shanghai", 31.23, 121.47, 24.9),
    ],
}


def fetch_city_mean_temp(latitude, longitude, from_date, to_date):
    """Daily mean 2m air temperature (Celsius) for one city, from
    Open-Meteo's historical archive - a pandas Series indexed by
    date."""
    params = {
        "latitude": latitude,
        "longitude": longitude,
        "start_date": from_date.isoformat(),
        "end_date": to_date.isoformat(),
        "daily": "temperature_2m_mean",
        "timezone": "UTC",
    }
    response = requests.get(ARCHIVE_URL, params=params, timeout=60)
    response.raise_for_status()
    daily = response.json()["daily"]
    index = pd.to_datetime(daily["time"])
    return pd.Series(daily["temperature_2m_mean"], index=index, dtype=float)


def fetch_city_forecast_temp(latitude, longitude, from_date, to_date):
    """Daily mean 2m air temperature (Celsius) forecast for one city,
    from Open-Meteo's forecast API (blended NWP models, not the
    archive endpoint's ERA5 reanalysis) - a pandas Series indexed by
    date. See module docstring for the not-verified-live caveat."""
    params = {
        "latitude": latitude,
        "longitude": longitude,
        "start_date": from_date.isoformat(),
        "end_date": to_date.isoformat(),
        "daily": "temperature_2m_mean",
        "timezone": "UTC",
    }
    response = requests.get(FORECAST_URL, params=params, timeout=60)
    response.raise_for_status()
    daily = response.json()["daily"]
    index = pd.to_datetime(daily["time"])
    return pd.Series(daily["temperature_2m_mean"], index=index, dtype=float)


def weighted_average(frame, weights):
    """Row-wise population-weighted average across frame's columns.
    Weights are re-normalized per row over whichever columns actually
    have data that day, so one city's occasional missing reading
    doesn't zero out or skew the rest - a row with no data anywhere
    comes back NaN rather than 0."""
    weight_array = np.array([weights[column] for column in frame.columns], dtype=float)
    values = frame.to_numpy(dtype=float)
    present = ~np.isnan(values)
    weighted_sum = np.nansum(values * weight_array, axis=1)
    weight_total = (present * weight_array).sum(axis=1)
    with np.errstate(invalid="ignore", divide="ignore"):
        result = weighted_sum / weight_total
    result[weight_total == 0] = np.nan
    return pd.Series(result, index=frame.index)


def build_basket_temperature(basket, from_date, to_date, forecast_from=None, forecast_to=None):
    """Population-weighted average daily mean temperature across a
    basket's cities (see BASKETS for each city's weight), with the
    forecast window (if given) appended onto each city's series before
    averaging."""
    city_series = {}
    weights = {}
    for city, latitude, longitude, population in basket:
        print(f"  Fetching {city}...", file=sys.stderr, flush=True)
        series = fetch_city_mean_temp(latitude, longitude, from_date, to_date)
        if forecast_from is not None:
            print(f"  Fetching {city} forecast...", file=sys.stderr, flush=True)
            forecast_series = fetch_city_forecast_temp(latitude, longitude, forecast_from, forecast_to)
            series = pd.concat([series, forecast_series])
        city_series[city] = series
        weights[city] = population
    return weighted_average(pd.DataFrame(city_series), weights)


def five_year_normal(full_series, current_dates, years_back=HISTORY_YEARS):
    """For each date in current_dates, the mean of full_series on that
    same month/day in each of the preceding years_back years (whatever
    subset of those years full_series actually has data for - e.g. a
    29 Feb normal averages over however many of those years were leap
    years)."""
    normal = {}
    for current_date in current_dates:
        samples = []
        for years_ago in range(1, years_back + 1):
            past_date = current_date - pd.DateOffset(years=years_ago)
            if past_date in full_series.index:
                samples.append(full_series.loc[past_date])
        normal[current_date] = sum(samples) / len(samples) if samples else float("nan")
    return pd.Series(normal)


def build_weather_dataframe(from_date, to_date, years_back=HISTORY_YEARS, forecast_from=None, forecast_to=None):
    """One Temp/HDD/HDD-5yr-avg/CDD/CDD-5yr-avg column set per region
    basket. Fetches each city once, covering both the current window
    and the years_back years of history the "5-year average" columns
    are built from. If forecast_from/forecast_to are given, each
    basket's temperature series is extended through that window too
    (see build_basket_temperature) - the 5-year normal then covers
    those forecast dates as well, comparing forecast actual against
    the same-calendar-date history, exactly like it already does for
    the historical window."""
    to_date_with_forecast = forecast_to or to_date
    current_dates = pd.date_range(from_date, to_date_with_forecast, freq="D")
    fetch_from = (pd.Timestamp(from_date) - pd.DateOffset(years=years_back)).date()

    columns = {}
    for region, basket in BASKETS.items():
        forecast_note = f" + forecast to {forecast_to.isoformat()}" if forecast_to else ""
        print(
            f"Fetching {region} basket ({fetch_from.isoformat()} to {to_date.isoformat()}){forecast_note}...",
            file=sys.stderr, flush=True,
        )
        full_temp = build_basket_temperature(basket, fetch_from, to_date, forecast_from, forecast_to)
        full_hdd = (DEGREE_DAY_BASE_C - full_temp).clip(lower=0)
        full_cdd = (full_temp - DEGREE_DAY_BASE_C).clip(lower=0)

        columns[f"{region} Temp (C)"] = full_temp.reindex(current_dates)
        columns[f"{region} HDD"] = full_hdd.reindex(current_dates)
        columns[f"{region} HDD (5yr Avg)"] = five_year_normal(full_hdd, current_dates, years_back)
        columns[f"{region} CDD"] = full_cdd.reindex(current_dates)
        columns[f"{region} CDD (5yr Avg)"] = five_year_normal(full_cdd, current_dates, years_back)
    df = pd.DataFrame(columns).sort_index()
    df.index.name = "date"
    return df


def get_excel():
    """Returns (excel, launched_here) - launched_here is True only if
    this call started a dedicated new Excel process (DispatchEx),
    False if it attached to the user's own already-running session
    (GetActiveObject) - see update_weather_chart()'s cleanup, which
    must not Quit() an Excel instance it didn't start (that could
    close the user's own unrelated open files)."""
    try:
        excel = win32.GetActiveObject("Excel.Application")
        launched_here = False
    except Exception:
        excel = win32.DispatchEx("Excel.Application")
        launched_here = True
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    return excel, launched_here


def get_workbook(excel):
    wanted = str(WORKBOOK_PATH.resolve()).lower()
    for workbook in excel.Workbooks:
        try:
            if str(Path(workbook.FullName).resolve()).lower() == wanted:
                return workbook
        except Exception:
            pass

    if not WORKBOOK_PATH.exists():
        # Normally created by one of the other americas/*.py scripts
        # first (MASTER_USA.py runs this one last) - fall back to
        # creating it if run standalone before any of those have.
        workbook = excel.Workbooks.Add()
        workbook.SaveAs(str(WORKBOOK_PATH.resolve()))
        return workbook

    return excel.Workbooks.Open(str(WORKBOOK_PATH.resolve()))


def get_or_create_chart_sheet(workbook):
    for ws in workbook.Worksheets:
        if ws.Name == CHART_SHEET_NAME:
            while ws.ChartObjects().Count:
                ws.ChartObjects(1).Delete()
            ws.Cells.Clear()
            return ws
    ws = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
    ws.Name = CHART_SHEET_NAME
    return ws


EXCEL_EPOCH = pd.Timestamp("1899-12-30")


def to_excel_serial(d):
    """Convert a date/datetime/Timestamp to an Excel serial date
    number. A single Range.Value = <python datetime> assignment works
    fine via win32com's automatic marshaling, but a bulk
    Range.Value = <tuple of tuples> array write does not get the same
    treatment - pywin32 raises "TypeError: must be a PyWinTypes time
    object" for a raw Python datetime/date inside that array, since the
    array-packing path doesn't do the datetime->PyTime conversion a
    scalar property-set does. Writing plain serial numbers sidesteps
    that entirely; NumberFormat below still renders them as dates."""
    return (pd.Timestamp(d) - EXCEL_EPOCH).total_seconds() / 86400


def write_chart_sheet(ws, df, today_date=None):
    headers = ["Date"] + list(df.columns)
    for col, header in enumerate(headers, 1):
        ws.Cells(1, col).Value = header

    rows = [
        (to_excel_serial(index),) + tuple(None if pd.isna(v) else float(v) for v in record)
        for index, record in zip(df.index, df.itertuples(index=False, name=None))
    ]
    end_row = 1 + len(rows)
    ws.Range(
        ws.Cells(2, 1),
        ws.Cells(end_row, len(headers)),
    ).Value = tuple(rows)
    ws.Range(f"A2:A{end_row}").NumberFormat = "dd/mm/yyyy"

    column_letters = {header: get_column_letter(i) for i, header in enumerate(headers, 1)}

    today_lines = {}
    if today_date is not None:
        # "Today" vertical-line helper cells: two rows (Y=0, Y=<per-
        # region chart max>) both at X=today, in their own columns well
        # clear of the main data. Excel positions a series by its own
        # XValues/Values range independently of the shared category
        # range, but still on the SAME time-scale x-axis every series in
        # the chart uses - so a 2-point series here still draws as a
        # correctly-positioned vertical line, not a separate mini-chart.
        today_col = len(headers) + 2
        today_serial = to_excel_serial(today_date)
        ws.Cells(2, today_col).Value = today_serial
        ws.Cells(3, today_col).Value = today_serial
        ws.Range(ws.Cells(2, today_col), ws.Cells(3, today_col)).NumberFormat = "dd/mm/yyyy"
        today_date_range = f"{get_column_letter(today_col)}2:{get_column_letter(today_col)}3"

        for i, region in enumerate(BASKETS):
            region_columns = [
                f"{region} HDD", f"{region} HDD (5yr Avg)",
                f"{region} CDD", f"{region} CDD (5yr Avg)",
            ]
            region_max = df[region_columns].max(numeric_only=True).max()
            chart_max = float(region_max) * 1.1 if pd.notna(region_max) and region_max > 0 else 10.0

            value_col = today_col + 1 + i
            ws.Cells(2, value_col).Value = 0
            ws.Cells(3, value_col).Value = chart_max
            today_value_range = f"{get_column_letter(value_col)}2:{get_column_letter(value_col)}3"
            today_lines[region] = (today_date_range, today_value_range, chart_max)

    return end_row, column_letters, today_lines


def create_line_chart(ws, chart_name, left, top, title, y_axis_title, end_row, column_letters, series_specs, today_line=None):
    """A simple multi-line comparison chart - house style (Aptos fonts,
    white background, Monochromatic Palette 1, legend at bottom), same
    550x270 size as the other native dashboard charts.

    series_specs: list of (column_name, dashed) tuples - "5yr Avg"
    reference series are drawn dashed so they read as the normal
    rather than the actual value. HDD series are blue and CDD series
    orange (HDD_COLOUR / CDD_COLOUR), so the two stand apart; anything
    else keeps Monochromatic Palette 1's shade.

    today_line: optional (date_range, value_range, chart_max) from
    write_chart_sheet() - a dashed grey vertical "Today" reference
    line, drawn as its own 2-point series rather than a fixed-position
    shape, so it always tracks the correct date on the time-scale axis
    regardless of chart size. chart_max also becomes the y-axis's
    explicit MaximumScale, so the line's top exactly matches the
    visible axis top instead of distorting it."""
    for chart_index in range(ws.ChartObjects().Count, 0, -1):
        chart_object = ws.ChartObjects(chart_index)
        if chart_object.Name == chart_name:
            chart_object.Delete()

    chart_object = ws.ChartObjects().Add(left, top, 550, 270)
    chart_object.Name = chart_name

    chart = chart_object.Chart
    chart.ChartType = XL_LINE

    date_range = f"A2:A{end_row}"
    for column_name, dashed in series_specs:
        column_letter = column_letters[column_name]
        series = chart.SeriesCollection().NewSeries()
        series.Name = column_name
        series.XValues = ws.Range(date_range)
        series.Values = ws.Range(f"{column_letter}2:{column_letter}{end_row}")
        series.ChartType = XL_LINE
        series.MarkerStyle = XL_MARKER_NONE
        series.Smooth = False
        series.Format.Line.Weight = 2.0
        if dashed:
            series.Format.Line.DashStyle = 4  # Dash, for the 5-year average reference line

    if today_line is not None:
        today_date_range, today_value_range, _ = today_line
        series = chart.SeriesCollection().NewSeries()
        series.Name = "Today"
        series.XValues = ws.Range(today_date_range)
        series.Values = ws.Range(today_value_range)
        series.ChartType = XL_LINE
        series.MarkerStyle = XL_MARKER_NONE
        series.Smooth = False
        series.Format.Line.Weight = 1.25
        series.Format.Line.DashStyle = 4
        series.Format.Line.ForeColor.RGB = TODAY_LINE_COLOUR

    chart.ChartColor = 5  # Monochromatic Palette 1 (house style) - then HDD/CDD recoloured below

    for index, (column_name, _) in enumerate(series_specs, 1):
        if "HDD" in column_name or "CDD" in column_name:
            line = chart.SeriesCollection(index).Format.Line
            line.Visible = MSO_TRUE
            line.ForeColor.RGB = HDD_COLOUR if "HDD" in column_name else CDD_COLOUR

    chart.HasTitle = True
    chart.ChartTitle.Text = title
    chart.ChartTitle.Font.Name = "Aptos Display"
    chart.ChartTitle.Font.Size = 16
    chart.ChartTitle.Font.Bold = True

    chart.HasLegend = True
    chart.Legend.Position = XL_LEGEND_BOTTOM
    chart.Legend.Font.Name = "Aptos"
    chart.Legend.Font.Size = 9

    x_axis = chart.Axes(XL_CATEGORY, XL_PRIMARY)
    x_axis.CategoryType = XL_TIME_SCALE
    x_axis.BaseUnit = XL_DAYS
    x_axis.TickLabels.NumberFormat = "dd-mmm"
    x_axis.TickLabels.Font.Name = "Aptos"
    x_axis.TickLabels.Font.Size = 9

    y_axis = chart.Axes(XL_VALUE, XL_PRIMARY)
    y_axis.HasTitle = True
    y_axis.AxisTitle.Text = y_axis_title
    y_axis.AxisTitle.Font.Name = "Aptos"
    y_axis.AxisTitle.Font.Size = 11
    y_axis.AxisTitle.Font.Bold = True
    y_axis.MinimumScale = 0
    if today_line is not None:
        # Explicit max, matching the Today line's top exactly - an
        # auto scale wouldn't necessarily land on the same value,
        # which would either clip the line or leave empty space above
        # the real data.
        y_axis.MaximumScaleIsAuto = False
        y_axis.MaximumScale = today_line[2]
    else:
        y_axis.MaximumScaleIsAuto = True
    y_axis.TickLabels.NumberFormat = "0"
    y_axis.TickLabels.Font.Name = "Aptos"
    y_axis.TickLabels.Font.Size = 9

    chart.ChartArea.Format.Fill.Visible = MSO_TRUE
    chart.ChartArea.Format.Fill.ForeColor.RGB = WHITE
    chart.ChartArea.Format.Line.Visible = MSO_FALSE
    chart.PlotArea.Format.Fill.Visible = MSO_TRUE
    chart.PlotArea.Format.Fill.ForeColor.RGB = WHITE
    chart.PlotArea.Format.Line.Visible = MSO_FALSE


def update_weather_chart(df, includes_forecast):
    global win32, pythoncom
    print("Connecting to Excel (pywin32's first run can be slow while it builds its COM cache)...", file=sys.stderr, flush=True)
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    excel = None
    launched_here = False
    workbook = None
    try:
        excel, launched_here = get_excel()
        workbook = get_workbook(excel)
        ws = get_or_create_chart_sheet(workbook)

        # Only draw the "Today" line when the chart actually extends
        # to today - the historical-only fallback (forecast fetch
        # failed) ends at TO_DATE (ARCHIVE_LAG_DAYS behind today), so
        # today would fall outside the plotted range and the line
        # would either not render or distort the axis.
        today_date = date.today() if includes_forecast else None
        end_row, column_letters, today_lines = write_chart_sheet(ws, df, today_date)

        chart_top = 20
        title_suffix = " (actual + forecast)" if includes_forecast else ""
        for region in BASKETS:
            chart_name = "".join(ch for ch in region if ch.isalnum()) + "Chart"
            create_line_chart(
                ws, chart_name, 560, chart_top,
                f"{region}: HDD & CDD vs {HISTORY_YEARS}-Year Average{title_suffix}", "Degree Days",
                end_row, column_letters,
                [
                    (f"{region} HDD", False),
                    (f"{region} HDD (5yr Avg)", True),
                    (f"{region} CDD", False),
                    (f"{region} CDD (5yr Avg)", True),
                ],
                today_line=today_lines.get(region),
            )
            chart_top += 290

        excel.Calculation = XL_CALCULATION_AUTOMATIC
        excel.CalculateFull()
        workbook.Save()
        print(f"Weather charts added to '{CHART_SHEET_NAME}' in {WORKBOOK_PATH.name}", file=sys.stderr, flush=True)
    finally:
        if workbook is not None:
            try:
                # Already Save()d above - close here so the file isn't
                # left locked "open for editing" against a hidden
                # background Excel process once this script exits.
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            excel.ScreenUpdating = True
            excel.EnableEvents = True
            excel.DisplayAlerts = True
            if launched_here:
                # Only quit an Excel instance THIS script started - an
                # attached-to instance may be the user's own session
                # with other files open, which must not be closed.
                try:
                    excel.Quit()
                except Exception:
                    pass
        pythoncom.CoUninitialize()


def main():
    print(
        f"Window: {FROM_DATE.isoformat()} to {TO_DATE.isoformat()} "
        f"(+ forecast to {FORECAST_TO_DATE.isoformat()})",
        file=sys.stderr, flush=True,
    )

    includes_forecast = True
    try:
        df = build_weather_dataframe(
            FROM_DATE, TO_DATE, forecast_from=FORECAST_FROM_DATE, forecast_to=FORECAST_TO_DATE,
        )
    except Exception as exc:
        print(
            f"Forecast fetch failed ({type(exc).__name__}: {exc}) - "
            "continuing with the historical window only.",
            file=sys.stderr, flush=True,
        )
        includes_forecast = False
        df = build_weather_dataframe(FROM_DATE, TO_DATE)

    if df.empty:
        print("No data returned.", file=sys.stderr, flush=True)
        sys.exit(1)

    print(df.tail())
    sys.stdout.flush()

    update_weather_chart(df, includes_forecast)


if __name__ == "__main__":
    main()
