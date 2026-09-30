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

Each region is a simple, equally-weighted average across a handful of
major demand-centre cities (not a true population-weighted index):
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

# Simple, equally-weighted proxy baskets - not true population-weighted
# indices. Pick major gas/power demand centres per region.
BASKETS = {
    "USA": [
        ("New York", 40.71, -74.01),
        ("Chicago", 41.85, -87.65),
        ("Houston", 29.76, -95.37),
        ("Los Angeles", 34.05, -118.24),
        ("Atlanta", 33.75, -84.39),
    ],
    "NW Europe": [
        ("London", 51.51, -0.13),
        ("Paris", 48.85, 2.35),
        ("Amsterdam", 52.37, 4.90),
        ("Brussels", 50.85, 4.35),
        ("Frankfurt", 50.11, 8.68),
    ],
    "JKTC": [
        ("Tokyo", 35.68, 139.65),
        ("Seoul", 31.23, 121.47),
        ("Taipei", 25.03, 121.56),
        ("Shanghai", 31.23, 121.47),
    ],
}


def fetch_city_mean_temp(latitude, longitude, from_date, to_date):
    """Daily mean 2m air temperature (Excel)... 
