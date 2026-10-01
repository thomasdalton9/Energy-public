"""
Shared helper: add an AGSI-style "water year" sheet with a native Excel
chart to a workbook - the same look as the gas storage charts:
  - shaded band = min-max of the 5 complete water years before this one
  - dashed grey = 5-year average
  - orange      = previous water year
  - dark blue   = current water year (to the latest day)
Water year runs Oct-Sep (storage/level charts in this repo use Oct-Sep).

Usage (after the workbook has been written):
    import water_year_chart
    water_year_chart.add_water_year_chart(path, series, "Rhine at Kaub", "cm")
where `series` is a pandas Series of daily values indexed by date.
"""
import os

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import AreaChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties

SHEET = "Water year chart"
BAND_FILL = "AFC6D9"
AVG_LINE = "6E6E6E"
PREV_LINE = "EB6834"
CURR_LINE = "0B3A66"


def _day_of_water_year(idx):
    start = pd.DatetimeIndex([pd.Timestamp(y if m >= 10 else y - 1, 10, 1) for y, m in zip(idx.year, idx.month)])
    return (idx - start).days + 1


def water_year_table(series):
    """One row per day of the water year (1-365): label, band, average, previous and current year."""
    s = series.dropna().sort_index()
    s.index = pd.DatetimeIndex(s.index)
    last = s.index.max()
    wy_start = pd.Timestamp(last.year if last.month >= 10 else last.year - 1, 10, 1)
    prev_start = wy_start - pd.DateOffset(years=1)
    hist_start = wy_start - pd.DateOffset(years=5)

    def by_day(part):
        d = part.groupby(_day_of_water_year(part.index)).mean()
        return d.reindex(range(1, 366))   # Feb 29 (day 152 in leap years) folds into its neighbours' days

    hist = s[(s.index >= hist_start) & (s.index < wy_start)]
    hday = pd.Series(hist.values, index=_day_of_water_year(hist.index))
    hday = hday[hday.index <= 365]
    band = hday.groupby(level=0).agg(["min", "max", "mean"]).reindex(range(1, 366)).interpolate().ffill().bfill()
    labels = [(pd.Timestamp(2001, 10, 1) + pd.Timedelta(days=d - 1)).strftime("%d-%b") for d in range(1, 366)]
    t = pd.DataFrame({
        "Day": labels,
        "5Y min": band["min"].round(2).values,
        "5Y range (max-min)": (band["max"] - band["min"]).round(2).values,
        "5Y max": band["max"].round(2).values,
        "5Y average": band["mean"].round(2).values,
        f"WY {prev_start.year}/{str(prev_start.year + 1)[2:]}": by_day(s[(s.index >= prev_start) & (s.index < wy_start)]).round(2).values,
        f"WY {wy_start.year}/{str(wy_start.year + 1)[2:]}": by_day(s[s.index >= wy_start]).round(2).values,
    })
    meta = {"hist": f"{hist_start.year}/{str(hist_start.year + 1)[2:]}-{prev_start.year}/{str(prev_start.year + 1)[2:]}",
            "last": last.date()}
    return t, meta


def _line(series, colour, width_emu, dash=None):
    series.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=colour, w=width_emu, prstDash=dash))
    series.smooth = False
    series.marker.symbol = "none"



def write_table(ws, table):
    ws.append(list(table.columns))
    for row in table.itertuples(index=False):
        ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in row])
    ws.column_dimensions["A"].width = 9
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 14


def build_chart(ws, table, meta, title, unit, width=26, height=12):
    """AGSI-style chart over a table written by write_table on ws (the chart can be placed on any sheet)."""
    n = len(table) + 1
    cats = Reference(ws, min_col=1, min_row=2, max_row=n)
    # Band: stacked area of (min, max-min) with the min part invisible.
    area = AreaChart()
    area.grouping = "stacked"
    area.add_data(Reference(ws, min_col=2, max_col=3, min_row=1, max_row=n), titles_from_data=True)
    area.set_categories(cats)
    base, rng = area.series
    base.graphicalProperties = GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))
    rng.graphicalProperties = GraphicalProperties(solidFill=BAND_FILL, ln=LineProperties(noFill=True))
    rng.tx = SeriesLabel(v=f"5-year range ({meta['hist']})")
    base.tx = SeriesLabel(v=" ")

    lines = LineChart()
    lines.add_data(Reference(ws, min_col=5, max_col=7, min_row=1, max_row=n), titles_from_data=True)
    lines.set_categories(cats)
    avg, prev, curr = lines.series
    _line(avg, AVG_LINE, 19050, "dash")
    _line(prev, PREV_LINE, 19050)
    _line(curr, CURR_LINE, 38100)

    area.title = f"{title} - water year (Oct-Sep), data to {meta['last']}"
    area.y_axis.title = unit
    area.x_axis.tickLblSkip = 30
    area.x_axis.tickMarkSkip = 30
    area.x_axis.tickLblPos = "low"   # dates below the plot, not on the zero line
    area.x_axis.delete = False
    area.y_axis.delete = False
    vals = pd.concat([table["5Y min"], table["5Y max"], table.iloc[:, 5], table.iloc[:, 6]]).dropna()
    span = float(vals.max() - vals.min()) or 1.0
    area.y_axis.scaling.min = float(vals.min() - 0.05 * span)
    area.y_axis.scaling.max = float(vals.max() + 0.05 * span)
    lines.y_axis.scaling.min = area.y_axis.scaling.min
    lines.y_axis.scaling.max = area.y_axis.scaling.max
    lines.y_axis.delete = True   # one visible y axis (shared id); the area chart draws it
    area.legend.position = "b"
    # no borders: chart frame and plot area
    area.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    area.plot_area.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    area.height, area.width = height, width
    area += lines
    return area

def add_water_year_chart(path, series, title, unit, sheet_name=SHEET):
    """sheet_name: pass a different name to put several water-year charts in one workbook."""
    table, meta = water_year_table(series)
    wb = load_workbook(path)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name, 1 if len(wb.sheetnames) > 1 else None)
    write_table(ws, table)
    area = build_chart(ws, table, meta, title, unit)
    ws.add_chart(area, "I2")

    root, ext = os.path.splitext(path)
    tmp = f"{root}.tmp{os.getpid()}{ext}"
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return table
