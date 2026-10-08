"""
Shared helper: add an AGSI-style "water year" sheet with a native Excel
chart to a workbook - the same look as the gas storage charts:
  - shaded band = min-max of the 5 complete water years before this one
  - dashed grey = 5-year average
  - mid blue    = previous water year
  - dark blue   = current water year (to the latest day)
Water year runs Oct-Sep (storage/level charts in this repo use Oct-Sep).

Usage (after the workbook has been written):
    import water_year_chart
    water_year_chart.add_water_year_chart(path, series, "Rhine at Kaub", "cm")
where `series` is a pandas Series of daily values indexed by date.
"""
import math
import os

import numpy as np
import pandas as pd

import xlsx_charts
from openpyxl import load_workbook
from openpyxl.chart import AreaChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties

SHEET = "Water year chart"
BAND_FILL = "AFC6D9"
AVG_LINE = "6E6E6E"
PREV_LINE = "2A78D6"   # previous year: mid blue (owner, Oct 2026; was orange)
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


def week_year_table(series):
    """Same layout as water_year_table but one row per WEEK (1-52) of the calendar year, for weekly series (US petroleum
    stocks): shaded band = min-max of the 5 complete calendar years before the current one, 5-year average, last year and
    current year (to the latest week). Week = (day of year - 1) // 7 + 1, week 53 folded into 52."""
    s = series.dropna().sort_index()
    s.index = pd.DatetimeIndex(s.index)
    last = s.index.max()
    cur = int(last.year)
    wk = pd.Series(np.minimum(52, (s.index.dayofyear - 1) // 7 + 1), index=s.index)
    df = pd.DataFrame({"v": s.values, "year": s.index.year, "wk": wk.values}).groupby(["wk", "year"]).v.last().unstack("year")
    df = df.reindex(range(1, 53))
    hist = [y for y in range(cur - 5, cur) if y in df.columns]
    h = df[hist]
    t = pd.DataFrame({
        "Day": [f"Week {w}" for w in range(1, 53)],
        "5Y min": h.min(axis=1).round(2).values,
        "5Y range (max-min)": (h.max(axis=1) - h.min(axis=1)).round(2).values,
        "5Y max": h.max(axis=1).round(2).values,
        "5Y average": h.mean(axis=1).round(2).values,
        str(cur - 1): df[cur - 1].round(2).values if (cur - 1) in df.columns else np.nan,
        str(cur): df[cur].round(2).values if cur in df.columns else np.nan,
    })
    meta = {"hist": f"{hist[0]}-{hist[-1]}" if hist else "", "last": last.date(), "weekly": True}
    return t, meta


def _nice_step(span, ticks=7):
    """A 1/2/2.5/5 x 10^k step giving about `ticks` gridlines over span."""
    raw = span / ticks
    mag = 10 ** math.floor(math.log10(raw))
    return next(m * mag for m in (1, 2, 2.5, 5, 10) if m * mag >= raw)


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


def build_chart(ws, table, meta, title, unit, width=26, height=12, gridlines=True, inner=None, short_title=False,
                y_decimals=None):
    """AGSI-style chart over a table written by write_table on ws (the chart can be placed on any sheet)."""
    n = len(table) + 1
    # Axis labels only on the 1st of each month (a category axis can only skip evenly, so label the other
    # days blank and show every label): written next to the table, column H.
    label_col = table.shape[1] + 1
    ws.cell(row=1, column=label_col, value="Axis label")
    for i, day in enumerate(table["Day"], start=2):
        if meta.get("weekly"):   # label weeks 1, 5, 9 ... (every 4th)
            keep = (i - 2) % 4 == 0
            day = str(day).replace("Week ", "Wk ")   # short, so the labels stay horizontal and are never cut off
        else:
            keep = str(day).startswith("01-")
        ws.cell(row=i, column=label_col, value=day if keep else None)
    cats = Reference(ws, min_col=label_col, min_row=2, max_row=n)
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

    if short_title:   # dashboards: keep the title to one line so it never runs into the plot
        area.title = f"{title} (to {pd.Timestamp(meta['last']):%d/%m/%y})"
    else:
        basis = "calendar-year weeks" if meta.get("weekly") else "water year (Oct-Sep)"
        area.title = f"{title} - {basis}, data to {meta['last']}"
    area.y_axis.title = unit
    area.x_axis.tickLblSkip = 1
    area.x_axis.tickMarkSkip = 1
    area.x_axis.tickLblPos = "low"   # dates below the plot, not on the zero line
    # same angled date labels as the monthly power/gas charts; the short week labels stay horizontal (angled
    # "Week n" labels were truncated to "We..." and ran into the legend)
    xlsx_charts.rotated_labels(area.x_axis, 0 if meta.get("weekly") else -45)
    area.x_axis.delete = False
    area.y_axis.delete = False
    if "%" in str(unit):   # percent full: fixed 0-100 scale, gridline/label every 20%
        area.y_axis.scaling.min, area.y_axis.scaling.max, area.y_axis.majorUnit = 0.0, 100.0, 20.0
    else:
        vals = pd.concat([table["5Y min"], table["5Y max"], table.iloc[:, 5], table.iloc[:, 6]]).dropna()
        lo, hi = float(vals.min()), float(vals.max())
        step = _nice_step((hi - lo) or 1.0)   # round axis ends and gridline step (Excel ticks start at the min)
        area.y_axis.scaling.min = math.floor(lo / step) * step
        area.y_axis.scaling.max = math.ceil(hi / step) * step
        if area.y_axis.scaling.max - area.y_axis.scaling.min < step * 2:
            area.y_axis.scaling.max = area.y_axis.scaling.min + step * 2
        area.y_axis.majorUnit = step
    if y_decimals == 0 and "%" not in str(unit):   # whole-number labels: whole-number axis ends too
        area.y_axis.scaling.min = float(math.floor(area.y_axis.scaling.min))
        area.y_axis.scaling.max = float(math.ceil(area.y_axis.scaling.max))
    if y_decimals is not None:
        area.y_axis.number_format = "0" if y_decimals == 0 else "0." + "0" * y_decimals
        area.y_axis.numFmt.sourceLinked = False
    lines.y_axis.scaling.min = area.y_axis.scaling.min
    lines.y_axis.scaling.max = area.y_axis.scaling.max
    lines.y_axis.majorUnit = area.y_axis.majorUnit
    lines.y_axis.delete = True   # one visible y axis (shared id); the area chart draws it
    area.legend.position = "b"
    # no borders: chart frame and plot area
    area.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    area.plot_area.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    area.height, area.width = height, width
    if not gridlines:
        lines.y_axis.majorGridlines = None
    area += lines
    return xlsx_charts.tidy_layout(area, gridlines, inner)

def add_water_year_chart(path, series, title, unit, sheet_name=SHEET, y_decimals=None, wb=None, weekly=False):
    """sheet_name: pass a different name to put several water-year charts in one workbook.
    wb: an open workbook to add the sheet to (the caller saves it). openpyxl drops the formatting of charts it
    reads back in (axis min/max and step, axis titles), so several charts must go into one workbook in a single
    load/save - add_charts.py does that."""
    table, meta = week_year_table(series) if weekly else water_year_table(series)
    own = wb is None
    if own:
        wb = load_workbook(path)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name, 1 if len(wb.sheetnames) > 1 else None)
    write_table(ws, table)
    area = build_chart(ws, table, meta, title, unit, y_decimals=y_decimals)
    ws.add_chart(area, "I2")
    if not own:
        return table

    root, ext = os.path.splitext(path)
    tmp = f"{root}.tmp{os.getpid()}{ext}"
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return table
