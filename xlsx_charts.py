"""
Shared helper: add a chart sheet with a native Excel chart to a workbook
that has already been written (e.g. by xlsx_notes.write_workbook).

Every xlsx this repo pulls should carry a chart of its data (see
CLAUDE.md). The chart sheet holds its own copy of the plotted table, so
it works whatever layout the data sheets use (wide, long, with an
index column...): pass a wide DataFrame indexed by date, one column per
series.

    import xlsx_charts
    xlsx_charts.add_chart_sheet(path, wide_df, "Bolivia gas demand by sector",
                                "million m3/day", kind="stacked_bar")

kind: "line", "stacked_area", "stacked_bar".
For storage/level series use water_year_chart.add_water_year_chart instead.
"""
import os

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import AreaChart, BarChart, LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties

# Validated categorical order (dataviz palette): fixed order, never cycled.
PALETTE = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "4A3AA7", "E34948"]
OTHER_GREY = "9A9A9A"


def _fold_to_palette(df):
    """More series than palette slots -> keep the largest, fold the rest into 'Other'."""
    if df.shape[1] <= len(PALETTE):
        return df
    order = df.abs().mean().sort_values(ascending=False).index
    keep = list(order[: len(PALETTE) - 1])
    out = df[keep].copy()
    out["Other"] = df.drop(columns=keep).sum(axis=1, min_count=1)
    return out


def add_chart_sheet(path, df, title, y_title, kind="line", sheet_name="Chart", date_format="%Y-%m",
                    width=28, height=13):
    df = df.copy()
    df.index = pd.to_datetime(df.index)
    df = df.sort_index()
    df = df.apply(pd.to_numeric, errors="coerce")
    df = df.loc[:, df.notna().any()]
    if df.empty:
        return
    df = _fold_to_palette(df)

    wb = load_workbook(path)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name, 1 if len(wb.sheetnames) > 1 else None)
    ws.append(["Date"] + [str(c) for c in df.columns])
    for ts, row in df.iterrows():
        ws.append([ts.strftime(date_format)] + [None if pd.isna(v) else float(v) for v in row.values])
    n = len(df) + 1
    ws.column_dimensions["A"].width = 11

    if kind == "stacked_bar":
        chart = BarChart()
        chart.type = "col"
        chart.grouping = "stacked"
        chart.overlap = 100
        chart.gapWidth = 30
    elif kind == "stacked_area":
        chart = AreaChart()
        chart.grouping = "stacked"
    else:
        chart = LineChart()
    chart.add_data(Reference(ws, min_col=2, max_col=df.shape[1] + 1, min_row=1, max_row=n), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=n))
    for i, s in enumerate(chart.series):
        colour = OTHER_GREY if str(df.columns[i]) == "Other" else PALETTE[i % len(PALETTE)]
        if kind == "line":
            s.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=colour, w=22225))
            s.smooth = False
            s.marker.symbol = "none"
        else:
            s.graphicalProperties = GraphicalProperties(solidFill=colour, ln=LineProperties(noFill=True))
    chart.title = f"{title} (to {df.index.max().strftime(date_format)})"
    chart.y_axis.title = y_title
    chart.x_axis.tickLblPos = "low"
    step = max(1, len(df) // 12)
    chart.x_axis.tickLblSkip = step
    chart.x_axis.tickMarkSkip = step
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.legend.position = "b"
    chart.width, chart.height = width, height
    ws.add_chart(chart, f"{chr(ord('A') + min(df.shape[1] + 2, 20))}2")

    root, ext = os.path.splitext(path)
    tmp = f"{root}.tmp{os.getpid()}{ext}"
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
