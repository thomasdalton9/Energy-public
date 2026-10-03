"""
Capacity factors by generation type for the master workbooks (South & Central America, North America,
Australia + NZ): capacity factor = energy generated / (installed capacity x hours in the period).

Inputs per country, as the masters already build them:
  generation  monthly GWh in the dashboard fuel groups (Hydro, Gas, Wind, Solar, Coal, Nuclear, Other)
  capacity    the country's capacity workbook (sheet "Monthly": date, <Fuel>_MW for Hydro, Gas, Wind, Solar, Coal,
              Nuclear, Oil, Bioenergy, Other; storage columns are not used)

Capacity is a stock: annual rows (year-end values dated 1 January) are placed in December and carried forward
month by month; a month before the first capacity figure takes that figure for up to 12 months (flagged in the
notes) and is otherwise left out. "Other" generation (oil, bioenergy, geothermal, other) is compared with Oil +
Bioenergy + Other capacity. Values up to 110% are kept (output above a net or summer rating is real); above
110% (generation and capacity classify the plant differently) and fuels under 50 MW are left blank.

add_capacity_factor_sheets() writes a "Capacity factors" summary tab (trailing 12 months, country x fuel) and one
regional chart (monthly %, all countries summed) for the master's power dashboard.
"""
import os

import pandas as pd
from openpyxl.styles import Font, PatternFill

import add_charts
import xlsx_charts

FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Other"]
CAP_FOR = {"Hydro": ["Hydro"], "Gas": ["Gas"], "Wind": ["Wind"], "Solar": ["Solar"], "Coal": ["Coal"],
           "Nuclear": ["Nuclear"], "Other": ["Oil", "Bioenergy", "Other"]}
MIN_MW = 50.0
MAX_CF = 110.0   # slightly over 100% is real (output above the net / summer rating in cool months); far above it
                 # means generation and capacity classify the plant differently (e.g. dual-fuel turbines burning
                 # diesel: oil generation, gas capacity) and is left blank
BACKFILL_MONTHS = 12


def monthly_capacity(path):
    """Capacity workbook -> monthly MW by fuel group (FUELS), carried forward; (frame, annual?, first month)."""
    d = add_charts.by_date(add_charts.read(path, "Monthly"), "date")
    d = d.apply(pd.to_numeric, errors="coerce")
    annual = len(d) > 1 and d.index.to_series().diff().median().days > 300
    if annual:
        d.index = d.index + pd.DateOffset(months=11)   # year-end value
    cap = pd.DataFrame({f: d[[f"{c}_MW" for c in CAP_FOR[f] if f"{c}_MW" in d]].sum(axis=1, min_count=1)
                        for f in FUELS}, index=d.index)
    cap.index = cap.index.to_period("M").to_timestamp()
    cap = cap.groupby(level=0).last()
    return cap, annual, cap.index.min()


def capacity_factor(gen_gwh, cap_mw):
    """Monthly GWh x monthly MW -> monthly capacity factor %, with the months/fuels that can't be compared blank."""
    gen = gen_gwh.reindex(columns=FUELS).apply(pd.to_numeric, errors="coerce")
    # a month whose total generation is under half the typical month is a gap in the feed, not plant standing idle
    tot = gen.sum(axis=1, min_count=1)
    gen = gen[tot >= 0.5 * tot.rolling(13, center=True, min_periods=3).median()]
    months = gen.index
    cap = cap_mw.reindex(cap_mw.index.union(months)).sort_index().ffill()
    cap = cap.bfill(limit=BACKFILL_MONTHS).reindex(months)
    hours = pd.Series(months.days_in_month * 24.0, index=months)
    cf = gen * 1000.0 / cap.mul(hours, axis=0) * 100.0
    cf = cf.where(cap >= MIN_MW).where(cf <= MAX_CF).where(cf >= 0)
    return cf.dropna(how="all", axis=1).dropna(how="all"), gen, cap


def trailing(gen, cap, months=12):
    """Energy-weighted capacity factor over the last `months` months both series cover, per fuel (%)."""
    ok = gen.notna() & cap.notna() & (cap >= MIN_MW)
    last = gen.where(ok).dropna(how="all").index
    if not len(last):
        return pd.Series(dtype=float), None
    window = last[-months:]
    g = gen.loc[window].where(ok.loc[window])
    hrs = pd.Series(window.days_in_month * 24.0, index=window)
    c = cap.loc[window].where(ok.loc[window]).mul(hrs, axis=0)
    out = g.sum() * 1000.0 / c.sum() * 100.0
    return out.where((out > 0) & (out <= MAX_CF)), (window.min(), window.max())


def add_capacity_factor_sheets(wb, used, sheet_name_fn, gen_frames, cap_paths, chart_w, chart_h, source_note,
                               region, exclude=None):
    """gen_frames: {country: monthly GWh frame}; cap_paths: {country: capacity workbook path}. Writes a
    "Capacity factors" summary tab (trailing 12 months, country x fuel) and ONE regional chart: monthly capacity
    factor by generation type over all countries (summed generation / summed capacity x hours; a month sums the
    countries reporting it, shown when they hold >= 90% of the region's capacity). Returns (chart, index_row, missing) for the master's power dashboard."""
    missing, summary, gens, caps, country_charts = [], [], {}, {}, []
    for country, gen in gen_frames.items():
        path = cap_paths.get(country)
        if not path or not os.path.exists(path):
            missing.append(f"{country}: no capacity workbook (capacity factor not calculated)")
            continue
        try:
            cap, annual, first = monthly_capacity(path)
            for (ec, fuel), _ in (exclude or {}).items():
                if ec == country and fuel in gen:
                    gen = gen.assign(**{fuel: float("nan")})
            cf, g, c = capacity_factor(gen, cap)
        except Exception as e:  # noqa: BLE001
            missing.append(f"{country}: capacity factor failed ({type(e).__name__}: {e})")
            continue
        if cf.empty:
            missing.append(f"{country}: generation and capacity have no month in common")
            continue
        t12, window = trailing(g, c)
        note = ("annual year-end capacity carried forward" if annual else "monthly capacity") + (
            f"; capacity from {first:%b/%y} applied to earlier months" if not annual and cf.index.min() < first else "") + \
            "".join(f"; {fuel} left blank: {why}" for (ec, fuel), why in (exclude or {}).items() if ec == country)
        summary.append((country, window, t12, note))
        ok = g.notna() & c.notna() & (c >= MIN_MW)
        gens[country], caps[country] = g.where(ok), c.where(ok)
        cdf = cf[cf.index >= "2021-01-01"]
        if not cdf.empty:
            ws = wb.create_sheet(sheet_name_fn(f"{country[:20].rstrip()} CF data", used))
            df, _ = xlsx_charts.prepare(cdf.round(1))
            xlsx_charts.write_table(ws, df)
            ws.cell(row=1, column=df.shape[1] + 4, value=f"Capacity factor = generation / (capacity x hours); {note}")
            title = f"{country} capacity factor by generation type"
            country_charts.append(
                ((xlsx_charts.build_chart(ws, df, df.shape[1], title, "% of installed capacity", "line",
                                          width=chart_w, height=chart_h, gridlines=False,
                                          inner=xlsx_charts.DASHBOARD_INNER), (source_note, None)),
                 (country, title, df.index.max().strftime("%b/%y"), ws.title, source_note, None)))

    chart, row = None, None
    if gens:
        # each month sums the countries that have that month (a lagging feed drops out of both generation and
        # capacity, so the ratio stays comparable); only months where >= 90% of the region's capacity reports
        months = pd.DatetimeIndex(sorted(set.union(*(set(g.dropna(how="all").index) for g in gens.values()))))
        cap_all = sum(c.reindex(c.index.union(months)).ffill().bfill().reindex(months).fillna(0).sum(axis=1)
                      for c in caps.values())   # a country's capacity counts even in months it has no figure
        cap_rep = sum(c.reindex(months).fillna(0).sum(axis=1) for c in caps.values())
        months = months[(cap_rep >= 0.9 * cap_all).values]
        hours = pd.Series(months.days_in_month * 24.0, index=months)
        g_sum = sum(g.reindex(months).fillna(0) for g in gens.values())
        c_sum = sum(c.reindex(months).fillna(0) for c in caps.values())
        reg = g_sum * 1000.0 / c_sum.mul(hours, axis=0) * 100.0
        reg = reg.where(c_sum >= MIN_MW).where((reg > 0) & (reg <= MAX_CF))
        reg = reg[reg.index >= "2021-01-01"].dropna(how="all", axis=1).dropna(how="all")
        if not reg.empty:
            ws = wb.create_sheet(sheet_name_fn(f"{region[:18].rstrip()} CF data", used))
            df, _ = xlsx_charts.prepare(reg.round(1))
            xlsx_charts.write_table(ws, df)
            ws.cell(row=1, column=df.shape[1] + 4, value="Capacity factor = generation / (capacity x hours), countries "
                                                         "reporting each month summed. Countries: " + ", ".join(gens))
            title = f"{region} capacity factor by generation type"

            def build():   # one chart object per dashboard (openpyxl can't place a chart twice)
                return xlsx_charts.build_chart(ws, df, df.shape[1], title, "% of installed capacity", "line",
                                               width=chart_w, height=chart_h, gridlines=False,
                                               inner=xlsx_charts.DASHBOARD_INNER)
            row = (region, title + f" ({len(gens)} countries)", df.index.max().strftime("%b/%y"), ws.title,
                   source_note, None)
            chart = (build(), (source_note, None))
            country_charts.insert(0, ((build(), (source_note, None)), row))   # leads the capacity-factor dashboard

    ws = wb.create_sheet(sheet_name_fn("Capacity factors", used))
    ws["A1"] = "Capacity factor by generation type, % - trailing 12 months (energy-weighted: total generation / " \
               "(capacity x hours) over the window)"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = ("Generation and capacity are the same series as on the power dashboard. 'Other' = oil, bioenergy, "
                "geothermal and other plant. Up to 110% is kept (plants can run above a net / summer rating); blank: "
                "under 50 MW installed, or above 110% (generation and capacity classify the plant differently, e.g. "
                "dual-fuel turbines burning diesel count as oil generation but gas capacity).")
    ws["A2"].font = Font(italic=True, color="6B6B6B")
    head = ["Country", "Window"] + FUELS + ["Capacity basis"]
    for j, h in enumerate(head, start=1):
        cell = ws.cell(row=4, column=j, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color="17365D")
    for i, (country, window, t12, note) in enumerate(summary, start=5):
        ws.cell(row=i, column=1, value=country)
        ws.cell(row=i, column=2, value=f"{window[0]:%b/%y}-{window[1]:%b/%y}" if window else "")
        for j, f in enumerate(FUELS, start=3):
            v = t12.get(f)
            if v is not None and pd.notna(v):
                ws.cell(row=i, column=j, value=round(float(v), 1)).number_format = "0.0"
        ws.cell(row=i, column=len(head), value=note)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions[ws.cell(row=4, column=len(head)).column_letter].width = 60
    return chart, row, missing, country_charts
