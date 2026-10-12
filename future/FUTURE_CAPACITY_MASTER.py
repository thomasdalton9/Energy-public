"""
MASTER FUTURE CAPACITY workbook -> output/Data and Chart Outputs/Master Outputs/future_capacity_master.xlsx

Combines the three section workbooks (read-only, no downloads):
  europe_future_power_capacity.xlsx       (europe/future/EUROPE_POWER_PIPELINE.py)
  europe_future_gas_infrastructure.xlsx   (europe/future/EUROPE_GAS_INFRA.py)
  global_lng_liquefaction_capacity.xlsx   (global_lng/GLOBAL_LIQUEFACTION.py)

Sheets: All projects, Power by year, Power targets, Regas, Regas by country, Pipelines & storage, Upstream,
Global LNG supply, Global LNG by country, Europe gas headroom, Coverage & gaps, Sources, chart sheets, Units.

Owner rules: raw sourced data only; TARGET rows are allowed, labelled and NEVER summed with projects; no back-tests;
nothing guessed (no ramp-up assumption is carried over from the LNG workbook).
"""
import os
import shutil
import subprocess
import sys
import tempfile

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_charts  # noqa: E402
import xlsx_notes  # noqa: E402

DATA = os.path.join(ROOT, "output", "Data and Chart Outputs")
OUT = os.path.join(DATA, "Master Outputs", "future_capacity_master.xlsx")
PNG_DIR = os.path.join(ROOT, "output", "PNG Charts")
F_POWER = os.path.join(DATA, "europe_future_power_capacity.xlsx")
F_GAS = os.path.join(DATA, "europe_future_gas_infrastructure.xlsx")
F_LNG = os.path.join(DATA, "global_lng_liquefaction_capacity.xlsx")

SCHEMA = ["region", "country", "asset_type", "project", "operator", "fuel_or_type", "capacity_value", "capacity_unit",
          "capacity_mw", "capacity_bcfd", "status", "expected_online", "online_date_basis", "first_seen", "last_seen",
          "source_publisher", "source_document", "source_url", "as_of_date", "evidence", "notes"]
EXTRAS = ["sum_flag", "overlaps_with", "dedup_key", "row_id", "source_key"]
SECTIONS = {"Europe power": F_POWER, "Europe gas": F_GAS, "Global LNG": F_LNG}
YEARS = list(range(2025, 2036))
FIRM = ["UNDER_CONSTRUCTION", "FID"]
INCL = FIRM + ["PERMITTED", "PLANNED"]
FUELS = ["Nuclear", "Gas", "Coal", "Oil", "Wind_Onshore", "Wind_Offshore", "Solar", "Hydro", "Bioenergy",
         "PumpedStorage", "Storage", "Other"]
LABEL = {"Wind_Onshore": "Wind onshore", "Wind_Offshore": "Wind offshore", "PumpedStorage": "Pumped storage",
         "Storage": "Batteries"}
MTPA_TO_BCFD = 1.36 * 35.3147 / 365.25   # DERIVED, as in the liquefaction workbook
PRE, POST, UND = "Before 2025 (date passed)", "2036 or later", "Date not published"
MONTHS = pd.date_range("2025-01-01", "2030-12-01", freq="MS")


def log(*a):
    print(*a, flush=True)


def read(path):
    return pd.read_excel(path, sheet_name=None)


def year_of(v):
    s = str(v)
    return int(s[:4]) if len(s) >= 4 and s[:4].isdigit() else None


# ---------------------------------------------------------------------------------------------------- all projects
def all_projects(sec):
    parts = []
    for name, d in sec.items():
        p = d["Projects"].copy()
        p.insert(0, "section", name)
        for c in SCHEMA + EXTRAS:
            if c not in p.columns:
                p[c] = pd.NA
        parts.append(p[["section"] + SCHEMA + EXTRAS])
    df = pd.concat(parts, ignore_index=True)
    for c in ("expected_online",):
        df[c] = df[c].map(lambda v: "" if pd.isna(v) else str(v))
    return df


# ---------------------------------------------------------------------------------------------------- power by year
def bucket(e):
    y = year_of(e)
    if y is None:
        return UND
    if y < 2025:
        return PRE
    if y > 2035:
        return POST
    return str(y)


COLS = [PRE] + [str(y) for y in YEARS] + [POST, UND]


def power_block(p, statuses, block):
    d = p[(p["sum_flag"] == "Y") & p["status"].isin(statuses)].copy()
    d["capacity_mw"] = d["capacity_mw"].astype(float)
    d["bucket"] = d["expected_online"].map(bucket)
    g = d.pivot_table(index=["country", "fuel_or_type"], columns="bucket", values="capacity_mw", aggfunc="sum",
                      fill_value=0.0)
    for c in COLS:
        if c not in g.columns:
            g[c] = 0.0
    g = g[COLS]
    g["Total, all dates (MW)"] = g.sum(axis=1)
    g["Net 2026-2030 dated (MW)"] = g[[str(y) for y in range(2026, 2031)]].sum(axis=1)
    g["rows"] = d.groupby(["country", "fuel_or_type"]).size()
    g = g.reset_index()
    tot = g.groupby("fuel_or_type")[COLS + ["Total, all dates (MW)", "Net 2026-2030 dated (MW)", "rows"]].sum().reset_index()
    tot.insert(0, "country", "EUROPE TOTAL")
    g = pd.concat([tot, g], ignore_index=True)
    g.insert(0, "block", block)
    return g.round(1)


def power_by_year(p):
    a = power_block(p, FIRM, "(a) firm: UNDER_CONSTRUCTION + FID")
    b = power_block(p, INCL, "(b) firm + PERMITTED + PLANNED")
    q = power_block(p[p["status"] == "QUEUE"].assign(status="QUEUE"), ["QUEUE"], "(c) QUEUE (grid-connection queue; shown separately, not additions)")
    return pd.concat([a, b, q], ignore_index=True), a, b, q


def power_targets(p):
    t = p[p["status"] == "TARGET"].copy()
    t["capacity_mw"] = t["capacity_mw"].astype(float)
    t["year"] = t["expected_online"].map(year_of)
    g = t.pivot_table(index=["country", "fuel_or_type", "source_document"], columns="year", values="capacity_mw",
                      aggfunc="sum")
    g.columns = [str(int(c)) for c in g.columns]
    g = g.reset_index()
    g.insert(0, "label", "TARGET: ERAA national estimate of installed capacity in the target year (MW stock, not additions); never summed with projects")
    return g


# ---------------------------------------------------------------------------------------------------- gas
def regas(gas):
    m = gas["Regas monthly"].copy()
    m["month"] = pd.to_datetime(m["month"].astype(str))
    m = m[m["month"].isin(MONTHS)].reset_index(drop=True)
    return m


def regas_by_country(gas, m):
    r = gas["Regas by month"].copy()
    mc = [c for c in r.columns if str(c)[:2] == "20" and pd.Timestamp(str(c) + "-01") in set(MONTHS)]
    g = r.groupby(["country", "group", "scenario"])[mc].sum().reset_index()
    # tie-out to the gas workbook's own totals (reported, never forced)
    for scen, col in (("committed", "Europe committed Bcf/d"), ("incl.", "Europe incl. planned Bcf/d")):
        s = g[g["scenario"].str.startswith(scen)][mc].sum()
        diff = (s.values - m[col].values)
        log(f"regas tie-out {scen}: max abs difference {abs(diff).max():.4f} Bcf/d")
    return g


def gas_assets(gas, types):
    p = gas["Projects"]
    d = p[p["asset_type"].isin(types)].copy()
    d["expected_online"] = d["expected_online"].map(lambda v: "" if pd.isna(v) else str(v))
    d["label"] = d["status"].map(lambda s: "TARGET (not summed)" if s == "TARGET" else "")
    return d


# ---------------------------------------------------------------------------------------------------- LNG
GROUPS = ["United States", "Qatar", "Australia", "Canada", "Mexico", "Russia", "Africa", "Other"]



def lng_ramp_columns(path):
    """The liquefaction workbook's own 'with ramp' columns (live formulas without cached values): its file is opened
    by LibreOffice in a temp copy so the workbook's formulas are evaluated, nothing is recomputed here.
    Returns (n_months, DataFrame indexed by month) or (None, None) when LibreOffice is unavailable."""
    from openpyxl import load_workbook
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        log("ramp columns skipped: LibreOffice not found")
        return None, None
    with tempfile.TemporaryDirectory() as d:
        src = os.path.join(d, "in.xlsx")
        shutil.copy(path, src)
        try:
            subprocess.run([soffice, "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir",
                            os.path.join(d, "out"), src], check=True, capture_output=True, timeout=240)
            ws = load_workbook(os.path.join(d, "out", "in.xlsx"), data_only=True)["Capacity by month"]
        except Exception as e:  # noqa: BLE001
            log("ramp columns skipped:", e)
            return None, None
    hdr = {str(c.value): c.column for c in ws[1]}
    ca = hdr["Operating + UC + FID, with ramp (mtpa, ASSUMPTION)"]
    cb = hdr["Incl. dated planned, with ramp (mtpa, ASSUMPTION)"]
    cn = hdr["Ramp months (ASSUMPTION, editable)"]
    n = ws.cell(2, cn).value
    rows = {}
    for r in range(2, ws.max_row + 1):
        m = ws.cell(r, 1).value
        if m is not None and pd.Timestamp(m) in set(MONTHS):
            rows[pd.Timestamp(m)] = (ws.cell(r, ca).value, ws.cell(r, cb).value)
    df = pd.DataFrame.from_dict(rows, orient="index")
    if df.isna().all().all() or not isinstance(n, (int, float)):
        log("ramp columns skipped: no values read")
        return None, None
    return int(n), df.astype(float)


def lng_supply(lng):
    cm = lng["Capacity by month"].copy()
    cn = lng["Chart - Nameplate"].copy()
    cm["Month"] = pd.to_datetime(cm["Month"])
    cn["Month"] = pd.to_datetime(cn["Month"])
    cm = cm[cm["Month"].isin(MONTHS)].set_index("Month")
    cn = cn[cn["Month"].isin(MONTHS)].set_index("Month")
    out = cn.copy()
    out.columns = [f"{c} (mtpa)" if c != "Offline (damaged / shut-in)" else "Offline: damaged / shut-in (mtpa)" for c in out.columns]
    keep = ["Operating (mtpa)", "Under construction, by expected month (cum.)", "FID, by expected month (cum.)",
            "Total sourced nameplate = operating + UC + FID", "Planned / permitted with a date (cum.)",
            "Total incl. dated planned", "Undated under construction (memo, constant)", "Undated FID (memo, constant)",
            "Undated planned / permitted (memo, constant)"]
    rename = {"Operating (mtpa)": "TOTAL operating (mtpa)",
              "Under construction, by expected month (cum.)": "TOTAL under construction, dated (mtpa, cum.)",
              "FID, by expected month (cum.)": "TOTAL FID, dated (mtpa, cum.)",
              "Total sourced nameplate = operating + UC + FID": "TOTAL sourced nameplate = operating + UC + FID (mtpa)",
              "Planned / permitted with a date (cum.)": "TOTAL planned / permitted, dated (mtpa, cum.)",
              "Total incl. dated planned": "TOTAL incl. dated planned (mtpa)",
              "Undated under construction (memo, constant)": "MEMO undated under construction (mtpa, not in totals)",
              "Undated FID (memo, constant)": "MEMO undated FID (mtpa, not in totals)",
              "Undated planned / permitted (memo, constant)": "MEMO undated planned / permitted (mtpa, not in totals)"}
    out = pd.concat([cm[keep].rename(columns=rename), out], axis=1)
    out["TOTAL operating (Bcf/d, DERIVED)"] = out["TOTAL operating (mtpa)"] * MTPA_TO_BCFD
    out["TOTAL sourced nameplate (Bcf/d, DERIVED)"] = out["TOTAL sourced nameplate = operating + UC + FID (mtpa)"] * MTPA_TO_BCFD
    out["TOTAL incl. dated planned (Bcf/d, DERIVED)"] = out["TOTAL incl. dated planned (mtpa)"] * MTPA_TO_BCFD
    out["MEMO undated UC + FID + planned (Bcf/d, DERIVED, not in totals)"] = (
        out["MEMO undated under construction (mtpa, not in totals)"] + out["MEMO undated FID (mtpa, not in totals)"]
        + out["MEMO undated planned / permitted (mtpa, not in totals)"]) * MTPA_TO_BCFD
    # tie-out with the workbook's own derived column
    ref = cm["Total sourced nameplate (Bcf/d, DERIVED)"]
    log("LNG Bcf/d tie-out max abs diff:", float((out["TOTAL sourced nameplate (Bcf/d, DERIVED)"] - ref).abs().max()))
    n, rp = lng_ramp_columns(F_LNG)
    if rp is not None:
        # separate ASSUMPTION columns at the end; never part of the sourced series or any total above
        out[f"ASSUMPTION: with {n}-month ramp, operating + UC + FID (mtpa)"] = rp[0]
        out[f"ASSUMPTION: with {n}-month ramp, incl. dated planned (mtpa)"] = rp[1]
        log("ramp columns added, months =", n, "Dec 2030:", rp.loc[pd.Timestamp("2030-12-01")].round(1).tolist())
    return out.round(4)


def headroom(m, ls):
    h = pd.DataFrame(index=MONTHS)
    mm = m.set_index("month")
    h["Europe regas ceiling, committed (Bcf/d)"] = mm["Europe committed Bcf/d"]
    h["Europe regas ceiling, incl. dated planned (Bcf/d)"] = mm["Europe incl. planned Bcf/d"]
    h["Global liquefaction, operating (Bcf/d, DERIVED)"] = ls["TOTAL operating (Bcf/d, DERIVED)"]
    h["Global liquefaction nameplate, sourced = operating + UC + FID (Bcf/d, DERIVED)"] = ls["TOTAL sourced nameplate (Bcf/d, DERIVED)"]
    h["Europe committed regas as % of global sourced nameplate"] = (
        100 * h["Europe regas ceiling, committed (Bcf/d)"] / h["Global liquefaction nameplate, sourced = operating + UC + FID (Bcf/d, DERIVED)"])
    h["Global sourced nameplate less Europe committed regas (Bcf/d)"] = (
        h["Global liquefaction nameplate, sourced = operating + UC + FID (Bcf/d, DERIVED)"] - h["Europe regas ceiling, committed (Bcf/d)"])
    h.index.name = "month"
    return h.round(4)


# ---------------------------------------------------------------------------------------------------- coverage, sources
def coverage(sec):
    rows = []
    t = sec["Europe power"]["Tried"]
    for r in t.itertuples(index=False):
        rows.append(("Europe power", r[0], r[1], r[2], r[3]))
    t = sec["Europe gas"]["Tried"]
    for r in t.itertuples(index=False):
        rows.append(("Europe gas", r[0], r[1], r[2], r[3]))
    t = sec["Global LNG"]["Tried"]
    for r in t.itertuples(index=False):
        rows.append(("Global LNG", r[0], r[1], r[2], ""))
    return pd.DataFrame(rows, columns=["section", "what", "where / url", "result", "checked"])


def sources(sec):
    rows = []
    for r in sec["Europe power"]["Sources"].itertuples(index=False):
        rows.append(("Europe power", r[0], r[1], r[2], r[3], "", ""))
    for r in sec["Europe gas"]["Sources"].itertuples(index=False):
        rows.append(("Europe gas", r[1], r[0], r[2], r[3], "", ""))
    for r in sec["Global LNG"]["Sources"].itertuples(index=False):
        rows.append(("Global LNG", r[0], r[2], r[3], r[5], r[1], r[4]))
    return pd.DataFrame(rows, columns=["section", "publisher", "document", "url", "used for", "role", "as of"])


def as_of_table(sec):
    rows = []
    for name, d in sec.items():
        p = d["Projects"]
        rows.append((name, str(p["last_seen"].dropna().astype(str).max()), len(p)))
    return rows


# ---------------------------------------------------------------------------------------------------- charts
def colour_for(col):
    base = col.replace(" (retirements)", "")
    key = {v: k for k, v in LABEL.items()}.get(base, base)
    c = xlsx_charts.fuel_colour(key)
    if key in ("Wind_Offshore", "PumpedStorage"):
        c = xlsx_charts._lighten(c, 0.35)
    if col.endswith("(retirements)"):
        c = xlsx_charts._lighten(c, 0.5)
    return c


def chart_table(block):
    """GW additions (positive) and retirements (negative) by fuel; rows = date buckets incl. 'Date not published'."""
    b = block[block["country"] == "EUROPE TOTAL"].set_index("fuel_or_type")[COLS]
    add = pd.DataFrame(index=COLS)
    for f in FUELS:
        if f in b.index:
            # additions and retirements need the sign split per row, so use the by-country rows
            pass
    return b


def split_chart_table(p, statuses):
    d = p[(p["sum_flag"] == "Y") & p["status"].isin(statuses)].copy()
    d["mw"] = d["capacity_mw"].astype(float)
    d["bucket"] = d["expected_online"].map(bucket)
    d["fuel"] = d["fuel_or_type"].map(lambda f: f if f in FUELS else "Other")
    t = pd.DataFrame(0.0, index=COLS, columns=[LABEL.get(f, f) for f in FUELS] + [LABEL.get(f, f) + " (retirements)" for f in FUELS])
    for r in d.itertuples():
        col = LABEL.get(r.fuel, r.fuel) + ("" if r.mw > 0 else " (retirements)")
        t.at[r.bucket, col] += r.mw / 1000.0
    t = t.loc[:, t.abs().sum() > 0]
    t.index.name = "Year / date bucket"
    return t.round(3)


def add_power_chart(ws, t, title, anchor_row=1):
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    from openpyxl.utils import get_column_letter
    ws.cell(row=anchor_row, column=1, value=t.index.name)
    for j, c in enumerate(t.columns):
        ws.cell(row=anchor_row, column=2 + j, value=c)
    for i, (lab, row) in enumerate(t.iterrows(), start=1):
        ws.cell(row=anchor_row + i, column=1, value=str(lab))
        for j, v in enumerate(row.values):
            ws.cell(row=anchor_row + i, column=2 + j, value=float(v))
    ws.column_dimensions["A"].width = 26
    n = anchor_row + len(t)
    ch = BarChart()
    ch.type, ch.grouping, ch.overlap, ch.gapWidth = "col", "stacked", 100, 40
    ch.add_data(Reference(ws, min_col=2, max_col=1 + t.shape[1], min_row=anchor_row, max_row=n), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=anchor_row + 1, max_row=n))
    for s, c in zip(ch.series, t.columns):
        s.graphicalProperties = GraphicalProperties(solidFill=colour_for(c), ln=LineProperties(noFill=True))
    ch.title = title
    ch.y_axis.title = "GW"
    ch.x_axis.tickLblPos = "low"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.legend.position = "b"
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.plot_area.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.width, ch.height = 28, 13
    xlsx_charts.tidy_layout(ch)
    ws.add_chart(ch, f"{get_column_letter(t.shape[1] + 4)}{anchor_row}")


def add_monthly_chart(wb, name, df, title, y_title, kind, n_bars, source_line):
    ws = wb.create_sheet(name)
    df = df.copy()
    df.index = pd.to_datetime(df.index)
    ws.cell(row=1, column=1, value=source_line)
    xlsx_charts.write_table(ws, df, "%Y-%m", start_row=2, start_col=1)
    ch = xlsx_charts.build_chart(ws, df, n_bars, title, y_title, kind, "%Y-%m", start_row=2, start_col=1)
    ws.add_chart(ch, f"{chr(ord('A') + min(df.shape[1] + 2, 20))}3")
    return ws


def build_charts(path, p, m, ls, h):
    from openpyxl import load_workbook
    wb = load_workbook(path)
    firm_t = split_chart_table(p, FIRM)
    incl_t = split_chart_table(p, INCL)
    src = ("Sources: MaStR, Bundesnetzagentur, DESNZ REPD + CfD results, NESO TEC, official documents; sum_flag = Y rows; "
           "QUEUE and TARGET rows excluded; 'Date not published' = firm/planned rows with no expected date in the source")
    ws = wb.create_sheet("Chart - Power firm")
    add_power_chart(ws, firm_t, "Europe power: net capacity additions by year, firm (under construction + FID), incl. date not published (GW)", 2)
    ws.cell(row=1, column=1, value=src)
    ws = wb.create_sheet("Chart - Power incl planned")
    add_power_chart(ws, incl_t, "Europe power: net capacity additions by year, firm + permitted + planned, incl. date not published (GW)", 2)
    ws.cell(row=1, column=1, value=src)

    rg = pd.DataFrame(index=MONTHS)
    for g in ("EU27", "United Kingdom", "Turkey", "Balkans", "other Europe"):
        c = f"{g} committed Bcf/d"
        if c in m.columns:
            rg[f"{g}, committed"] = m.set_index("month")[c]
    rg = rg.loc[:, rg.sum() > 0]
    n = rg.shape[1]
    rg["Europe incl. dated planned"] = m.set_index("month")["Europe incl. planned Bcf/d"]
    add_monthly_chart(wb, "Chart - Regas", rg, "Europe LNG regasification ceiling by month, committed by group and incl. dated planned (Bcf/d)",
                      "Bcf/d", "stacked_area", n, "Source: GIE ALSI declared send-out, GIE LNG Map, ENTSOG TYNDP 2024, operator documents (europe_future_gas_infrastructure.xlsx); ceiling, not send-out")

    status = pd.DataFrame({"Operating": ls["TOTAL operating (mtpa)"],
                           "Under construction (dated)": ls["TOTAL under construction, dated (mtpa, cum.)"],
                           "FID (dated)": ls["TOTAL FID, dated (mtpa, cum.)"],
                           "Planned / permitted (dated)": ls["TOTAL planned / permitted, dated (mtpa, cum.)"]})
    add_monthly_chart(wb, "Chart - LNG by status", status, "Global LNG liquefaction nameplate by month and status (mtpa)", "mtpa",
                      "stacked_area", 4, "Sources: EIA, DISR, QatarEnergy, operator documents (global_lng_liquefaction_capacity.xlsx); undated projects excluded (memo columns on 'Global LNG supply')")
    grp = pd.DataFrame(index=MONTHS)
    for g in GROUPS:
        grp[g] = sum(ls[f"{g} - {k} (mtpa)"] for k in ("operating", "UC + FID"))
    add_monthly_chart(wb, "Chart - LNG by group", grp, "Global LNG liquefaction nameplate by month and country group, operating + UC + FID (mtpa)",
                      "mtpa", "stacked_area", len(GROUPS), "Sources: EIA, DISR, QatarEnergy, operator documents; undated projects excluded")
    hh = h[["Europe regas ceiling, committed (Bcf/d)", "Europe regas ceiling, incl. dated planned (Bcf/d)",
            "Global liquefaction, operating (Bcf/d, DERIVED)",
            "Global liquefaction nameplate, sourced = operating + UC + FID (Bcf/d, DERIVED)"]].copy()
    hh.columns = ["Europe regas ceiling, committed", "Europe regas ceiling, incl. dated planned",
                  "Global liquefaction, operating", "Global liquefaction nameplate, sourced (operating + UC + FID)"]
    add_monthly_chart(wb, "Chart - Headroom", hh, "Europe regas ceiling vs global liquefaction nameplate (Bcf/d; context only, not a forecast)",
                      "Bcf/d", "line", 0, "Context only. Regas = send-out CEILING, liquefaction = NAMEPLATE; neither is a flow or a forecast.")
    xlsx_charts.save_atomic(wb, path)
    return firm_t, incl_t, rg, status, grp, hh


def pngs(firm_t, incl_t, rg, status, grp, hh):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    os.makedirs(PNG_DIR, exist_ok=True)
    out = []

    def finish(fig, ax, name, foot):
        for s in ("top", "right"):
            ax.spines[s].set_visible(False)
        fig.text(0.01, 0.005, foot, fontsize=7, color="#555")
        fig.tight_layout(rect=(0, 0.03, 1, 1))
        p = os.path.join(PNG_DIR, f"future_capacity_master__{name}.png")
        fig.savefig(p, dpi=130)
        plt.close(fig)
        out.append(p)

    for t, nm, ttl in ((firm_t, "Power_firm", "firm (under construction + FID)"), (incl_t, "Power_incl_planned", "firm + permitted + planned")):
        fig, ax = plt.subplots(figsize=(11, 5.6))
        x = list(range(len(t)))
        pos = [0.0] * len(t)
        neg = [0.0] * len(t)
        for c in t.columns:
            v = t[c].tolist()
            if c.endswith("(retirements)"):
                ax.bar(x, v, bottom=neg, color="#" + colour_for(c), width=0.7, label=c)
                neg = [a + b for a, b in zip(neg, v)]
            else:
                ax.bar(x, v, bottom=pos, color="#" + colour_for(c), width=0.7, label=c)
                pos = [a + b for a, b in zip(pos, v)]
        ax.axhline(0, color="#444", lw=0.6)
        ax.set_xticks(x)
        ax.set_xticklabels([str(i).replace(" (date passed)", "\n(date passed)").replace("Date not published", "Date not\npublished") for i in t.index], fontsize=8, rotation=0)
        ax.set_ylabel("GW", fontsize=9)
        ax.set_title(f"Europe power: net capacity additions by year, {ttl}, incl. date not published (GW)", fontsize=12, fontweight="bold", loc="left")
        ax.legend(ncol=5, fontsize=7.5, frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.12))
        finish(fig, ax, nm, "Sources: MaStR, Bundesnetzagentur, DESNZ REPD + CfD, NESO TEC, official documents; sum_flag = Y; QUEUE and TARGET excluded; retirements below zero.")

    def monthly(df, nm, ttl, ylab, foot, stacked_n=None, lines=()):
        fig, ax = plt.subplots(figsize=(11, 5.4))
        x = df.index
        cols = list(df.columns)
        bars = [c for c in cols if c not in lines]
        cs = [("#" + xlsx_charts.fuel_colour(c, i)) for i, c in enumerate(bars)]
        if bars and stacked_n != 0:
            ax.stackplot(x, [df[c].fillna(0) for c in bars], labels=bars, colors=cs, alpha=0.95)
        dash = ["-", "--", ":", "-."]
        for i, c in enumerate(lines):
            col = "#" + xlsx_charts.PALETTE[i % len(xlsx_charts.PALETTE)] if len(lines) > 1 else "#252525"
            ax.plot(x, df[c], color=col, ls=dash[i % 4] if len(lines) == 1 or "planned" in c else "-", lw=2.0, label=c)
        ax.set_ylabel(ylab, fontsize=9)
        ax.set_title(ttl, fontsize=12, fontweight="bold", loc="left")
        ax.legend(ncol=3, fontsize=7.5, frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.1))
        ax.margins(x=0)
        finish(fig, ax, nm, foot)

    rb = [c for c in rg.columns if c != "Europe incl. dated planned"]
    monthly(rg, "Regas_ceiling", "Europe LNG regasification ceiling by month (Bcf/d)", "Bcf/d",
            "Sources: GIE ALSI / LNG Map, ENTSOG TYNDP 2024, operator documents. Ceiling, not send-out. Committed = operating + under construction + FID; line adds dated planned.",
            lines=("Europe incl. dated planned",))
    monthly(status, "Global_LNG_by_status", "Global LNG liquefaction nameplate by month and status (mtpa)", "mtpa",
            "Sources: EIA, DISR, QatarEnergy, operator documents. Undated projects excluded (memo columns in the workbook). Nameplate, not output.")
    monthly(grp, "Global_LNG_by_group", "Global LNG liquefaction nameplate by month, operating + UC + FID (mtpa)", "mtpa",
            "Sources: EIA, DISR, QatarEnergy, operator documents. Undated projects excluded. Nameplate, not output.")
    monthly(hh, "Headroom", "Europe regas ceiling vs global liquefaction nameplate (Bcf/d) - context, not a forecast", "Bcf/d",
            "Regas = send-out ceiling; liquefaction = nameplate (mtpa x 1.36 bcm/mtpa -> Bcf/d, DERIVED). Neither is a flow; no forecast or balance is implied.",
            stacked_n=0, lines=tuple(hh.columns))
    return out


# ---------------------------------------------------------------------------------------------------- main
def main():
    for f in (F_POWER, F_GAS, F_LNG):
        if not os.path.exists(f):
            sys.exit(f"missing section workbook {f}")
    sec = {k: read(v) for k, v in SECTIONS.items()}
    allp = all_projects(sec)
    pw = allp[allp["section"] == "Europe power"]
    by_year, a, b, q = power_by_year(pw)
    targets = power_targets(pw)
    m = regas(sec["Europe gas"])
    rc = regas_by_country(sec["Europe gas"], m)
    pipes = gas_assets(sec["Europe gas"], ["gas_pipeline", "gas_storage"])
    up = gas_assets(sec["Europe gas"], ["gas_field"])
    gsum = sec["Europe gas"]["Summary"]
    ls = lng_supply(sec["Global LNG"])
    bc = sec["Global LNG"]["By country"]
    h = headroom(m, ls)
    cov = coverage(sec)
    src = sources(sec)
    asof = as_of_table(sec)

    # notes
    n = []
    T = []

    def sec_(t):
        T.append(t)
        n.extend([t])

    n.append("Master future-capacity workbook: what is planned, under construction or being retired, from three section workbooks read as committed (no downloads). Raw sourced data only; nothing guessed.")
    n.append("")
    sec_("Sections and as-of dates")
    for name, last, rows in asof:
        n.append(f"{name}: {rows} project rows, latest last_seen {last}. Source workbook rebuilt on its own 1st/15th schedule (gas 03:17, LNG 09:19, power 09:52 UTC); this master runs after all three.")
    n.append("")
    sec_("Status ladder")
    n.append("UNDER_CONSTRUCTION > FID > PERMITTED > PLANNED > QUEUE (grid-connection queue totals, not projects with funding) ; TARGET (national estimates, policy or TYNDP less-advanced). Also kept as in the sections: SUSPENDED, CANCELLED, COMMISSIONED_RECENT (online after 1 Jan 2025, kept for reconciliation) and, for liquefaction, OPERATING and CLOSED.")
    n.append("'firm' = UNDER_CONSTRUCTION + FID. 'firm + planned' = firm + PERMITTED + PLANNED. QUEUE is shown on its own block and never added. Closures: statutory or notified closures are FID, other announced closures PLANNED (negative MW).")
    n.append("")
    sec_("Sum rules")
    n.append("TARGET rows (ERAA national estimates of installed capacity, TYNDP less-advanced gas projects) are allowed, labelled and NEVER summed with projects: they are on 'Power targets' (power) or flagged 'TARGET (not summed)' on 'Pipelines & storage'. ERAA values are stock in the target year, not additions.")
    n.append("'Power by year' adds only rows with sum_flag = Y (a duplicate of a more official row has sum_flag = N). Retirements are negative. Plants and national aggregates (power_aggregate) both count where sum_flag = Y.")
    n.append("'Date not published' = rows with no expected_online in the source (e.g. UK REPD under-construction offshore wind and batteries have no commissioning date). They are in the totals but not in any year column; 'Before 2025' = dates already past; '2036 or later' beyond the table.")
    n.append("Gas and liquefaction rows carry no sum_flag: regas, pipelines and liquefaction are totalled only by the section workbooks' own monthly rules, copied here unchanged.")
    n.append("")
    sec_("Conversions")
    n.append("Power MW as published. Gas: bcm/y x 35.3147 / 365.25 = Bcf/d; GWh/d / 11.2 GWh per mcm x 0.0353147; GIE map bcm are normal m3, not corrected. LNG: mtpa x 1.36 bcm/y per mtpa x 35.3147 / 365.25 = 0.1315 Bcf/d (DERIVED, factor stated in the liquefaction workbook).")
    n.append("")
    sec_("Regas ('Regas', 'Regas by country')")
    n.append("Copied from the gas workbook's 'Regas monthly' / 'Regas by month' (2025-2030). Ceiling per terminal = ALSI declared send-out where published, else GIE LNG Map nominal. committed = operating + under construction + FID; incl. planned adds dated PLANNED / PERMITTED. A stated year counts from January. 'Regas by country' sums the gas workbook's terminal rows by country; the tie-out to its Europe totals is logged each run.")
    n.append("")
    sec_("Global LNG supply")
    n.append("Copied from the liquefaction workbook's 'Capacity by month' / 'Chart - Nameplate' (2025-2030): operating, dated under construction, dated FID, dated planned (cumulative from the placed start month), offline (damaged / shut-in) shown separately. Undated projects are MEMO columns, not in the monthly totals. The last two columns, headed 'ASSUMPTION: with <n>-month ramp', are the liquefaction workbook's own 'with ramp' columns (each addition after the run month rises linearly over an assumed number of months, owner-approved ASSUMPTION; values read from that workbook, not recomputed); they are never added into the sourced series, totals, charts or headroom.")
    n.append("Nameplate is not output: US plants run above baseload nameplate, Qatar and the UAE export through the Strait of Hormuz, Russian volumes are sanctioned.")
    n.append("")
    sec_("Europe gas headroom")
    n.append("CONTEXT ONLY, not a forecast and not a balance: the Europe regas send-out ceiling (committed) set beside global liquefaction nameplate (sourced = operating + UC + FID), month by month, in Bcf/d. Neither series is a flow; the ratio says how large Europe's import ceiling is against world nameplate, nothing about what will be traded.")
    n.append("")
    sec_("Rules")
    n.append("Raw sourced data; every number comes from a document opened in full by its section script; search snippets are never used. No back-tests. Nothing is guessed: what was tried and not sourced is on 'Coverage & gaps'.")
    n.append("Charts (native, no borders, 12 pt bold titles, power colours from fuel_colour) are on the 'Chart - ...' sheets and as PNGs future_capacity_master__<name>.png in output/PNG Charts.")

    sheets = {
        "All projects": allp.set_index("section"),
        "Power by year": by_year.set_index("block"),
        "Power targets": targets.set_index("label"),
        "Regas": m.set_index("month"),
        "Regas by country": rc.set_index("country"),
        "Pipelines & storage": pipes.set_index("label"),
        "Gas summary (section)": gsum.set_index(gsum.columns[0]),
        "Upstream": up.set_index("label"),
        "Global LNG supply": ls,
        "Global LNG by country": bc.set_index("country"),
        "Europe gas headroom": h,
        "Coverage & gaps": cov.set_index("section"),
        "Sources": src.set_index("section"),
    }
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    xlsx_notes.write_workbook(OUT, sheets, n, T)
    firm_t, incl_t, rg, status, grp, hh = build_charts(OUT, pw, m, ls, h)
    out = pngs(firm_t, incl_t, rg, status, grp, hh)

    def tot(block, label):
        t = block[block["country"] == "EUROPE TOTAL"].set_index("fuel_or_type")
        y = [str(i) for i in range(2026, 2031)]
        s = pd.DataFrame({"dated 2026-30 (MW)": t[y].sum(axis=1), "undated (MW)": t[UND], "all dates (MW)": t["Total, all dates (MW)"]})
        log(f"\n{label}\n", s.round(0).to_string(), "\n", s.sum().round(0).to_dict())
    tot(a, "FIRM"); tot(b, "FIRM + PLANNED")
    log("\nRegas ceiling Bcf/d:", m.set_index("month").loc[["2025-01-01", "2026-10-01", "2030-12-01"] if False else pd.to_datetime(["2026-10-01", "2030-12-01"]), ["Europe committed Bcf/d", "Europe incl. planned Bcf/d"]].to_dict())
    log("Global LNG:", ls.loc[pd.to_datetime(["2026-10-01", "2030-12-01"]), ["TOTAL sourced nameplate = operating + UC + FID (mtpa)", "TOTAL sourced nameplate (Bcf/d, DERIVED)", "TOTAL operating (mtpa)"]].to_dict())
    log("PNGs:", out)
    log("wrote", OUT)


if __name__ == "__main__":
    main()
