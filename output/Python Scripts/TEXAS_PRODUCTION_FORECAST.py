"""
Texas natural gas production forecast to Dec 2028 (Bcf/d, monthly) with a Permian takeaway cap, base and a
'takeaway delayed' scenario, plus the implied net interstate outflow. EIA is blocked from the Claude sandbox, so this
runs in GitHub Actions (.github/workflows/texas_production_forecast.yml, 1st and 15th; EIA_API_KEY secret).

  History       Texas marketed gas production, 'Balance' sheet of texas_gas_monthly.xlsx (read only; TEXAS_GAS.py owns it).
  Forecast base EIA STEO (API v2 route steo) regional MARKETED production: NGMPPM Permian, NGMPEF Eagle Ford,
                NGMPHA Haynesville (Bcf/d, monthly, STEO horizon = end of next year, i.e. Dec 2027). STEO's own series
                are already marketed, so no dry->marketed conversion is applied (the 16.4% last-12-month extraction-loss
                share TEXAS_GAS.py uses is shown on the Assumptions tab only to give the dry equivalent).
                Texas share: Permian x PERMIAN_TX_SHARE (rest is New Mexico), Haynesville x HAYNESVILLE_TX_SHARE (rest is
                Louisiana), Eagle Ford 100%. 'Other Texas' (Barnett, Panhandle, definitional differences) = the mean gap
                between Texas history and the shares x STEO over the months both exist, held flat (can be negative).
                2028 is beyond STEO: each month = same month of 2027 x (STEO 2027 / 2026 annual average) per region.
  Takeaway cap  Permian production = min(STEO Permian, existing takeaway + new pipelines in service + local demand),
                from the editable 'Takeaway' table. 'Delayed' shifts every new pipeline's in-service month by DELAY_MONTHS.
  Net outflow   production - Texas consumption (published sectors) - Mexico pipeline exports - LNG exports. For forecast
                months consumption = same month a year earlier, Mexico = last 12 months mean, LNG = column 'LNG exports
                (INPUT)' = last-12-month mean held flat; the LNG forecast can overwrite that column.
  Edits         The Assumptions and Takeaway tabs of the committed workbook are read back on every run, so edit them in the
                workbook (or in the repo file) and the next run (or a manual dispatch) recomputes the forecast.
  Fallback      If STEO cannot be fetched, the last STEO copy saved on the 'STEO raw' tab is used; if there is none,
                the forecast is the last-12-month trend (each month = same month a year earlier x last-12/previous-12 growth)
                and the Units tab says so.

Usage: python3 TEXAS_PRODUCTION_FORECAST.py [--out "output/Data and Chart Outputs/texas_production_forecast.xlsx"]
"""
import argparse
import os
import sys

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
DEFAULT_OUT = os.path.join(OUT_DIR, "texas_production_forecast.xlsx")
TEXAS_XLSX = os.path.join(OUT_DIR, "texas_gas_monthly.xlsx")
END = "2028-12-01"
STEO_IDS = {"Permian": "NGMPPM", "Eagle Ford": "NGMPEF", "Haynesville": "NGMPHA"}

DEFAULT_PARAMS = [
    ("Permian Texas share of STEO Permian", 0.85, "share", "STEO Permian = Texas + southeast New Mexico (NM about 13-15%). Assumption."),
    ("Haynesville Texas share of STEO Haynesville", 0.30, "share", "STEO Haynesville = East Texas + Louisiana. Assumption (unverified)."),
    ("Eagle Ford Texas share", 1.00, "share", "Eagle Ford is in Texas."),
    ("Local Permian demand (power, industrial, Mexico-bound not included)", 3.0, "Bcf/d", "Gas consumed in the basin, adds to takeaway in the cap. Unverified."),
    ("Takeaway delay in the delayed scenario", 6, "months", "Every pipeline with status not 'in service' slips by this many months."),
]
DEFAULT_PIPES = [
    ("Existing Permian takeaway (aggregate, Waha-bound + Midland/Eagle Ford routes incl. Matterhorn)", 26.5, "2025-01", "In service",
     "Calibration: with local demand it brackets STEO Permian marketed in 2026. Total is unverified."),
    ("Matterhorn Express (Kinder/WhiteWater/EnLink/Devon, Waha-Katy)", 2.5, "2024-10", "In service", "Already inside the existing aggregate above (shown for reference, not added)."),
    ("Gulf Coast Express expansion (Kinder Morgan)", 0.57, "2026-07", "Under construction", "unverified - company announcement, date from memory"),
    ("Blackcomb Pipeline (WhiteWater/Targa/MPLX/Enterprise)", 2.5, "2026-10", "Under construction", "unverified - company announcement, date from memory"),
    ("Hugh Brinson Pipeline (Energy Transfer, phase 1)", 1.5, "2026-12", "Under construction", "unverified - company announcement, date from memory"),
    ("Hugh Brinson Pipeline (Energy Transfer, phase 2 expansion)", 0.7, "2027-07", "Under construction", "unverified - company announcement, date from memory"),
    ("Eiger Express (WhiteWater/Targa/MPLX/Enterprise)", 2.5, "2028-07", "Planned", "unverified - company announcement, date from memory"),
    ("Traverse Pipeline (WhiteWater, Agua Dulce)", 1.75, "2027-10", "Planned", "unverified - company announcement, date from memory"),
]


def steo_fetch():
    key = os.environ.get("EIA_API_KEY")
    if not key:
        raise RuntimeError("EIA_API_KEY not set")
    cols = {}
    for name, sid in STEO_IDS.items():
        r = requests.get("https://api.eia.gov/v2/steo/data/", timeout=(10, 90), params={
            "api_key": key, "frequency": "monthly", "data[0]": "value", "facets[seriesId][]": sid, "start": "2024-01",
            "sort[0][column]": "period", "sort[0][direction]": "asc", "length": 200})
        r.raise_for_status()
        d = r.json()["response"]["data"]
        if not d:
            raise RuntimeError(f"no STEO rows for {sid}")
        cols[name] = pd.Series({pd.Timestamp(x["period"] + "-01"): float(x["value"]) for x in d})
    return pd.DataFrame(cols).sort_index()


def read_prev(path, sheet, **kw):
    try:
        return pd.read_excel(path, sheet_name=sheet, **kw)
    except Exception:  # noqa: BLE001
        return None


def months(a, b):
    return pd.date_range(a, b, freq="MS")


def run(out):
    bal = pd.read_excel(TEXAS_XLSX, sheet_name="Balance")
    bal["Month"] = pd.to_datetime(bal["Month"])
    bal = bal.set_index("Month").sort_index()
    hist = bal["Marketed production"].dropna()
    last = hist.index[-1]
    idx = months("2015-01-01", END)
    fc = [m for m in idx if m > last]

    # editable assumptions (read back from the committed workbook)
    pp = read_prev(out, "Assumptions")
    params = {}
    for lab, val, unit, note in DEFAULT_PARAMS:
        v = val
        if pp is not None and "Value" in pp:
            hit = pp[pp["Assumption"].astype(str).eq(lab)]
            if len(hit) and pd.notna(hit["Value"].iloc[0]):
                v = float(hit["Value"].iloc[0])
        params[lab] = v
    tk = read_prev(out, "Takeaway")
    if tk is None or "Capacity Bcf/d" not in tk:
        tk = pd.DataFrame(DEFAULT_PIPES, columns=["Pipeline", "Capacity Bcf/d", "In-service month", "Status", "Source / verification"])
    tk = tk.dropna(subset=["Pipeline"]).copy()
    tk["In-service month"] = tk["In-service month"].astype(str).str[:7]
    sh_p, sh_h, sh_e, local, delay = (params[k[0]] for k in DEFAULT_PARAMS)
    delay = int(delay)

    # --- STEO (live, else saved copy, else trend)
    method = "EIA STEO (live)"
    try:
        steo = steo_fetch()
    except Exception as e:  # noqa: BLE001
        print("STEO fetch failed:", e)
        raw = read_prev(out, "STEO raw")
        if raw is not None and len(raw):
            steo = raw.set_index(pd.to_datetime(raw["Month"]))[list(STEO_IDS)]
            method = "EIA STEO (saved copy from an earlier run - live fetch failed)"
        else:
            steo, method = None, "last-12-month trend (STEO unavailable)"
    print("method:", method)

    df = pd.DataFrame(index=idx)
    df["Type"] = ["History" if m <= last else "Forecast" for m in idx]
    df["Texas marketed production (history)"] = hist.reindex(idx)

    if steo is not None:
        # 2028 beyond STEO: same month of 2027 x annual-average growth 2027/2026 per region
        s = steo.reindex(idx)
        for reg in STEO_IDS:
            a26 = s.loc["2026-01-01":"2026-12-01", reg].mean()
            a27 = s.loc["2027-01-01":"2027-12-01", reg].mean()
            g = a27 / a26
            for m in months("2028-01-01", END):
                s.loc[m, reg] = s.loc[m - pd.DateOffset(years=1), reg] * g
        df["STEO Permian marketed (Bcf/d)"] = s["Permian"]
        df["STEO Eagle Ford marketed (Bcf/d)"] = s["Eagle Ford"]
        df["STEO Haynesville marketed (Bcf/d)"] = s["Haynesville"]
        df["STEO source"] = ["STEO" if m <= pd.Timestamp("2027-12-01") else "extension of 2027 at STEO 2027/2026 growth" for m in idx]
        # takeaway capacity by scenario
        add = tk[~tk["Pipeline"].str.contains("Existing", case=False) & tk["Status"].astype(str).str.lower().ne("in service")]
        base_ex = float(tk[tk["Pipeline"].str.contains("Existing", case=False)]["Capacity Bcf/d"].sum())

        def cap(delay_m):
            c = pd.Series(base_ex + local, index=idx, dtype=float)
            for _, r in add.iterrows():
                start = pd.Timestamp(r["In-service month"] + "-01") + pd.DateOffset(months=delay_m)
                c[c.index >= start] += float(r["Capacity Bcf/d"])
            return c
        df["Permian takeaway + local demand, base (Bcf/d)"] = cap(0)
        df["Permian takeaway + local demand, delayed (Bcf/d)"] = cap(delay)
        for sc, col in (("base", "Permian takeaway + local demand, base (Bcf/d)"), ("delayed", "Permian takeaway + local demand, delayed (Bcf/d)")):
            df[f"Permian marketed capped, {sc}"] = pd.concat([df["STEO Permian marketed (Bcf/d)"], df[col]], axis=1).min(axis=1)
        df["Permian capped by takeaway (base)"] = df["STEO Permian marketed (Bcf/d)"] > df["Permian takeaway + local demand, base (Bcf/d)"]
        ov = df.index[(df.index <= last) & df["STEO Permian marketed (Bcf/d)"].notna()][-6:]
        comp = (sh_p * df["STEO Permian marketed (Bcf/d)"] + sh_e * df["STEO Eagle Ford marketed (Bcf/d)"]
                + sh_h * df["STEO Haynesville marketed (Bcf/d)"])
        other = float((hist.reindex(ov) - comp.reindex(ov)).mean())
        df["Other Texas / calibration (held flat)"] = other
        for sc in ("base", "delayed"):
            tx = (sh_p * df[f"Permian marketed capped, {sc}"] + sh_e * df["STEO Eagle Ford marketed (Bcf/d)"]
                  + sh_h * df["STEO Haynesville marketed (Bcf/d)"] + other)
            df[f"Texas production forecast, {sc}"] = tx.where(df.index > last)
    else:
        tail = hist.iloc[-12:].sum() / hist.iloc[-24:-12].sum()
        full = hist.reindex(idx)
        for m in fc:
            full[m] = full[m - pd.DateOffset(years=1)] * tail
        for sc in ("base", "delayed"):
            df[f"Texas production forecast, {sc}"] = full.where(df.index > last)
        other = float("nan")

    # --- implied net interstate outflow
    cons = bal["Consumption (published sectors)"].reindex(idx)
    mex = bal["Pipeline exports to Mexico"].reindex(idx)
    lng = bal["LNG exports"].reindex(idx)
    mex_flat, lng_flat = mex.dropna().iloc[-12:].mean(), lng.dropna().iloc[-12:].mean()
    cons_f, mex_f, lng_f = cons.copy(), mex.copy(), lng.copy()
    for m in idx:
        if m > last:
            cons_f[m] = cons_f[m - pd.DateOffset(years=1)]
            mex_f[m] = mex_flat
            lng_f[m] = lng_flat
    df["Texas consumption, published sectors (forecast = same month prior year)"] = cons_f
    df["Pipeline exports to Mexico (forecast = last-12-month mean)"] = mex_f
    df["LNG exports (INPUT: forecast = last-12-month mean held flat; replace with the LNG forecast)"] = lng_f
    lng_col = "LNG exports (INPUT: forecast = last-12-month mean held flat; replace with the LNG forecast)"
    for sc in ("base", "delayed"):
        prod = df["Texas marketed production (history)"].fillna(df[f"Texas production forecast, {sc}"])
        df[f"Implied net interstate outflow, {sc}"] = (prod - df["Texas consumption, published sectors (forecast = same month prior year)"]
                                                        - df["Pipeline exports to Mexico (forecast = last-12-month mean)"] - df[lng_col])
    df.index.name = "Month"

    # --- workbook
    assum = pd.DataFrame([(a, params[a], u, n) for a, _, u, n in DEFAULT_PARAMS], columns=["Assumption", "Value", "Unit", "Note"])
    loss = bal["Dry production"].dropna()
    ratio = 1 - (bal["Dry production"] / bal["Marketed production"]).dropna().iloc[-12:].mean()
    steo_raw = (steo.reset_index().rename(columns={"index": "Month"}) if steo is not None else pd.DataFrame(columns=["Month"]))
    summary = []
    for y in (2026, 2027, 2028):
        d = df[df.index == f"{y}-12-01"]
        summary.append({"Month": f"Dec {y}", "Base Bcf/d": d["Texas production forecast, base"].iloc[0],
                        "Delayed Bcf/d": d["Texas production forecast, delayed"].iloc[0],
                        "Base net outflow Bcf/d": d["Implied net interstate outflow, base"].iloc[0],
                        "Delayed net outflow Bcf/d": d["Implied net interstate outflow, delayed"].iloc[0]})
    notes = [
        "Notes", "", "UNITS",
        "All flows are Bcf/d (billion cubic feet per day), monthly averages. Texas marketed production, consumption, Mexico and LNG history come from texas_gas_monthly.xlsx (EIA).",
        "", "METHOD",
        f"Forecast base: {method}. Series: NGMPPM (Permian), NGMPEF (Eagle Ford), NGMPHA (Haynesville), regional marketed production. STEO ends Dec 2027; 2028 extends 2027 by the STEO 2027/2026 annual growth per region (stated, not an EIA forecast).",
        f"Texas = {sh_p:.0%} Permian + {sh_e:.0%} Eagle Ford + {sh_h:.0%} Haynesville + 'Other Texas' calibration ({other:+.2f} Bcf/d, mean gap to Texas history over the last 6 overlapping months, held flat). Shares are editable assumptions.",
        f"STEO regional series are already marketed, so no dry-to-marketed conversion is applied; the last-12-month extraction-loss share used in TEXAS_GAS.py is {ratio:.1%} of marketed (dry equivalent = marketed x {1 - ratio:.3f}).",
        "Permian cap: production = min(STEO Permian, existing takeaway + new pipelines in service + local demand) from the Takeaway tab. Delayed scenario shifts every pipeline not yet in service by the delay on the Assumptions tab.",
        "Takeaway table: only Matterhorn (in service) is firm; every other capacity and date is UNVERIFIED (company announcements, from memory) and the existing-takeaway aggregate is a calibration. Edit the tab and re-run to update.",
        "Net interstate outflow = marketed production - consumption (published sectors; EIA withholds Texas lease/plant and pipeline fuel) - Mexico pipeline exports - LNG exports. Forecast months: consumption = same month prior year, Mexico and LNG = last-12-month means.",
        "LNG column is an INPUT: it holds the last-12-month level flat for now and is meant to be replaced by the LNG forecast.",
        f"Latest history month: {last:%b/%y}. Not an EIA or company forecast of Texas; a transparent scenario tool.",
    ]
    sheets = {"Summary": pd.DataFrame(summary), "Assumptions": assum, "Takeaway": tk, "Forecast": df.reset_index(),
              "STEO raw": steo_raw}
    xlsx_notes.write_workbook(out, sheets, notes, ["UNITS", "METHOD"])
    print(pd.DataFrame(summary).to_string())
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    a = ap.parse_args()
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    run(a.out)
