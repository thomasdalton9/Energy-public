"""Chart specs for the Taiwan workbooks (registered into add_charts.REGISTRY by register()):

  taiwan_esist_monthly.xlsx    Energy Administration monthly statistics (taiwan/TAIWAN_ESIST_MONTHLY.py)
  taiwan_taipower.xlsx         Taipower daily peak load / reserve margin and EMS-unit generation (taiwan/TAIWAN_TAIPOWER.py)
  taiwan_reservoirs_daily.xlsx WRA reservoir storage (taiwan/TAIWAN_WRA_RESERVOIRS.py)
  taiwan_generation_rollup.xlsx 10-minute generation by fuel rolled up to daily / monthly (taiwan/TAIWAN_GEN_ROLLUP.py)
  taiwan_live_daily.xlsx       daily means of Taipower's hourly-polled live snapshot (taiwan/TAIWAN_LIVE_SNAPSHOT.py)

Power types use the fixed names of xlsx_charts.FUEL_COLOURS (Hydro, Gas, Wind, Solar, Coal, Nuclear, Oil, Bioenergy,
Geothermal, Pumped storage).
"""
import os

import pandas as pd

ESIST = "taiwan_esist_monthly.xlsx"
TAIPOWER = "taiwan_taipower.xlsx"
RESERVOIRS = "taiwan_reservoirs_daily.xlsx"
ROLLUP = "taiwan_generation_rollup.xlsx"
LIVE = "taiwan_live_daily.xlsx"
GEN_COLS = {"Hydro": "Renewable Energy - Hydro", "Gas": "Thermal - LNG-Fired", "Wind": "Renewable Energy - Wind",
            "Solar": "Renewable Energy - Solar PV", "Coal": "Thermal - Coal-Fired", "Nuclear": "Nuclear",
            "Oil": "Thermal - Oil-Fired",
            "Other": ("Renewable Energy - Biomass", "Renewable Energy - Waste", "Renewable Energy - Geothermal", "Pumped Storage")}
CAP_ORDER = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Other"]   # Other = biomass, waste, geothermal, pumped storage


def _sheet(path, name, dates):
    try:
        d = pd.read_excel(path, sheet_name=name, index_col=0)
    except Exception:  # noqa: BLE001
        return pd.DataFrame()
    d.index = pd.to_datetime(d.index) if dates else pd.to_datetime(d.index.astype(int).astype(str), format="%Y")
    return d.sort_index()


def _pick(d, mapping, scale=1.0):
    out = {}
    for name, col in mapping.items():
        cols = col if isinstance(col, tuple) else (col,)
        have = [c for c in cols if c in d.columns]
        if have:
            out[name] = d[have].sum(axis=1, min_count=1) * scale
    return pd.DataFrame(out, index=d.index)


def _top(d, n=6, drop=("Total", "Grand Total")):
    """Keep the n largest columns (by mean), fold the rest, and any 'Others' column, into 'Others'."""
    d = d[[c for c in d.columns if c not in drop]].apply(pd.to_numeric, errors="coerce")
    named = [c for c in d.columns if c != "Others"]
    keep = list(d[named].mean().sort_values(ascending=False).index[:n])
    out = d[keep].copy()
    rest = [c for c in d.columns if c not in keep]
    if rest:
        out["Others"] = d[rest].sum(axis=1, min_count=1)
    return out


def _bar(ac, name, df, title, units, kind="stacked_bar", fmt="%Y-%m"):
    df = df.dropna(how="all")
    return ac.spec(name, df, title, units, kind, fmt) if not df.empty else None


def esist_specs(ac):
    def f(p):
        M = lambda t: _sheet(p, f"{t} M", True)    # noqa: E731
        A = lambda t: _sheet(p, f"{t} A", False)   # noqa: E731
        out = []
        gm, ga = M("GEN"), A("GEN")
        out.append(_bar(ac, "Generation", _pick(gm, GEN_COLS), "Taiwan power generation by source (Energy Administration)",
                        "GWh per month"))
        out.append(_bar(ac, "Generation annual", _pick(ga, GEN_COLS), "Taiwan annual power generation by source (Energy Administration)",
                        "GWh per year", fmt="%Y"))
        cm, ca = M("CAP"), A("CAP")
        cap = GEN_COLS
        out.append(_bar(ac, "Capacity", _pick(cm, cap, 0.001)[CAP_ORDER], "Taiwan installed generating capacity by source, end of month (Energy Administration)",
                        "GW installed"))
        out.append(_bar(ac, "Capacity annual", _pick(ca, cap, 0.001)[CAP_ORDER], "Taiwan installed generating capacity by source, end of year (Energy Administration)",
                        "GW installed", fmt="%Y"))
        cs = M("CONS")
        if not cs.empty:
            c = pd.DataFrame({"Industry": cs.get("Industrial Sector"), "Services": cs.get("Service Sector"),
                              "Residential": cs.get("Residential Sector"), "Energy sector own use": cs.get("Energy Sector Own Use"),
                              "Agriculture": cs.get("Agricultural Sector"), "Transport": cs.get("Transport Sector")})
            out.append(_bar(ac, "Consumption", c, "Taiwan electricity consumption by sector (Energy Administration)", "GWh per month"))
        rm = M("REN")
        if not rm.empty:
            r = pd.DataFrame({"Hydro": rm.get("Hydro"), "Solar": rm.get("Solar PV"), "Wind": rm.get("Wind - Subtotal"),
                              "Bioenergy": rm[[c for c in ("Biomass - Subtotal", "Waste") if c in rm]].sum(axis=1, min_count=1),
                              "Geothermal": rm.get("Geothermal")}) / 1000.0
            out.append(_bar(ac, "Renewables", r, "Taiwan renewable power generation by source (Energy Administration)", "GWh per month"))
        gs, gsa = M("GAS"), A("GAS")
        for tag, g, per in (("", gs, "bcm per month"),):
            if g.empty:
                continue
            k = 1e-6   # thousand m3 -> bcm
            sup = pd.DataFrame({"LNG imports": g.get("Natural Gas Supply - Import"),
                                "Domestic production": g.get("Natural Gas Supply - Indigenous Production")}) * k
            out.append(_bar(ac, "Gas supply", sup, "Taiwan natural gas supply: LNG imports and domestic production (Energy Administration)", per))
            use = pd.DataFrame({
                "Power generation": g.get("Transformation Input - Electricity Generation and Cogeneration",
                                          g.get("Transformation Input - Electricity Generation a...")),
                "Industry": g.get("Natural Gas Consumption - Industrial Sector"),
                "Residential": g.get("Natural Gas Consumption - Residential Sector"),
                "Services": g.get("Natural Gas Consumption - Service Sector"),
                "Energy sector own use": g.get("Natural Gas Consumption - Energy Sector Own Use"),
                "Refineries": g.get("Transformation Input - Petroleum Refineries")}) * k
            if use["Power generation"].isna().all():
                pcol = [c for c in g.columns if c.startswith("Transformation Input - Electricity")]
                if pcol:
                    use["Power generation"] = g[pcol[0]] * k
            out.append(_bar(ac, "Gas use", use, "Taiwan natural gas use by sector (Energy Administration)", per))
        if not gsa.empty:
            k = 1e-6
            use = pd.DataFrame({"Power generation": gsa[[c for c in gsa.columns if c.startswith("Transformation Input - Electricity")][0]] * k,
                                "Industry": gsa.get("Natural Gas Consumption - Industrial Sector") * k,
                                "Residential": gsa.get("Natural Gas Consumption - Residential Sector") * k,
                                "Services": gsa.get("Natural Gas Consumption - Service Sector") * k,
                                "Energy sector own use": gsa.get("Natural Gas Consumption - Energy Sector Own Use") * k})
            out.append(_bar(ac, "Gas use annual", use, "Taiwan natural gas use by sector, annual (Energy Administration)", "bcm per year", fmt="%Y"))
        lm, la = M("LNG"), A("LNG")
        for nm, d, units, fmt in (("LNG imports", lm, "Mt per month", "%Y-%m"), ("LNG imports annual", la, "Mt per year", "%Y")):
            if not d.empty:
                x = _top(d / 1000.0, 6)
                out.append(_bar(ac, nm, x, "Taiwan LNG imports by origin" + (", annual" if "annual" in nm else "") + " (Energy Administration)", units, fmt=fmt))
        ip, op = M("IMPPRICE"), M("OILPRICE")
        if not ip.empty:
            out.append(_bar(ac, "LNG price", ip[[c for c in ip.columns if c.startswith("LNG")]].rename(columns=lambda c: "LNG import price (CIF)"),
                            "Taiwan LNG import price (Energy Administration)", "US$ per tonne", "line"))
            out.append(_bar(ac, "Coal price", ip[[c for c in ip.columns if "coal" in c]].rename(columns=lambda c: c.split(" (")[0]), "Taiwan coal import prices (Energy Administration)", "US$ per tonne", "line"))
        if not op.empty:
            out.append(_bar(ac, "Oil price", op.rename(columns={"WTI": "WTI", "BRENT": "Brent", "DUBAI": "Dubai"}), "International crude oil prices (Energy Administration)", "US$ per barrel", "line"))
        cr = M("CRUDESRC")
        if not cr.empty:
            out.append(_bar(ac, "Crude imports", _top(cr / 1000.0, 6), "Taiwan crude oil imports by origin (Energy Administration)", "million barrels per month"))
        cv = M("CRUDE")
        if not cv.empty:
            out.append(_bar(ac, "Refinery intake", cv[["Refinery Intake"]].rename(columns={"Refinery Intake": "Refinery intake"}) / 1000.0,
                            "Taiwan refinery crude intake (Energy Administration)", "Mtoe per month", "line"))
        co = M("COALSRC")
        if not co.empty:
            out.append(_bar(ac, "Coal imports", _top(co / 1000.0, 6), "Taiwan coal imports by origin (Energy Administration)", "Mt per month"))
        sa = A("SUP")
        if not sa.empty:
            cols = {"Coal": "Coal and Coal Products", "Oil": "Crude Oil and Petroleum Products", "Gas": "Natural Gas", "Nuclear": "Nuclear",
                    "Renewables": "Renewable Energy - Subtotal"}
            out.append(_bar(ac, "Energy supply annual", _pick(sa, cols, 0.001), "Taiwan primary energy supply by fuel, annual (Energy Administration)",
                            "Mtoe per year", fmt="%Y"))
        return [s for s in out if s]
    return f


def taipower_specs(ac):
    def f(p):
        out = []
        k = _sheet(p, "Peak", True)
        if not k.empty:
            out.append(ac.spec("Peak load", k[["Peak_load_GW", "Net_peak_supply_capacity_GW"]].rename(columns={
                "Peak_load_GW": "Peak load", "Net_peak_supply_capacity_GW": "Net peak supply capacity"}),
                "Taiwan daily peak load and supply capacity (Taipower)", "GW", "line", "%Y-%m-%d"))
            out.append(ac.spec("Reserve margin", k[["Reserve_margin_pct"]].rename(columns={"Reserve_margin_pct": "Reserve margin"}),
                               "Taiwan daily peak reserve margin (Taipower)", "% of peak load", "line", "%Y-%m-%d"))
            u = k[["Industrial_use_GWh", "Residential_use_GWh"]].rename(columns={"Industrial_use_GWh": "Industrial", "Residential_use_GWh": "Residential"})
            out.append(ac.spec("Industrial and residential use", u / 24.0, "Taiwan daily electricity use, industrial and residential (Taipower)",
                               "GW (daily average)", "line", "%Y-%m-%d"))
        return out
    return f


def reservoir_specs(ac):
    def f(p):
        out = []
        names = {}
        try:
            names = dict(pd.read_excel(p, sheet_name="Reservoirs", index_col=0)["name"])
            vol = _sheet(p, "Daily storage", True)
        except Exception:  # noqa: BLE001
            return out
        if vol.empty:
            return out
        big = vol.max().sort_values(ascending=False).index[:12]
        top = vol[list(big[:7])] / 100.0
        top.columns = [str(names.get(int(c), c)) for c in top.columns]
        out.append(ac.spec("Storage daily", top.dropna(how="all"), "Taiwan largest reservoirs: daily storage (Water Resources Agency)",
                           "million m3", "line", "%Y-%m-%d"))
        for rid in big:
            nm = names.get(int(rid) if str(rid).isdigit() else rid, str(rid))
            s = vol[rid].dropna()
            if len(s) < 2:
                continue
            out.append({"name": str(nm), "water_year": s, "y_decimals": 0, "title": f"{nm} reservoir storage (Water Resources Agency)",
                        "units": "million m3", "sheet": f"Water year - {nm}"[:31]})
        return out
    return f


ROLL_FUELS = [("Hydro", "Hydro"), ("Gas", "Gas"), ("Wind", "Wind"), ("Solar", "Solar"), ("Coal", "Coal"), ("Nuclear", "Nuclear"),
              ("Oil", "Oil"), ("Cogeneration", "Cogeneration"), ("Pumped_storage", "Pumped storage"),
              ("Other_renewables", "Other renewables")]


def _roll_gw(d, flag, last_days=None, freq=None):
    """Average-GW columns of a roll-up sheet, charted periods only (flag column = 1), a gap row for every missing period."""
    d = d[d[flag] == 1]
    if d.empty:
        return pd.DataFrame()
    if last_days:
        d = d[d.index >= d.index.max() - pd.Timedelta(days=last_days)]
    out = pd.DataFrame({new: d[f"{old}_GW"] for old, new in ROLL_FUELS if f"{old}_GW" in d.columns}, index=d.index)
    out = out.loc[:, out.abs().sum() > 0]
    return out.reindex(pd.date_range(out.index.min(), out.index.max(), freq=freq)) if freq else out


def rollup_specs(ac):
    def f(p):
        out = []
        for tag, who, what in (("Zenodo", "Taipower units + IPPs", "2017-22"), ("EMS", "Taipower EMS units", "")):
            dd, mm = _sheet(p, f"Daily {tag}", True), _sheet(p, f"Monthly {tag}", True)
            if not mm.empty:
                m = _roll_gw(mm, "charted", None, "MS")
                if not m.empty:
                    out.append(ac.spec(f"Generation monthly {tag}", m, f"Taiwan generation by fuel, monthly, {who} (Taipower open data)",
                                       "GW (monthly average)", "stacked_bar", "%Y-%m"))
            if not dd.empty:
                d = _roll_gw(dd, "complete", 730, "D")
                if not d.empty:
                    out.append(ac.spec(f"Generation daily {tag}", d, f"Taiwan generation by fuel, daily, {who} (Taipower open data)",
                                       "GW (daily average)", "stacked_area", "%Y-%m-%d"))
        return out
    return f


def live_specs(ac):
    def f(p):
        d = _sheet(p, "Live daily", True)
        if d.empty:
            return []
        d = d[d["charted"] == 1]
        if d.empty:
            return []
        names = [("Hydro", "Hydro"), ("Gas", "Gas"), ("Wind", "Wind"), ("Solar", "Solar"), ("Coal", "Coal"), ("Oil", "Oil"),
                 ("Cogeneration", "Cogeneration"), ("Other_renewables", "Other renewables"),
                 ("Storage_discharge", "Pumped storage and batteries")]
        g = pd.DataFrame({new: d[f"{old}_GW"] for old, new in names}, index=d.index)
        g = g.loc[:, g.abs().sum() > 0].reindex(pd.date_range(d.index.min(), d.index.max(), freq="D"))
        return [ac.spec("Generation daily live", g, "Taiwan generation by fuel, daily mean of hourly snapshots (Taipower live data)",
                        "GW (daily average)", "stacked_area", "%Y-%m-%d")]
    return f


def register(registry, ac):
    registry[ROLLUP] = rollup_specs(ac)
    registry[LIVE] = live_specs(ac)
    registry[ESIST] = esist_specs(ac)
    registry[TAIPOWER] = taipower_specs(ac)
    registry[RESERVOIRS] = reservoir_specs(ac)
