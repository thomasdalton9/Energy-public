"""Chart specs for the Taiwan workbooks (registered into add_charts.REGISTRY by register()):

  taiwan_esist_monthly.xlsx    Energy Administration monthly statistics (taiwan/TAIWAN_ESIST_MONTHLY.py)
  taiwan_taipower.xlsx         Taipower daily peak load / reserve margin and EMS-unit generation (taiwan/TAIWAN_TAIPOWER.py)
  taiwan_reservoirs_daily.xlsx WRA reservoir storage (taiwan/TAIWAN_WRA_RESERVOIRS.py)

Power types use the fixed names of xlsx_charts.FUEL_COLOURS (Hydro, Gas, Wind, Solar, Coal, Nuclear, Oil, Bioenergy,
Geothermal, Pumped storage).
"""
import os

import pandas as pd

ESIST = "taiwan_esist_monthly.xlsx"
TAIPOWER = "taiwan_taipower.xlsx"
RESERVOIRS = "taiwan_reservoirs_daily.xlsx"
GEN_COLS = {"Hydro": "Renewable Energy - Hydro", "Gas": "Thermal - LNG-Fired", "Wind": "Renewable Energy - Wind",
            "Solar": "Renewable Energy - Solar PV", "Coal": "Thermal - Coal-Fired", "Nuclear": "Nuclear",
            "Oil": "Thermal - Oil-Fired", "Bioenergy": ("Renewable Energy - Biomass", "Renewable Energy - Waste"),
            "Geothermal": "Renewable Energy - Geothermal", "Pumped storage": "Pumped Storage"}
CAP_ORDER = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Geothermal", "Pumped storage"]


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
        cap = {k: v for k, v in GEN_COLS.items()}
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
        origins = ["Qatar", "Australia", "United States", "Papua New Guinea", "Malaysia", "Brunei Darussalam", "Indonesia", "Russia", "Nigeria", "Others"]
        for nm, d, units, fmt in (("LNG imports", lm, "Mt per month", "%Y-%m"), ("LNG imports annual", la, "Mt per year", "%Y")):
            if not d.empty:
                x = d[[c for c in origins if c in d.columns]] / 1000.0
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
            top = [c for c in cr.columns if c != "Total"]
            out.append(_bar(ac, "Crude imports", cr[top] / 1000.0, "Taiwan crude oil imports by origin (Energy Administration)", "million barrels per month"))
        cv = M("CRUDE")
        if not cv.empty:
            out.append(_bar(ac, "Refinery intake", cv[["Refinery Intake"]].rename(columns={"Refinery Intake": "Refinery intake"}) / 1000.0,
                            "Taiwan refinery crude intake (Energy Administration)", "Mtoe per month", "line"))
        co = M("COALSRC")
        if not co.empty:
            top = [c for c in co.columns if c != "Grand Total"]
            out.append(_bar(ac, "Coal imports", co[top] / 1000.0, "Taiwan coal imports by origin (Energy Administration)", "Mt per month"))
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
        for rid in big:
            nm = names.get(int(rid) if str(rid).isdigit() else rid, str(rid))
            s = vol[rid].dropna()
            if len(s) < 2:
                continue
            out.append({"name": str(nm), "water_year": s, "y_decimals": 0, "title": f"{nm} reservoir storage (Water Resources Agency)",
                        "units": "million m3", "sheet": f"Water year - {nm}"[:31]})
        return out
    return f


def register(registry, ac):
    registry[ESIST] = esist_specs(ac)
    registry[TAIPOWER] = taipower_specs(ac)
    registry[RESERVOIRS] = reservoir_specs(ac)
