"""Chart specs for the Ghana workbooks (registered into add_charts.REGISTRY):

  ghana_wem_weekly_generation_daily.xlsx   Energy Commission weekly WEM statistics (africa/GHANA_WEM_WEEKLY_POWER.py)
  ghana_energy_commission_power.xlsx       Energy Commission annual statistics (africa/GHANA_ENERGY_COMMISSION.py)
"""
import pandas as pd


def _sheet(path, name):
    d = pd.read_excel(path, sheet_name=name, index_col=0)
    d.index = pd.to_datetime(d.index, errors="coerce")
    return d[d.index.notna()].sort_index()


def wem_specs(ac):
    def f(p):
        d = _sheet(p, "Daily")
        out = []
        g = d[["Hydro_GWh", "Thermal_GWh", "Solar_GWh"]].rename(columns=lambda c: c.replace("_GWh", ""))
        m = ac.complete_months(g, g.resample("MS").sum(min_count=1))
        m = m.dropna(how="all").fillna(0)
        if not m.empty:
            out.append(ac.spec("Monthly", m.round(1), "Ghana power generation by type (Energy Commission, weekly WEM statistics)",
                               "GWh per month", "stacked_bar"))
        last = g.dropna(how="all").tail(180)
        if len(last) > 7:
            out.append(ac.spec("Daily", last.fillna(0).round(2), "Ghana daily generation by type, latest 180 days (Energy Commission)",
                               "GWh per day", "stacked_bar", "%Y-%m-%d"))
        return out
    return f


def annual_specs(ac):
    def f(p):
        out = []
        a = _sheet(p, "Annual")
        a = a[["Hydro_GWh", "Thermal_GWh", "Other_Renewables_GWh"]].rename(
            columns={"Hydro_GWh": "Hydro", "Thermal_GWh": "Thermal", "Other_Renewables_GWh": "Solar & other renewables"}).fillna(0)
        out.append(ac.spec("Annual", a, "Ghana electricity generation by type, annual (Energy Commission; GRIDCo, ECG)",
                           "GWh per year", "stacked_bar", "%Y"))
        for sheet, title in (("Akosombo", "Akosombo Dam (Lake Volta) month-end water level"), ("Bui", "Bui Dam month-end water level")):
            try:
                s = _sheet(p, sheet)["Level_ft"].dropna()
            except Exception:  # noqa: BLE001
                continue
            if len(s) < 14:
                continue
            s.index = s.index + pd.offsets.MonthEnd(0)   # a month-end reading is dated its last day
            out.append({"name": sheet, "water_year": s.resample("D").interpolate(limit=31, limit_area="inside"),
                        "title": f"{title} (Energy Commission)", "units": "feet above sea level", "y_decimals": 0,
                        "sheet": f"Water year - {sheet}"})
        return out
    return f
