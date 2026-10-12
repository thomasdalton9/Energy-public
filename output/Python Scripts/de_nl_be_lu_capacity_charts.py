"""Chart specs for germany/belgium/netherlands/luxembourg_power_capacity_{monthly,annual}.xlsx, registered into
add_charts.REGISTRY (small hook at the end of add_charts.py).

  Capacity           month-end (or year-end) installed capacity by fuel, stacked, GW; battery power and pumped storage
                     are added as their own stacks (they are outside Total_MW)
  Net additions      change in solar and wind capacity per month (per year for annual workbooks), GW
"""
import pandas as pd

STACK = [("Hydro_MW", "Hydro"), ("Gas_MW", "Gas"), ("Wind_MW", "Wind"), ("Solar_MW", "Solar"), ("Coal_MW", "Coal"),
         ("Nuclear_MW", "Nuclear"), ("Oil_MW", "Oil"), ("Bioenergy_MW", "Bioenergy"), ("Other_MW", "Other"),
         ("Storage_MW", "Battery storage"), ("PumpedStorage_MW", "Pumped storage")]


def specs(ac, title, annual=False, start="2021-01-01"):
    def f(p):
        d = ac.by_date(ac.read(p, "Monthly"), "date")
        d = d[d.index >= start]
        g = pd.DataFrame({lab: pd.to_numeric(d[c], errors="coerce") / 1000.0 for c, lab in STACK if c in d})
        g = g.loc[:, g.fillna(0).abs().sum() > 0]
        fmt = "%Y" if annual else "%b/%y"
        out = [ac.spec("Capacity", g, title, "GW installed (year end)" if annual else "GW installed (month end)", "stacked_bar", fmt)]
        full = ac.by_date(ac.read(p, "Monthly"), "date")
        add = pd.DataFrame({lab: pd.to_numeric(full[c], errors="coerce").diff() / 1000.0
                            for c, lab in (("Solar_MW", "Solar"), ("Wind_MW", "Wind")) if c in full})
        add = add[add.index >= start].dropna(how="all")
        if len(add) > 1 and add.abs().sum().sum() > 0:
            out.append(ac.spec("Net additions", add, title.split(" installed")[0] + (" net solar and wind capacity additions per year" if annual else " net solar and wind capacity additions per month"),
                               "GW added" + (" per year" if annual else " per month"), "stacked_bar", fmt))
        return out
    return f
