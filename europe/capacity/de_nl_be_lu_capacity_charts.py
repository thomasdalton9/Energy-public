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
        # the chart palette holds 8 series: fold pumped storage into hydro, then the smallest groups into one grey 'Other' stack
        if len(g.columns) > 8 and "Pumped storage" in g:
            hyd = g["Pumped storage"].fillna(0) + (g["Hydro"].fillna(0) if "Hydro" in g else 0)
            g = g.drop(columns=[c for c in ("Hydro", "Pumped storage") if c in g])
            g.insert(0, "Hydro (incl. pumped storage)", hyd)
        merged = []
        for c in ("Oil", "Other", "Nuclear"):
            if len(g.columns) - max(len(merged) - 1, 0) <= 8:
                break
            if c in g:
                merged.append(c)
        if len(merged) > 1:
            tot = g[merged].fillna(0).sum(axis=1)
            g = g.drop(columns=merged)
            g["Other (" + ", ".join(m.lower() for m in merged) + ")"] = tot
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
