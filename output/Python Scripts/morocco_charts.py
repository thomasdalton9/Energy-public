"""Chart specs for morocco_gas_daily.xlsx (africa/MOROCCO_GAS.py), registered into add_charts.REGISTRY.

  Tarifa daily            ENTSOG physical flow at Tarifa, both directions, Bcf/d (DERIVED from GWh/d at 11.7 GWh/MNm3)
  Gas to power YTD        Ministry 'Chiffres cles' year-to-date gas for power (bars) vs Tarifa year to date (line)
  Tarifa monthly          ENTSOG monthly net balance vs the Enagas bulletin Tarifa row (cross-check, from 2017)
  Plants generation       HCP annual generation per plant
  Nm3 per MWh             HCP gas burned / generation per plant (DERIVED)
  Estimate vs Ember       monthly gas-fired generation ESTIMATE vs Ember (FALLBACK)
  Comtrade price          UN Comtrade import value per MMBtu (DERIVED), partner total
"""
import pandas as pd

ESTIMATE = "Gas-fired generation ESTIMATE (Tarifa gas / official heat rate)"
EMBER = "Gas generation - Ember (FALLBACK)"


def _sheet(path, name, idx):
    d = pd.read_excel(path, sheet_name=name)
    d = d.drop(columns=[c for c in d.columns if str(c).startswith("Unnamed")])
    if idx not in d.columns or not len(d):
        return pd.DataFrame()
    d[idx] = pd.to_datetime(d[idx], errors="coerce")
    return d.dropna(subset=[idx]).set_index(idx).sort_index()


def specs(ac):
    def f(p):
        out = []
        d = _sheet(p, "Tarifa daily", "date")
        if len(d):
            g = d[["Exit_ES_to_MA_Bcf_d_DERIVED", "Entry_MA_to_ES_Bcf_d_DERIVED"]].apply(pd.to_numeric, errors="coerce")
            g = g.rename(columns={"Exit_ES_to_MA_Bcf_d_DERIVED": "Spain to Morocco (Enagas exit; gas to power from 2023)",
                                  "Entry_MA_to_ES_Bcf_d_DERIVED": "Morocco to Spain (Enagas entry; Algerian transit to Oct 2021)"})
            out.append(ac.spec("Tarifa daily", g.dropna(how="all"),
                               "Morocco gas: GME flow at Tarifa, daily (ENTSOG / Enagas)", "Bcf/d", "line", "%Y-%m-%d"))
        y = _sheet(p, "Ministry YTD", "date")
        m = _sheet(p, "Tarifa monthly", "month")
        if len(m):
            t = m[(m.index >= "2023-01-01")]
            t = t[t["Complete"].astype(str) == "yes"]
            mn = pd.to_numeric(t["Gas_to_power_MNm3_DERIVED"], errors="coerce")
            # running total within each calendar year; it stops at the year's first missing month (no filling)
            run, last, tot = {}, None, 0.0
            for d_, v in mn.dropna().items():
                if d_.month == 1:
                    last, tot = d_, 0.0
                elif last is None or (d_.year, d_.month - 1) != (last.year, last.month):
                    last = None
                    continue
                tot += v
                last = d_
                run[d_] = round(tot, 1)
            tab = pd.DataFrame({"Tarifa exit, year to date (DERIVED at 11.7 GWh per MNm3)": pd.Series(run, dtype=float)})
            if len(y):
                cur = y[y["Role"].astype(str) == "current year"]
                prev = y[y["Role"].astype(str) != "current year"]
                mi = pd.concat([cur["Gas_to_power_YTD_MNm3"], prev["Gas_to_power_YTD_MNm3"]])
                mi = mi[~mi.index.duplicated()].sort_index()
                tab = tab.join(mi.rename("Ministry gas for power, year to date (Chiffres cles, ONEE)"), how="outer")
            tab = tab[tab.index >= "2023-01-01"]
            cols = ["Ministry gas for power, year to date (Chiffres cles, ONEE)",
                    "Tarifa exit, year to date (DERIVED at 11.7 GWh per MNm3)"]
            tab = tab[[c for c in cols if c in tab.columns]]
            out.append(ac.spec("Gas to power YTD", tab,
                               "Morocco gas for power, year to date: Ministry (bars) vs Tarifa exit (line)",
                               "million Nm3 (year to date)", "stacked_bar",
                               line_cols=("Tarifa exit, year to date (DERIVED at 11.7 GWh per MNm3)",)))
            c = pd.DataFrame(index=m.index)
            full = m["Complete"].astype(str) == "yes"
            c["ENTSOG net (entry - exit), complete months"] = pd.to_numeric(m["Net_entry_MA_to_ES_GWh"], errors="coerce").where(full)
            if "Enagas_bulletin_net_GWh" in m:
                c["Enagas monthly bulletin, Tarifa net"] = pd.to_numeric(m["Enagas_bulletin_net_GWh"], errors="coerce")
            out.append(ac.spec("Tarifa monthly", c.dropna(how="all"),
                               "Tarifa net flow, monthly: ENTSOG vs Enagas bulletin (+ = to Spain, - = to Morocco)",
                               "GWh per month", "line"))
        a = _sheet(p, "Plants annual", "date")
        if len(a):
            g = a[["Tahaddart_GWh", "AinBeniMathar_GWh"]].apply(pd.to_numeric, errors="coerce").rename(
                columns={"Tahaddart_GWh": "Tahaddart (EET, combined cycle)", "AinBeniMathar_GWh": "Ain Beni Mathar (ONEE, ISCC)"})
            out.append(ac.spec("Plants generation", g.dropna(how="all"),
                               "Morocco gas-fired generation by plant, annual (HCP yearbook, ONEE)", "GWh per year",
                               "stacked_bar", "%Y"))
            h = a[["Tahaddart_Nm3_per_MWh_DERIVED", "AinBeniMathar_Nm3_per_MWh_DERIVED",
                   "Both_plants_Nm3_per_MWh_DERIVED"]].apply(pd.to_numeric, errors="coerce").rename(columns={
                       "Tahaddart_Nm3_per_MWh_DERIVED": "Tahaddart", "AinBeniMathar_Nm3_per_MWh_DERIVED": "Ain Beni Mathar",
                       "Both_plants_Nm3_per_MWh_DERIVED": "Both plants"})
            out.append(ac.spec("Nm3 per MWh", h.dropna(how="all"),
                               "Morocco gas plants: gas burned per MWh, DERIVED from HCP (2022 = restart year)",
                               "Nm3 per MWh", "line", "%Y"))
        e = _sheet(p, "Estimate", "month")
        if len(e):
            cols = {"Gas_fired_GWh_ESTIMATE": ESTIMATE, "Gas_GWh_Ember_FALLBACK": EMBER}
            g = e[[c for c in cols if c in e]].apply(pd.to_numeric, errors="coerce").rename(columns=cols)
            out.append(ac.spec("Estimate vs Ember", g.dropna(how="all"),
                               "Morocco gas-fired generation, monthly: ESTIMATE from Tarifa vs Ember (FALLBACK)",
                               "GWh per month", "line"))
        t = _sheet(p, "Comtrade monthly", "date")
        if len(t):
            w = t[pd.to_numeric(t["Partner_code"], errors="coerce") == 0]
            g = pd.to_numeric(w["USD_per_MMBtu_DERIVED"], errors="coerce").rename(
                "Import value per MMBtu, all origins (DERIVED)").to_frame()
            out.append(ac.spec("Comtrade price", g.dropna(how="all"),
                               "Morocco gas import price (UN Comtrade HS 271121, value / MMBtu)", "US$/MMBtu", "line"))
        return out
    return f
