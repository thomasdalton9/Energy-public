"""Demand-per-head forecast along the GDP-per-head curve, with a backtest. Shared by the forecast animations."""
import numpy as np
import pandas as pd
import fundamentals as F


def curve_fit(x, kwh, w):
    v = np.log(kwh); best = None
    for K in np.linspace(5000, 20000, 31):
        for g0 in np.linspace(5, 80, 38):
            for b in np.linspace(0.8, 3.0, 23):
                r = v - np.log(K / (1 + (g0 / x) ** b)); s = np.sum(w * r * r)
                if best is None or s < best[0]:
                    best = (s, K, g0, b)
    K, g0, b = best[1:]
    return lambda g: K / (1 + (g0 / g) ** b)


def load(M):
    d = F.load(M + "/output/Data and Chart Outputs")
    return d


def panel(d, years):
    dem, pop, ppp = d["Demand_TWh"], d["Population_m"], d["GDP_PPP_USD_bn"]
    ok = [c for c in dem.columns if c != "WLD" and c in pop and c in ppp and
          all(not pd.isna(dem[c].get(y)) and not pd.isna(pop[c].get(y)) and not pd.isna(ppp[c].get(y)) and
              dem[c].get(y) > 0 for y in years) and pop[c].get(2024, 0) >= 0.3]
    gdp = pd.DataFrame({c: [ppp[c][y] / pop[c][y] for y in years] for c in ok}, index=years)
    kwh = pd.DataFrame({c: [dem[c][y] * 1e9 / (pop[c][y] * 1e6) for y in years] for c in ok}, index=years)
    popf = pd.DataFrame({c: [pop[c][y] for y in years] for c in ok}, index=years)
    return ok, gdp, kwh, popf


def project(kwh0, g0, path_g, curve, lam):
    """kWh per head along a GDP-per-head path: follow the curve, with the country's gap to the curve (log ratio)
    closing by lam per year."""
    gap = np.log(kwh0 / curve(g0))
    out = []
    for t, g in enumerate(path_g, start=1):
        out.append(curve(g) * np.exp(gap * (1 - lam) ** t))
    return np.array(out)


def backtest(d, base=2015, end=2024):
    yrs = list(range(base, end + 1))
    ok, gdp, kwh, popf = panel(d, yrs)
    curve = curve_fit(gdp.loc[base].values, kwh.loc[base].values, np.sqrt(popf.loc[base].values))
    res = {}
    for lam in [0, 0.02, 0.04, 0.06, 0.08, 0.1, 0.15]:
        pred = {c: project(kwh[c][base], gdp[c][base], gdp[c].loc[base + 1:end].values, curve, lam)[-1] for c in ok}
        err = pd.Series({c: np.log(pred[c] / kwh[c][end]) for c in ok})
        w = popf.loc[end] * kwh.loc[end]   # demand-weighted
        tot_pred = sum(pred[c] * popf[c][end] for c in ok); tot_act = sum(kwh[c][end] * popf[c][end] for c in ok)
        naive = pd.Series({c: np.log(kwh[c][base] * gdp[c][end] / gdp[c][base] / kwh[c][end]) for c in ok})
        res[lam] = dict(median_abs_err=float(np.median(np.abs(np.expm1(err)))) * 100,
                        demand_weighted_abs_err=float(np.average(np.abs(np.expm1(err)), weights=w)) * 100,
                        world_total_err=(tot_pred / tot_act - 1) * 100,
                        naive_elasticity1_median_abs_err=float(np.median(np.abs(np.expm1(naive)))) * 100)
    return pd.DataFrame(res).T, ok, gdp, kwh, popf, curve


def forecast(d, base=2024, horizon=2035, lam=0.04):
    """GDP per head from IMF real GDP growth / IMF population to 2031, then the 2029-31 average growth held to the
    horizon; kWh per head along the 2024 world curve with gap closing lam per year; demand = kWh x population."""
    ok, gdp, kwh, popf = panel(d, [base])
    curve = curve_fit(gdp.loc[base].values, kwh.loc[base].values, np.sqrt(popf.loc[base].values))
    gr, lp = d["GDP_growth_IMF_pct"], d["Population_IMF_m"]
    years = list(range(base + 1, horizon + 1))
    G, P, Kw, flags = {}, {}, {}, {}
    for c in ok:
        if c not in gr or c not in lp:
            continue
        g_rate = gr[c].reindex(range(base + 1, 2032))
        p_lvl = lp[c].reindex(range(base, 2032))
        if g_rate.isna().any() or p_lvl.isna().any():
            continue
        p_rate = p_lvl.pct_change().dropna()
        lg, lpop = g_rate.loc[2029:2031].mean() / 100, p_rate.loc[2029:2031].mean()
        g_path, p_path = [], []
        g, p = gdp[c][base], popf[c][base]
        for y in years:
            gy = g_rate.get(y, lg * 100) / 100
            py = p_rate.get(y, lpop)
            p *= 1 + py
            g *= (1 + gy) / (1 + py)
            g_path.append(g); p_path.append(p)
        G[c], P[c] = g_path, p_path
        Kw[c] = project(kwh[c][base], gdp[c][base], np.array(g_path), curve, lam)
    G = pd.DataFrame(G, index=years); P = pd.DataFrame(P, index=years); Kw = pd.DataFrame(Kw, index=years)
    return G, P, Kw, curve


def slope_fn(curve):
    def s(g, h=1e-3):
        return (np.log(curve(g * (1 + h))) - np.log(curve(g))) / np.log(1 + h)
    return s


def forecast_v2(d, base=2024, horizon=2035, own_from=2014, weight_own=0.5):
    """Backtested method (2015 -> 2024: median country error 12.8%, demand-weighted 5.6%, world total +3.3%): each year, demand per head
    grows by elasticity x GDP-per-head growth, elasticity = 50% the country's own 2014-24 elasticity (clipped -0.5..2)
    + 50% the slope of the 2024 world curve at its current income (so it falls as the country gets richer).
    GDP per head: IMF WEO real GDP growth and population to 2031, then the 2029-31 averages held to the horizon."""
    yrs = list(range(own_from, base + 1))
    ok, gdp, kwh, popf = panel(d, yrs)
    curve = curve_fit(gdp.loc[base].values, kwh.loc[base].values, np.sqrt(popf.loc[base].values))
    slope = slope_fn(curve)
    gr, lp = d["GDP_growth_IMF_pct"], d["Population_IMF_m"]
    years = list(range(base + 1, horizon + 1))
    G, P, Kw, EL = {}, {}, {}, {}
    for c in ok:
        if c not in gr or c not in lp:
            continue
        g_rate = gr[c].reindex(range(base + 1, 2032))
        p_lvl = lp[c].reindex(range(base, 2032))
        if g_rate.isna().any() or p_lvl.isna().any():
            continue
        p_rate = p_lvl.pct_change().dropna()
        lg, lpop = g_rate.loc[2029:2031].mean() / 100, p_rate.loc[2029:2031].mean()
        dl = np.log(gdp[c][base] / gdp[c][own_from])
        own = np.clip(np.log(kwh[c][base] / kwh[c][own_from]) / dl, -0.5, 2) if abs(dl) > 0.05 else None
        if kwh[c][base] > 3 * curve(gdp[c][base]):
            own = None   # more than 3x the curve (e.g. Bhutan's crypto-mining load): its own past jump is not
            # extrapolated, the country moves with the curve's slope only
        g, p, k = gdp[c][base], popf[c][base], kwh[c][base]
        gp, pp, kp = [], [], []
        for y in years:
            gy = g_rate.get(y, lg * 100) / 100
            py = p_rate.get(y, lpop)
            g_new = g * (1 + gy) / (1 + py)
            el = slope(g) if own is None else weight_own * own + (1 - weight_own) * slope(g)
            k = k * (g_new / g) ** el
            g, p = g_new, p * (1 + py)
            gp.append(g); pp.append(p); kp.append(k)
        G[c], P[c], Kw[c] = gp, pp, kp
        EL[c] = own
    return (pd.DataFrame(G, index=years), pd.DataFrame(P, index=years), pd.DataFrame(Kw, index=years), curve,
            pd.Series(EL))
