"""
Generic gap-filling of installed-capacity series: from whatever mix of monthly, quarterly and annual sources a country
has, build ONE gap-free monthly series per fuel, with every cell flagged by how it was obtained. Config-driven and free of
any country knowledge: europe/capacity/EUROPE_CAPACITY_COMPLETE.py holds the Europe config and the workbook output; later
regions (Americas, Asia, Oceania, East Med, Africa) write their own config and call the same functions.

CONVENTION  A row dated the 1st of a month is the stock at the END of the previous month (the repo-wide capacity
convention; ENTSO-E 1-January rows are the end-of-year stock). Callers shift sources that date a row by the month it
closes (`shift_months=1`) BEFORE passing them in. Values are MW.

INTERFACE
  Series(pts, res, kind=None, source="", url="", note="")
      pts   pandas Series, index = month starts (convention above), values MW; NaN rows are ignored
      res   resolution of the source: "M" monthly, "Q" quarterly, "A" annual, "EMBER" Ember yearly fallback
      kind  for res "M": "RAW" (published monthly) or "REBUILT" (register rebuilt to month-end)
  FuelCfg(primary=[Series, ...], ref=None, scale_to=None, linear=None, included_in=None, note="", band=(0.7, 1.3),
          memo=None)
      primary     ordered candidates; the FIRST that holds any point is used for the whole span (one basis per fuel)
      ref         annual (or quarterly) reference series used to back-fill months before a late-starting primary (splice)
                  and to report the level ratio; None = no splice
      scale_to    official anchors (quarterly/annual) that a REBUILT primary is scaled to, linearly interpolated factor;
                  if any anchor ratio is outside `band`, the primary is dropped and the anchors are used instead (a Break)
      linear      True/False to override the fuel class (default: LINEAR_FUELS grow continuously, all others change in steps)
      included_in name of another fuel column that already contains this one (flag INCLUDED_IN_<FUEL>, value blank)
      memo        unflagged pass-through Series (e.g. Lignite, Wind_Onshore) kept as memo columns
  fill_country(name, fuels, start, end, fuel_cfg, linear_fuels=LINEAR_FUELS) -> Result
      fuel_cfg  {fuel: FuelCfg}; fuels absent from the dict, or with no usable source, are left blank and flagged NONE
      Result.values  DataFrame (index = month starts start..end, one column per fuel, MW)
      Result.flags   DataFrame of the same shape, flag strings (below)
      Result.basis   list of dicts (country, fuel, source, resolution, rule, ratio, caveat) for a 'Basis' sheet
      Result.breaks  list of dicts (country, fuel, month, kind, ratio, detail) for a 'Breaks' sheet
      Result.memo    DataFrame of memo columns
  coverage(flags) -> DataFrame, share of cells per flag, plus observed / filled totals
  total(values, exclude=("PumpedStorage", "Storage")) -> Series, sum of the fuel columns (blank only if all blank)

FILL RULES (in order of preference, per fuel)
  1. monthly source: used as is. RAW (published) or REBUILT (register rebuild). REBUILT_SCALED = a rebuild multiplied by the
     interpolated ratio official anchors / rebuild (monthly shape kept, level pinned to the anchors).
  2. quarterly official: linear between quarter-end stocks. ANCHOR_Q at the stock itself, INTERP_Q between.
  3. annual official: ANCHOR_A at the stock; wind, solar, storage and bioenergy (LINEAR_FUELS) are linear between
     consecutive stocks (INTERP_A); every other fuel is held at the last stock until the next one (HELD_A).
     Short holes inside a monthly source (<= 6 months) are filled the same way: INTERP_GAP / HELD_GAP.
  4. after the last point: held flat to `end` (HELD_LATEST), no trend.
  5. splice: months before a monthly primary that starts after `start` are back-filled from `ref` by rules 3, multiplied by the
     ratio primary/ref at the first overlapping 1 January (SPLICED). A ratio outside `band` is not applied (SPLICED_UNSCALED)
     and is listed in Breaks.
  6. EMBER: a series from Ember's yearly data (res "EMBER") is filled like an annual one but every cell is FALLBACK_EMBER.
  7. no source at all: NONE (blank).
FLAGS  RAW REBUILT REBUILT_SCALED ANCHOR_Q ANCHOR_A INTERP_Q INTERP_A HELD_A INTERP_GAP HELD_GAP HELD_LATEST SPLICED
       SPLICED_UNSCALED FALLBACK_EMBER NONE INCLUDED_IN_<FUEL>
OBSERVED = RAW, REBUILT, REBUILT_SCALED, ANCHOR_Q, ANCHOR_A (a published stock at its own date); everything else is filled.
"""
from dataclasses import dataclass, field
from typing import Optional

import numpy as np
import pandas as pd

LINEAR_FUELS = ("Wind", "Solar", "Storage", "Bioenergy")
FUEL_ORDER = ["Hydro", "PumpedStorage", "Gas", "Coal", "Nuclear", "Oil", "Wind", "Solar", "Bioenergy", "Storage", "Other"]
BAND = (0.7, 1.3)
MAX_GAP_MONTHS = 6          # hole inside a monthly source that counts as a gap rather than an annual step
MAX_Q_GAP_MONTHS = 4        # distance between points of a quarterly source
OBSERVED = ("RAW", "REBUILT", "REBUILT_SCALED", "ANCHOR_Q", "ANCHOR_A")
FLAGS = ["RAW", "REBUILT", "REBUILT_SCALED", "ANCHOR_Q", "ANCHOR_A", "INTERP_Q", "INTERP_A", "HELD_A", "INTERP_GAP", "HELD_GAP",
         "HELD_LATEST", "SPLICED", "SPLICED_UNSCALED", "FALLBACK_EMBER", "NONE"]


@dataclass
class Series:
    pts: pd.Series
    res: str                       # "M", "Q", "A", "EMBER"
    kind: Optional[str] = None     # "RAW" / "REBUILT" for res "M"
    source: str = ""
    url: str = ""
    note: str = ""

    def clean(self):
        p = pd.to_numeric(self.pts, errors="coerce").dropna().sort_index()
        p.index = pd.DatetimeIndex(p.index).to_period("M").to_timestamp()
        return p[~p.index.duplicated(keep="last")]


@dataclass
class FuelCfg:
    primary: list = field(default_factory=list)
    ref: Optional[Series] = None
    scale_to: Optional[Series] = None
    linear: Optional[bool] = None
    included_in: Optional[str] = None
    note: str = ""
    band: tuple = BAND
    memo: Optional[dict] = None


@dataclass
class Result:
    values: pd.DataFrame
    flags: pd.DataFrame
    basis: list
    breaks: list
    memo: pd.DataFrame


def _months(a, b):
    return (b.year - a.year) * 12 + b.month - a.month


def expand(s, fuel_linear, idx):
    """Series -> (values, flags) on idx by rules 1-4/6. Months before the first point stay NaN / ''."""
    pts = s.clean()
    out = pd.Series(np.nan, index=idx)
    fl = pd.Series("", index=idx, dtype=object)
    if pts.empty:
        return out, fl
    dates = pts.index
    native = {"M": s.kind or "RAW", "Q": "ANCHOR_Q", "A": "ANCHOR_A", "EMBER": "FALLBACK_EMBER"}[s.res]
    for m in idx:
        if m in pts.index:
            out[m], fl[m] = pts[m], native
        elif m < dates[0]:
            continue
        elif m > dates[-1]:
            out[m], fl[m] = pts.iloc[-1], "HELD_LATEST"
        else:
            i = dates.searchsorted(m)
            a, b = dates[i - 1], dates[i]
            gap = _months(a, b)
            if s.res == "M":
                short = gap <= MAX_GAP_MONTHS
                lin_flag, held_flag = ("INTERP_GAP", "HELD_GAP") if short else ("INTERP_A", "HELD_A")
                linear = fuel_linear
            elif s.res == "Q" and gap <= MAX_Q_GAP_MONTHS:
                lin_flag, held_flag, linear = "INTERP_Q", "INTERP_Q", True
            else:
                lin_flag, held_flag, linear = "INTERP_A", "HELD_A", fuel_linear
            if linear:
                out[m] = pts[a] + (pts[b] - pts[a]) * _months(a, m) / gap
                fl[m] = lin_flag
            else:
                out[m], fl[m] = pts[a], held_flag
            if s.res == "EMBER":
                fl[m] = "FALLBACK_EMBER"
    if s.res == "EMBER":
        fl[fl != ""] = "FALLBACK_EMBER"
    return out, fl


def first_overlap_ratio(prim, ref):
    """Ratio primary / reference at the first 1 January where both hold a native point (else at the primary's first month
    where the reference's expanded value exists). Returns (ratio, date) or (None, None)."""
    p, r = prim.clean(), ref.clean()
    common = [d for d in p.index.intersection(r.index) if d.month == 1 and r[d] not in (0, np.nan) and p[d] is not None]
    for d in common:
        if r[d] > 0:
            return p[d] / r[d], d
    both = p.index.intersection(r.index)
    for d in both:
        if r[d] > 0:
            return p[d] / r[d], d
    if len(p) and len(r) and r.index[0] <= p.index[0] <= r.index[-1]:
        idx = pd.date_range(r.index[0], r.index[-1], freq="MS")
        v, _ = expand(ref, True, idx)
        d = p.index[0]
        if d in v.index and v[d] > 0:
            return p[d] / v[d], d
    return None, None


def scale_series(prim, anchors, band):
    """Multiply a rebuilt series by the linearly interpolated ratio anchors / rebuild. Returns (scaled Series, ratios Series,
    ok). ok False = some anchor ratio outside band (caller drops the rebuild)."""
    p, a = prim.clean(), anchors.clean()
    d = a.index.intersection(p.index)
    d = [x for x in d if p[x] > 0 and a[x] > 0]
    if not d:
        return None, pd.Series(dtype=float), False
    ratios = pd.Series([a[x] / p[x] for x in d], index=pd.DatetimeIndex(d))
    ok = bool(((ratios >= band[0]) & (ratios <= band[1])).all())
    if not ok:
        return None, ratios, False
    xs = np.array([x.value for x in ratios.index], dtype=float)
    f = np.interp(np.array([x.value for x in p.index], dtype=float), xs, ratios.values)
    return pd.Series(p.values * f, index=p.index), ratios, True


def _fmt(x):
    return "" if x is None else f"{x:.3f}"


def fill_fuel(country, fuel, cfg, idx, linear_fuels):
    """One fuel -> (values, flags, basis dict, breaks list)."""
    linear = cfg.linear if cfg.linear is not None else fuel in linear_fuels
    basis = {"country": country, "fuel": fuel, "source": "", "resolution": "", "rule": "", "ratio": "", "caveat": cfg.note}
    breaks = []
    nan = pd.Series(np.nan, index=idx)
    if cfg.included_in:
        flag = f"INCLUDED_IN_{cfg.included_in.upper()}"
        basis.update(source=f"included in {cfg.included_in}", rule=flag)
        return nan, pd.Series(flag, index=idx, dtype=object), basis, breaks
    prim = next((s for s in cfg.primary if s is not None and not s.clean().empty), None)
    vals, flg = nan.copy(), pd.Series("", index=idx, dtype=object)
    if prim is not None:
        use = prim
        if cfg.scale_to is not None and prim.kind == "REBUILT":
            sc, ratios, ok = scale_series(prim, cfg.scale_to, cfg.band)
            if ok:
                use = Series(sc, "M", "REBUILT_SCALED", prim.source, prim.url, prim.note)
                basis["ratio"] = f"{ratios.min():.3f}-{ratios.max():.3f} (anchors/rebuild)"
            else:
                rr = f"{ratios.min():.2f}-{ratios.max():.2f}" if len(ratios) else "n/a"
                breaks.append({"country": country, "fuel": fuel, "month": "", "kind": "REBUILD OUTSIDE BAND",
                               "ratio": rr, "detail": f"{prim.source}: rebuild is not within {cfg.band[0]}-{cfg.band[1]} of the "
                                                      f"official anchors ({cfg.scale_to.source}); anchors used instead"})
                use = cfg.scale_to
                basis["ratio"] = rr
        vals, flg = expand(use, linear, idx)
        basis.update(source=use.source or prim.source, resolution=use.res,
                     rule=use.kind if use.res == "M" else {"Q": "QUARTERLY", "A": "ANNUAL", "EMBER": "FALLBACK_EMBER"}[use.res])
        pts = use.clean()
        if cfg.ref is not None and use.res in ("M", "Q") and cfg.scale_to is None:
            r, d = first_overlap_ratio(use, cfg.ref)
            if r is not None:
                basis["ratio"] = f"{r:.3f} at {d:%Y-%m-%d} (vs {cfg.ref.source})"
                if not cfg.band[0] <= r <= cfg.band[1]:
                    breaks.append({"country": country, "fuel": fuel, "month": f"{d:%Y-%m}", "kind": "LEVEL DIFFERS FROM ANNUAL",
                                   "ratio": round(r, 3), "detail": f"{use.source} is {r:.2f} x {cfg.ref.source} at the first "
                                                                    f"1 January both hold; different basis, kept as published"})
        # splice: months before the first native point
        missing = vals.isna()
        if missing.any() and cfg.ref is not None and not cfg.ref.clean().empty:
            rv, rf = expand(cfg.ref, linear, idx)
            r, d = first_overlap_ratio(use, cfg.ref)
            scaled = r is not None and cfg.band[0] <= r <= cfg.band[1]
            for m in idx[missing.values]:
                if np.isnan(rv[m]):
                    continue
                vals[m] = rv[m] * (r if scaled else 1.0)
                flg[m] = "SPLICED" if scaled else "SPLICED_UNSCALED"
            if (flg.isin(["SPLICED", "SPLICED_UNSCALED"])).any():
                first_native = pts.index[0]
                basis["rule"] += " + " + ("SPLICED" if scaled else "SPLICED_UNSCALED")
                basis["ratio"] = f"{r:.3f} at {d:%Y-%m-%d}" if r is not None else "n/a"
                basis["caveat"] = (basis["caveat"] + " " if basis["caveat"] else "") + \
                    f"months before {first_native:%Y-%m} back-filled from {cfg.ref.source}"
                if not scaled:
                    breaks.append({"country": country, "fuel": fuel, "month": f"{first_native:%Y-%m}", "kind": "SPLICE LEVEL BREAK",
                                   "ratio": "" if r is None else round(r, 3),
                                   "detail": f"ratio {_fmt(r)} outside {cfg.band}: annual source used unscaled before the monthly source starts"})
    elif cfg.ref is not None and not cfg.ref.clean().empty:
        vals, flg = expand(cfg.ref, linear, idx)
        basis.update(source=cfg.ref.source, resolution=cfg.ref.res,
                     rule={"A": "ANNUAL", "Q": "QUARTERLY", "M": cfg.ref.kind or "RAW", "EMBER": "FALLBACK_EMBER"}[cfg.ref.res])
    for m in idx[flg.eq("").values]:
        flg[m] = "NONE"
    if (flg == "NONE").all():
        basis.update(source="none", rule="NONE", caveat=(cfg.note + " " if cfg.note else "") + "no source holds this fuel")
    elif (flg == "NONE").any():
        basis["caveat"] = (basis["caveat"] + " " if basis["caveat"] else "") + \
            f"{int((flg == 'NONE').sum())} months before the first source point have no value (NONE)"
    return vals, flg, basis, breaks


def fill_country(name, fuels, start, end, fuel_cfg, linear_fuels=LINEAR_FUELS):
    idx = pd.date_range(start, end, freq="MS")
    vals, flags, basis, breaks, memo = {}, {}, [], [], {}
    for f in fuels:
        cfg = fuel_cfg.get(f) or FuelCfg()
        v, fl, b, br = fill_fuel(name, f, cfg, idx, linear_fuels)
        vals[f], flags[f] = v, fl
        basis.append(b)
        breaks.extend(br)
        for k, s in (cfg.memo or {}).items():
            p = s.clean()
            memo[k] = p.reindex(idx)
    return Result(pd.DataFrame(vals, index=idx), pd.DataFrame(flags, index=idx), basis, breaks, pd.DataFrame(memo, index=idx))


def coverage(flags):
    """Share of the country's fuel-months per flag, plus OBSERVED and FILLED (NONE excluded from FILLED, listed on its own)."""
    cells = flags.values.ravel()
    cells = cells[~pd.isna(cells)]
    n = len(cells)
    out = {f: float((cells == f).sum()) / n * 100 if n else 0.0 for f in FLAGS}
    out["INCLUDED"] = float(np.char.startswith(cells.astype(str), "INCLUDED_IN_").sum()) / n * 100 if n else 0.0
    out["OBSERVED"] = sum(out[f] for f in OBSERVED)
    out["FILLED"] = 100.0 - out["OBSERVED"] - out["NONE"] - out["INCLUDED"]
    return pd.Series(out)


def total(values, exclude=("PumpedStorage", "Storage")):
    cols = [c for c in values.columns if c not in exclude]
    return values[cols].sum(axis=1, min_count=1)
