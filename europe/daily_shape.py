"""
Give monthly official statistics a real daily shape.

Several power/gas workbooks take a MONTHLY official total (Eurostat, CBS, BFE) that used to be spread evenly over the days of the month, so
the daily series was flat inside each month and stepped at the cut-over to the ENTSO-E/ENTSOG daily values. `reshape()` keeps the official
monthly total exactly and takes the day-to-day shape from a daily series of the same quantity (ENTSO-E, ENTSOG):

    daily_i = shape_i x (official monthly total / shape monthly total)

Rules: a shape with missing days is filled with its own month mean for those days (needs >= MIN_COVER of the days); if the shape is zero or
missing for the month, the `fallback` shape (e.g. ENTSO-E total generation or load) gives the weights; if neither exists the month is spread
evenly. The scale factor must lie within [lo, hi]; outside it the month is spread evenly instead and logged (a shape that far from the official
figure is not a credible profile). Signed series (net imports) use an additive shift (daily_i = shape_i + (official - shape total) / days)
because a ratio is meaningless when the total is near zero or changes sign.
"""
import numpy as np
import pandas as pd

MIN_COVER = 0.8


def reshape(official_month, shape, fallback=None, lo=0.2, hi=5.0, additive=False, label="", log=None, max_shift_ratio=None):
    """official_month: Series indexed by month start -> monthly total (same unit as the output days). shape / fallback: daily Series.
    Returns the daily Series for every day of every official month; log (a list) collects the months that did not use `shape`."""
    log = log if log is not None else []
    parts = []
    for m, tot in official_month.dropna().items():
        days = pd.date_range(m, m + pd.offsets.MonthEnd(0))
        even = pd.Series(tot / len(days), index=days)
        res, how = even, "even (no shape)"
        if tot == 0:
            parts.append(even)
            continue
        for name, src in (("shape", shape), ("fallback", fallback)):
            if src is None:
                continue
            s = pd.to_numeric(src.reindex(days), errors="coerce")
            if s.notna().mean() < MIN_COVER:
                continue
            s = s.fillna(s.mean())
            if additive:
                if name == "fallback":
                    continue
                shift = (tot - s.sum()) / len(days)
                if max_shift_ratio is not None and abs(shift) > max_shift_ratio * max(abs(s).mean(), 1.0):
                    how = f"even (additive shift {shift:.0f} too large vs mean {abs(s).mean():.0f})"
                    break
                res, how = s + shift, "shape"
                break
            s = s.clip(lower=0)
            if s.sum() <= 0:
                continue
            scale = tot / s.sum()
            if name == "shape" and not (lo <= scale <= hi):
                how = f"even (scale {scale:.2f} outside {lo}-{hi})"
                continue
            res, how = s * scale, name
            break
        if how != "shape":
            log.append(f"{label} {m:%Y-%m}: {how}")
        parts.append(res)
    return pd.concat(parts) if parts else pd.Series(dtype=float)


def check_monthly(daily, official_month, label):
    """Assert daily sums equal the official monthly totals; return the largest absolute deviation."""
    got = daily.groupby(daily.index.to_period("M").to_timestamp()).sum()
    dev = (got.reindex(official_month.index) - official_month).abs().max()
    return 0.0 if pd.isna(dev) else float(dev)


def print_log(log, prefix="  shape fallback: "):
    """Print the fallback log compactly: one line per series with the count and first/last months."""
    by = {}
    for ln in log:
        lab, rest = ln.rsplit(" ", 2)[0], ln.split(" ", )[-1]
        key = (ln.split(": ", 1)[0].rsplit(" ", 1)[0], ln.split(": ", 1)[1])
        by.setdefault(key, []).append(ln.split(": ", 1)[0].rsplit(" ", 1)[1])
    for (lab, how), ms in by.items():
        print(f"{prefix}{lab}: {how} in {len(ms)} month(s): {', '.join(ms[:4])}{' ...' if len(ms) > 4 else ''}{' ' + ms[-1] if len(ms) > 4 else ''}", flush=True)
