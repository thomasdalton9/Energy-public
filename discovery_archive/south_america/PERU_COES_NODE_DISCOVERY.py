"""
Why does COES's marginal-cost export miss half-hours at SANTA ROSA 220 kV (node STAROSA220), and which other
node would give a complete Lima price? One-off probe for SA_POWER_PRICES_DAILY.py (231 Santa Rosa days since
2021 have fewer than 40 of 48 half-hours).

For a few months (gap-heavy ones plus a clean control month) it downloads COES's ExportarMasivo file and reports:
  1. Santa Rosa: rows per day, rows with a blank TOTAL, and every bar/node whose name looks like Santa Rosa
     (a renamed node would show as a second name taking over on the missing days).
  2. Every node: share of days with all 48 half-hours; for nodes complete on Santa Rosa's missing days, how
     close they track Santa Rosa on days both are complete (mean absolute difference, correlation of daily means).
Writes peru_node_probe.csv (per node and month) for the workflow artifact.

Usage: python3 PERU_COES_NODE_DISCOVERY.py [YYYY-MM ...]
"""
import io
import sys
import time

import pandas as pd
import requests

COES = "https://www.coes.org.pe/Portal/mercadomayorista/costosmarginales/ExportarMasivo"
NODE = "STAROSA220"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0 Safari/537.36"}
MONTHS = sys.argv[1:] or ["2026-03", "2023-09", "2026-07", "2024-04", "2025-07"]   # last = no Santa Rosa gaps


def month_file(m):
    a = pd.Timestamp(m + "-01")
    b = a + pd.offsets.MonthEnd(0) + pd.Timedelta(days=1)
    url = f"{COES}?fechaInicio={a:%d/%m/%Y}&fechaFin={b:%d/%m/%Y}"
    for i in range(4):
        try:
            r = requests.get(url, headers=UA, timeout=600)
            r.raise_for_status()
            break
        except requests.RequestException as e:
            print(f"  retry {i + 1}: {type(e).__name__}", flush=True)
            time.sleep(10 * (i + 1))
    if r.content[:2] != b"PK":
        print(f"{m}: not an xlsx ({r.headers.get('content-type')})", flush=True)
        return None
    raw = pd.read_excel(io.BytesIO(r.content), header=None)
    hdr = next(i for i in range(15) if "NOMBRE BARRA" in [str(v).strip() for v in raw.iloc[i].values])
    raw.columns = [str(c).strip() for c in raw.iloc[hdr].values]
    raw = raw.iloc[hdr + 1:].copy()
    print(f"{m}: {len(r.content) / 1e6:.1f} MB, columns {list(raw.columns)}", flush=True)
    raw["NODO EMD"] = raw["NODO EMD"].astype(str).str.strip()
    raw["NOMBRE BARRA"] = raw["NOMBRE BARRA"].astype(str).str.strip()
    raw["t"] = pd.to_datetime(raw["FECHA HORA"].astype(str), format="%d/%m/%Y %H:%M", errors="coerce")
    raw["TOTAL"] = pd.to_numeric(raw["TOTAL"], errors="coerce")
    raw["date"] = (raw["t"] - pd.Timedelta(minutes=1)).dt.normalize()
    lo, hi = pd.Timestamp(m + "-01"), pd.Timestamp(m + "-01") + pd.offsets.MonthEnd(0)
    return raw[(raw["date"] >= lo) & (raw["date"] <= hi)]


rows = []
for m in MONTHS:
    print(f"\n===== {m}", flush=True)
    df = month_file(m)
    if df is None:
        continue
    print(f"  rows {len(df):,}; nodes {df['NODO EMD'].nunique()}; bars {df['NOMBRE BARRA'].nunique()}; "
          f"intervals {df['t'].nunique()} (a full month = {df['date'].nunique() * 48})", flush=True)
    # half-hours present per node per day: all rows, and rows with a value
    allrows = df.groupby(["NODO EMD", "date"])["t"].nunique().unstack(fill_value=0)
    valued = df.dropna(subset=["TOTAL"]).groupby(["NODO EMD", "date"])["t"].nunique().unstack(fill_value=0)
    days = allrows.columns
    # 1. Santa Rosa
    sr_like = df[df["NOMBRE BARRA"].str.contains("ROSA", case=False) | df["NODO EMD"].str.contains("ROSA", case=False)]
    print("  Santa-Rosa-like names:", sr_like.groupby(["NODO EMD", "NOMBRE BARRA"]).size().to_dict(), flush=True)
    if NODE in allrows.index:
        a, v = allrows.loc[NODE], valued.reindex(index=[NODE]).fillna(0).iloc[0]
        bad = [d for d in days if v.get(d, 0) < 40]
        print(f"  {NODE}: days with <40 valued half-hours: {len(bad)} -> "
              f"{', '.join(f'{d:%d}({int(a[d])} rows/{int(v.get(d, 0))} valued)' for d in bad)}", flush=True)
        # which half-hours go missing on bad days (time of day)
        if bad:
            sr = df[(df["NODO EMD"] == NODE) & df["date"].isin(bad)].dropna(subset=["TOTAL"])
            have = sr.groupby("date")["t"].apply(lambda s: set(s.dt.strftime("%H:%M")))
            alltimes = {f"{h:02d}:{mm:02d}" for h in range(24) for mm in (0, 30)}
            missing = pd.Series([t for d in have.index for t in sorted(alltimes - have[d])]).value_counts()
            print(f"  missing half-hours by time of day (top 12): {missing.head(12).to_dict()}", flush=True)
    else:
        bad = list(days)
        print(f"  {NODE} not present at all this month", flush=True)
    # 2. every node: completeness, and fit to Santa Rosa
    full = (valued.reindex(index=allrows.index).fillna(0) >= 48)
    sr_mean = df[df["NODO EMD"] == NODE].groupby("date")["TOTAL"].mean()
    sr_ok = valued.reindex(index=[NODE]).fillna(0).iloc[0] >= 40 if NODE in valued.index else pd.Series(False, index=days)
    for node in allrows.index:
        nm = df.loc[df["NODO EMD"] == node, "NOMBRE BARRA"].iloc[0]
        node_mean = df[df["NODO EMD"] == node].groupby("date")["TOTAL"].mean()
        both = [d for d in days if sr_ok.get(d, False) and full.loc[node].get(d, False)]
        diff = (node_mean.reindex(both) - sr_mean.reindex(both)).abs().mean() if both else None
        corr = node_mean.reindex(both).corr(sr_mean.reindex(both)) if len(both) > 3 else None
        rows.append({"month": m, "node": node, "bar": nm, "days_full_48": int(full.loc[node].sum()),
                     "days": len(days), "full_on_SR_bad_days": int(sum(full.loc[node].get(d, False) for d in bad)),
                     "SR_bad_days": len(bad), "mean_abs_diff_vs_SR": None if diff is None else round(diff, 2),
                     "corr_vs_SR": None if corr is None else round(corr, 4)})

res = pd.DataFrame(rows)
res.to_csv("peru_node_probe.csv", index=False)
if not res.empty:
    agg = res.groupby(["node", "bar"]).agg(full_days=("days_full_48", "sum"), days=("days", "sum"),
                                         full_on_SR_bad=("full_on_SR_bad_days", "sum"), SR_bad=("SR_bad_days", "sum"),
                                         mad=("mean_abs_diff_vs_SR", "mean"), corr=("corr_vs_SR", "mean"))
    agg["complete_share"] = (agg["full_days"] / agg["days"]).round(3)
    lima = agg[agg.index.get_level_values("bar").str.contains(
        "ROSA|CHAVARR|SAN JUAN|BALNEARIO|VENTANILLA|ZAPALLAL|INDUSTRIAL|CARABAYLLO|PLANICIE|CHILCA|SANTA|HUACHO|LIMA",
        case=False, regex=True)]
    pd.set_option("display.width", 250)
    print("\n===== LIMA-AREA NODES (all probed months)", flush=True)
    print(lima.sort_values(["full_on_SR_bad", "mad"], ascending=[False, True]).head(40).to_string(), flush=True)
    print("\n===== MOST COMPLETE NODES OVERALL, closest to Santa Rosa first", flush=True)
    print(agg[agg["complete_share"] >= 0.95].sort_values("mad").head(30).to_string(), flush=True)
    print(f"\nnodes 100% complete in every probed month: {int((agg['complete_share'] == 1).sum())} of {len(agg)}",
          flush=True)
