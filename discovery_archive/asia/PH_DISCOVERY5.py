"""
Philippines discovery, round 5: test the generation-by-fuel extension of asia/PHILIPPINES_IEMOP.py.
  1. one day from both DIPCEF (final) and DIPCER (raw): columns, interval range / count, MWh by fuel side by side
  2. the full pull into temporary workbooks (every listed day), then daily GWh by fuel for a few days, monthly
     totals and shares, the unmapped share, and add_charts.py on the result
"""
import os
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "asia"))
import PHILIPPINES_IEMOP as P  # noqa: E402


def out(*a):
    print(*a, flush=True)


def compare():
    raw, fin = P.listing("DIPCER"), P.listing("DIPCEF")
    both = sorted(d for d in fin if len(fin[d]) >= 24 and len(raw.get(d, [])) >= 24)
    out(f"raw days {min(raw)}..{max(raw)} ({len(raw)}); final days {min(fin)}..{max(fin)} ({len(fin)}); both {len(both)}")
    fuel_of = P.FuelMap()
    d = both[-1]
    rows = {}
    for src, lst in (("DIPCEF", fin), ("DIPCER", raw)):
        with ThreadPoolExecutor(8) as ex:
            res = list(ex.map(P.get, lst[d]))
        mixes = []
        for c in res:
            _, m, _ = P.dipc_hour(c, fuel_of)
            mixes.append(m)
        allm = pd.concat(mixes)
        out(f"  {src} {d}: intervals {len(allm)} unique {allm.index.nunique()} from {allm.index.min()} to "
            f"{allm.index.max()}; files {sorted(u.rsplit('/', 1)[-1] for u in lst[d])[:2]}..")
        rows[src] = P.mix_row(mixes)
    import io
    import zipfile
    z = zipfile.ZipFile(io.BytesIO(P.get(fin[d][0])))
    x = pd.read_csv(z.open(z.namelist()[0]))
    out(f"  DIPCEF columns {list(x.columns)} rows {len(x)}")
    out((pd.DataFrame(rows) / 1000).round(2).to_string())


def full():
    tmp = os.path.join(ROOT, "tmp_ph")
    os.makedirs(tmp, exist_ok=True)
    mk, gen = os.path.join(tmp, "philippines_power_market.xlsx"), os.path.join(tmp, "philippines_power_generation_daily.xlsx")
    t0 = time.time()
    r = subprocess.run([sys.executable, os.path.join(ROOT, "asia", "PHILIPPINES_IEMOP.py"), "--out", mk, "--mix-out", gen],
                       capture_output=True, text=True, timeout=1150)
    out(f"pull took {time.time() - t0:.0f}s, exit {r.returncode}")
    out("\n".join(r.stdout.splitlines()[-60:]))
    out(r.stderr[-3000:])
    d = pd.read_excel(gen, sheet_name="Daily", index_col=0)
    d.index = pd.to_datetime(d.index)
    fuel = [c for c in d.columns if c.endswith("_MWh")]
    out("\nDaily GWh (last 6 days and 3 early days):")
    out((pd.concat([d.head(3), d.tail(6)])[fuel] / 1000).round(1).to_string())
    out(d[["Intervals", "Source"]].value_counts().to_string()[:1500])
    m = d[fuel].resample("MS").sum() / 1000
    m["days"] = d["Total_MWh"].resample("MS").count()
    out("\nMonthly GWh:")
    out(m.round(0).to_string())
    sh = d[fuel].sum() / d["Total_MWh"].sum() * 100
    out("\nShares over all days (%):")
    out(sh.round(2).to_string())
    dem = pd.read_excel(gen, sheet_name="Demand", index_col=0)
    out(dem.tail(3).to_string())
    r = subprocess.run([sys.executable, os.path.join(ROOT, "add_charts.py"), gen], capture_output=True, text=True)
    out("add_charts:", r.stdout[-1500:], r.stderr[-1500:])


if __name__ == "__main__":
    for f in sys.argv[1:] or ["compare", "full"]:
        out(f"\n==================== {f}")
        try:
            globals()[f]()
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            out(f"!! {f}: {e!r}")
