"""
South & Southeast Asia gas, round 8: re-test the pulls after the parser fixes (Migas table layouts, NSO old .xls, PBS totals) (GitHub Actions). Runs
asia/VIETNAM_NSO_GAS.py, asia/INDONESIA_MIGAS_GAS.py and asia/PAKISTAN_PBS_GAS.py with --out to a temp dir (twice,
to check the incremental second run), then prints each workbook's sheets and a sanity check against known
magnitudes (Vietnam ~0.6-0.9 bcf/d, Indonesia ~5.5-7 bcf/d, Pakistan ~2.9-3.5 bcf/d domestic).
"""
import os
import subprocess
import sys
import tempfile
import time

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TMP = tempfile.mkdtemp()
RUNS = [("VIETNAM_NSO_GAS", "vietnam_gas.xlsx"), ("INDONESIA_MIGAS_GAS", "indonesia_gas.xlsx"),
        ("PAKISTAN_PBS_GAS", "pakistan_gas.xlsx")]


def out(*a):
    print(*a, flush=True)


for script, fn in RUNS:
    path = os.path.join(TMP, fn)
    for k in (1, 2):
        t = time.time()
        out(f"\n{'#' * 20} {script} run {k} {'#' * 20}")
        p = subprocess.run([sys.executable, os.path.join(ROOT, "asia", f"{script}.py"), "--out", path],
                           capture_output=True, text=True, timeout=2400)
        lines = (p.stdout + p.stderr).splitlines()
        out("\n".join(lines if len(lines) < 400 else lines[:150] + ["..."] + lines[-200:]))
        out(f"-- exit {p.returncode} in {time.time() - t:.0f}s")
    if os.path.exists(path):
        x = pd.read_excel(path, sheet_name=None, index_col=0)
        for name, df in x.items():
            out(f"\n== {fn} / {name}: {df.shape}")
            out(df.head(3).to_string()[:1500] if name in ("Units", "Files", "Books") else df.tail(5).iloc[:, :10].to_string()[:2500])
        try:
            if fn == "vietnam_gas.xlsx":
                s = x["Production"]["Natural_gas_mcm_per_day"].dropna()
                out(f"SANITY Vietnam: last 12m {s.tail(12).mean() * 35.315 / 1000:.2f} bcf/d; months {len(s)} "
                    f"{s.index.min()}..{s.index.max()}")
                yr = x["Production"]["Natural_gas_mcm"].groupby(pd.to_datetime(x["Production"].index).year).agg(["sum", "count"])
                out(yr.to_string())
            if fn == "indonesia_gas.xlsx":
                s = x["Production"]["Total"].dropna()
                out(f"SANITY Indonesia: last 12m {s.tail(12).mean() / 1000:.2f} bcf/d; months {len(s)}")
                out(s.to_string())
            if fn == "pakistan_gas.xlsx":
                if "Production" in x:
                    s = x["Production"]["Pakistan_total"].dropna()
                    out(f"SANITY Pakistan: last 12m {s.tail(12).mean() / 1000:.2f} bcf/d; months {len(s)}")
                    out(x["Production"].to_string()[:4000])
                if "LNG imports" in x:
                    out(x["LNG imports"].to_string())
        except Exception as e:  # noqa: BLE001
            out(f"sanity failed: {e}")
