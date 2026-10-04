"""
Capacity round 6 (rerun: pick the per-grid PDF): run asia/PHILIPPINES_DOE_CAPACITY.py twice to a temp file (the second run should find no new
release) and print the sheets.
"""
import os
import subprocess
import sys
import tempfile

import pandas as pd

pd.set_option("display.width", 250)
p = os.path.join(tempfile.mkdtemp(), "philippines_power_capacity.xlsx")
for _ in range(2):
    subprocess.run([sys.executable, "asia/PHILIPPINES_DOE_CAPACITY.py", "--out", p], timeout=900)
if os.path.exists(p):
    for s in ("Monthly", "Dependable", "Release"):
        print(f"--- {s}", flush=True)
        print(pd.read_excel(p, sheet_name=s).to_string(), flush=True)
    g = pd.read_excel(p, sheet_name="By grid")
    print(g.groupby(["measure", "scope", "grid"]).size().to_string(), flush=True)
