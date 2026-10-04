"""Cambodia probe 9: re-run (listing fallback, HEAD-or-GET check) (text fallback for the rotated 2019/2020 annexes, pypdfium2 page search) asia/CAMBODIA_EAC.py to a temp workbook (twice: the second run must download nothing) and
print a summary."""
import os
import subprocess
import sys
import tempfile

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
tmp = os.path.join(tempfile.mkdtemp(), "cambodia_power_generation.xlsx")
for run in (1, 2):
    print(f"######## run {run}", flush=True)
    subprocess.run([sys.executable, os.path.join(ROOT, "asia", "CAMBODIA_EAC.py"), "--out", tmp], check=False)
pd.set_option("display.width", 250)
for sh in ("Units", "Daily", "Monthly", "Basis"):
    d = pd.read_excel(tmp, sheet_name=sh)
    print(f"######## {sh} {d.shape}")
    print(d.to_string(max_colwidth=150))
raw = pd.read_excel(tmp, sheet_name="Raw")
print("######## Raw", raw.shape)
print(raw.groupby("Publication").size().to_string())
