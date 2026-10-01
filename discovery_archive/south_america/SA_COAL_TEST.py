"""
Test run of south_america/SOUTH_AMERICA_COAL.py in GitHub Actions without committing anything: writes the
workbook to a temp path, adds the charts, prints every sheet (head / tail), then runs it a second time against
that workbook to check the incremental path. Also prints the Peru Anuario annex 'Produccion' rows around coal.
"""
import io
import os
import re
import subprocess
import sys
import tempfile

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "south_america"))
pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 30)

out = os.path.join(tempfile.mkdtemp(), "south_america_coal_production.xlsx")
for run in (1, 2):
    print(f"\n######## run {run}", flush=True)
    subprocess.run([sys.executable, os.path.join(ROOT, "south_america", "SOUTH_AMERICA_COAL.py"), "--out", out], check=True)
    subprocess.run([sys.executable, os.path.join(ROOT, "add_charts.py"), out], check=True)
xl = pd.ExcelFile(out)
print(f"\nsheets: {xl.sheet_names}")
for s in xl.sheet_names:
    df = xl.parse(s)
    print(f"\n--- {s} {df.shape}")
    if s == "Units":
        print("\n".join(str(v) for v in df.iloc[:, 0].tolist()))
    else:
        print(df.head(6).to_string(max_colwidth=40))
        print(df.tail(6).to_string(max_colwidth=40))

import SOUTH_AMERICA_COAL as m  # noqa: E402
for y, path in list(m.peru_anuarios().items())[:1]:
    r = m.get("https://www.gob.pe" + path)
    for f in re.findall(r'https://cdn\.www\.gob\.pe/uploads/document/file/\d+/[^"?#\s]+\.xlsx', r.text)[:1]:
        x = m.get(f)
        xl = pd.ExcelFile(io.BytesIO(x.content))
        for s in [s for s in xl.sheet_names if m.norm(s).lower().strip().startswith("produccion")]:
            df = xl.parse(s, header=None)
            print(f"\n--- Peru annex {s} {df.shape}")
            print(df.head(12).to_string(max_colwidth=22)[:4000])
            for i, row in df.iterrows():
                t = " | ".join(str(v) for v in row.values if str(v) != "nan")
                if re.search(r"carb", m.norm(t), re.I):
                    print(f"  r{i}: {t[:300]}")
