"""
One-off probe (prints only): how far back do AEMO's NEMWEB folders reach for the DWGM (INT041) and Wallumbilla
benchmark prices that AU_GAS_HUB_PRICES.py reads? Decides whether au_sttm_prices.yml can move from daily to the
1st/15th schedule (the longest gap between those runs is 17 days).
"""
import io
import os
import re
import sys
import zipfile

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "australia_nz"))
from AU_GAS_HUB_PRICES import GSH, NEM, VIC, aemo_csv, dwgm_frame, get  # noqa: E402

LIST_RE = r'([A-Za-z]+, [A-Za-z]+ \d+, \d{4}\s+[\d:]+ [AP]M)\s+(\d+|&lt;dir&gt;)\s+<a href="([^"]+)"'


def section(t):
    print("\n" + "=" * 25, t, "=" * 25, flush=True)


def listing(url):
    rows = re.findall(LIST_RE, get(url).text, re.I)
    return [(h.split("/")[-1], size, when) for when, size, h in rows]


def rng(f, label):
    print(f"   {label}: {len(f)} gas days {f.index.min():%Y-%m-%d}..{f.index.max():%Y-%m-%d}" if len(f) else f"   {label}: empty",
          flush=True)


section("DWGM: CURRENT/VicGas")
vic = listing(VIC)
print(f"   {len(vic)} entries")
for e in [e for e in vic if re.search(r"int041|PublicRpts", e[0], re.I)]:
    print("   ", e)
rng(dwgm_frame(pd.read_csv(io.BytesIO(get(VIC + "int041_v4_market_and_reference_prices_1.csv").content))), "rolling CSV")
allparts = []
for i in range(1, 15):
    try:
        z = zipfile.ZipFile(io.BytesIO(get(f"{VIC}PublicRpts{i:02d}.zip").content))
    except Exception as e:  # noqa: BLE001
        print(f"   PublicRpts{i:02d}: {type(e).__name__}")
        continue
    parts = [pd.read_csv(z.open(n)) for n in z.namelist() if n.lower().startswith("int041")]
    if parts:
        f = dwgm_frame(pd.concat(parts, ignore_index=True))
        allparts += parts
        rng(f, f"PublicRpts{i:02d} ({len(parts)} INT041 files)")
    else:
        print(f"   PublicRpts{i:02d}: no INT041; e.g. {z.namelist()[:3]}")
if allparts:
    rng(dwgm_frame(pd.concat(allparts, ignore_index=True)), "ALL PublicRpts combined")

section("Wallumbilla: GSH/Benchmark_Price")
gsh = listing(GSH)
files = [e for e in gsh if re.search(r"\.(zip|csv)$", e[0], re.I)]
print(f"   {len(files)} files")
for e in files[:3] + [("...", "", "")] + files[-3:]:
    print("   ", e)
for name, _, _ in files[:1] + files[-1:]:
    content = get(GSH + name).content
    frames = []
    if name.lower().endswith(".zip"):
        z = zipfile.ZipFile(io.BytesIO(content))
        frames = [aemo_csv(z.read(n)) for n in z.namelist()]
    else:
        frames = [aemo_csv(content)]
    d = pd.concat(frames, ignore_index=True)
    d.columns = [str(c).strip().upper() for c in d.columns]
    dcol = next((c for c in d.columns if "GAS_DATE" in c or c.endswith("DATE")), None)
    dt = pd.to_datetime(d[dcol].astype(str).str.strip('"').str[:10], format="%Y/%m/%d", errors="coerce").dropna()
    print(f"   {name}: {len(d)} rows, {dcol} {dt.min():%Y-%m-%d}..{dt.max():%Y-%m-%d}" if len(dt) else f"   {name}: no dates, cols {list(d.columns)[:10]}")
