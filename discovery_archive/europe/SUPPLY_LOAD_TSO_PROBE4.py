"""Probe 4: SiStat Slovenia electricity balance (annual, 1817602S) - every measure, 2021-2025. Prints only."""
import csv
import io
import signal
import sys

import requests

signal.signal(signal.SIGALRM, lambda *_: (_ for _ in ()).throw(TimeoutError("hard timeout")))
signal.alarm(120)
U = "https://pxweb.stat.si/SiStatData/api/v1/en/Data/1817602S.px"
r = requests.post(U, json={"query": [], "response": {"format": "csv"}}, headers={"User-Agent": "Mozilla/5.0"}, timeout=(10, 60))
print(r.status_code, len(r.text), flush=True)
rows = list(csv.reader(io.StringIO(r.text)))
print(rows[0][-5:])
for row in rows[1:]:
    print(f"{row[0][:80]:82s}", row[-5:], flush=True)
sys.exit(0)
