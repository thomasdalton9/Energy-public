"""
Probe: Balkan cross-border flows, ENTSO-E physical (A11) vs finalised schedules (A09), both directions, GWh per year 2023-2025,
for the Bosnia-Montenegro border and Montenegro's other borders (the BA-ME physical flow does not close either side's balance).
Prints only.
"""
import os
import sys
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)).replace("discovery_archive/europe", "europe"))
import entsoe_common as C  # noqa: E402

E = {"BA": "10YBA-JPCC-----D", "ME": "10YCS-CG-TSO---S", "RS": "10YCS-SERBIATSOV", "AL": "10YAL-KESH-----5", "XK": "10Y1001C--00100H",
     "MK": "10YMK-MEPSO----8", "HR": "10YHR-HEP------M", "IT-CS": "10Y1001A1001A71M"}
PAIRS = [("BA", "ME"), ("ME", "BA"), ("RS", "ME"), ("ME", "RS"), ("AL", "ME"), ("ME", "AL"), ("XK", "ME"), ("ME", "XK"), ("IT-CS", "ME"), ("ME", "IT-CS"),
         ("HR", "BA"), ("BA", "HR"), ("RS", "BA"), ("BA", "RS")]


def annual(doc, a, b, y):
    tot, days = 0.0, 0
    d0 = date(y, 1, 1)
    d1 = min(date(y + 1, 1, 1), C.today_utc().date())
    for w0, w1 in C.windows(C.utc_midnight(d0), C.utc_midnight(d1), 60):
        try:
            r = C.fetch_split({"documentType": doc, "out_Domain": E[a], "in_Domain": E[b]}, w0, w1, C.parse_energy, C.merge_energy)
        except Exception as e:  # noqa: BLE001
            return f"ERR {str(e)[:40]}"
        if r:
            tot += sum(r["e"].values()) / 1000.0
            days += len(r["e"])
    return f"{tot:7.0f} ({days}d)"


print("pair      | year | A11 physical      | A09 schedules", flush=True)
for a, b in PAIRS:
    for y in (2023, 2024, 2025):
        print(f"{a}>{b:6} | {y} | {annual('A11', a, b, y)} | {annual('A09', a, b, y)}", flush=True)
sys.exit(0)
