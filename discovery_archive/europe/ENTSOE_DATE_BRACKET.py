"""
Where does ENTSO-E have generation data? The 2026-05-04..10 week returned
"No matching data found" for DE-LU, FR, ES, IT-North, PL, NL, BE, AT, CZ, PT
even though the same zones return fresh data for the last 3 days (and FR 2025,
GB 2019 returned data). This asks for single days across ~18 months for a few
big zones, with Montenegro (which did return May data) as a control.

Usage: python3 ENTSOE_DATE_BRACKET.py
"""
import os
import sys
import time
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import EU_KEYED_PROBE as K  # noqa: E402

ZONES = {k: K.ZONES[k] for k in ("DE-LU", "FR", "ES", "NL")}
ZONES["ME(control)"] = "10YCS-CG-TSO---S"
DAYS = ["2026-09-28", "2026-09-20", "2026-09-10", "2026-09-01", "2026-08-15", "2026-07-15", "2026-06-15", "2026-06-01",
        "2026-05-20", "2026-05-12", "2026-05-04", "2026-04-15", "2026-03-15", "2026-01-15", "2025-10-15", "2025-06-15"]


def main():
    if not K.ENTSOE:
        K.log("ENTSOE_API_KEY not set")
        return
    K.log("day        | " + " | ".join(f"{z:>11}" for z in ZONES))
    for day in DAYS:
        d = datetime.strptime(day, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        cells = []
        for z in ZONES.values():
            res = K.entsoe_gen(z, d, d + timedelta(days=1))
            cells.append(f"{res[0]:>11}" if isinstance(res, tuple) else f"{len({k.split(':')[0] for k in res}):>2} types   ".rjust(11))
            time.sleep(0.3)
        K.log(f"{day} | " + " | ".join(cells))


if __name__ == "__main__":
    main()
