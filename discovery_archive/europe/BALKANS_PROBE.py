"""
Balkans coverage probe (follow-up to EU_KEYED_PROBE.py; reuses its ENTSO-E helpers).

For Montenegro, Serbia, Bosnia-Herzegovina, North Macedonia, Albania, Kosovo
(+ Croatia, Slovenia, Bulgaria for comparison):
  - ENTSO-E actual generation per type: fuel types present, resolution, lag of last point (last 3 days)
  - settled week 2026-05-04..10 by fuel group (GWh) - zero/missing fuels show where the feed is thin
  - history depth: is there data in a 2022 and a 2019 sample week?
  - ENTSO-E load (A65) coverage as a fallback signal
  - Energy-Charts public_power for me/rs/ba/mk/al/xk

Usage: python3 BALKANS_PROBE.py
"""
import os
import sys
import time
from datetime import datetime, timedelta, timezone

import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import EU_KEYED_PROBE as K  # noqa: E402

ZONES = {
    "ME": "10YCS-CG-TSO---S", "RS": "10YCS-SERBIATSOV", "BA": "10YBA-JPCC-----D", "MK": "10YMK-MEPSO----8",
    "AL": "10YAL-KESH-----5", "XK": "10Y1001C--00100H", "HR": "10YHR-HEP------M", "SI": "10YSI-ELES-----O",
    "BG": "10YCA-BULGARIA-R",
}
ALL_GEN = ["B01", "B02", "B03", "B04", "B05", "B06", "B07", "B08", "B09", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "B17", "B18", "B19", "B20"]


def day_gen(zone, d):
    res = K.entsoe_gen(zone, d, d + timedelta(days=1))
    return res


def main():
    if not K.ENTSOE:
        K.log("ENTSOE_API_KEY not set")
        return
    now = datetime.now(timezone.utc).replace(minute=0, second=0, microsecond=0)
    K.log("== ENTSO-E generation, last 3 days ==")
    for name, z in ZONES.items():
        res = K.entsoe_gen(z, now - timedelta(days=3), now)
        if isinstance(res, tuple):
            K.log(f"  {name}: {res[0]} {str(res[1])[:140]}")
        elif not res:
            K.log(f"  {name}: empty")
        else:
            types = sorted({k.split(":")[0] for k in res})
            dirs = sorted({k.split(":")[1] for k in res})
            last = max(v["last"] for v in res.values())
            K.log(f"  {name}: types={types} dir={dirs} res={sorted({v['res'] for v in res.values()})}min "
                  f"last={last:%Y-%m-%d %H:%M} lag_h={(now - last).total_seconds() / 3600:.0f}")
        time.sleep(0.4)

    K.log("\n== settled week 2026-05-04..10: GWh by ENTSO-E psr type (generation only, both directions summed) ==")
    for name, z in ZONES.items():
        tot, ok = {}, 0
        d = K.W0
        while d < K.W1:
            res = day_gen(z, d)
            if not isinstance(res, tuple):
                ok += 1
                for k, v in res.items():
                    psr = k.split(":")[0]
                    tot[psr] = tot.get(psr, 0.0) + v["mwh"]
            d += timedelta(days=1)
            time.sleep(0.3)
        txt = ", ".join(f"{p} {v / 1000:.1f}" for p, v in sorted(tot.items()) if p in ALL_GEN)
        K.log(f"  {name}: days_ok={ok}/7 | {txt or 'no data'}")

    K.log("\n== history depth (one-day samples) ==")
    for name, z in ZONES.items():
        out = []
        for label, d in (("2022-06-01", datetime(2022, 6, 1, tzinfo=timezone.utc)), ("2019-06-01", datetime(2019, 6, 1, tzinfo=timezone.utc))):
            res = day_gen(z, d)
            out.append(f"{label}: " + (f"{res[0]}" if isinstance(res, tuple) else f"{len({k.split(':')[0] for k in res})} types"))
            time.sleep(0.3)
        K.log(f"  {name}: " + "; ".join(out))

    K.log("\n== ENTSO-E actual load (A65), last 3 days ==")
    for name, z in ZONES.items():
        r, err = K.entsoe_get({"documentType": "A65", "processType": "A16", "outBiddingZone_Domain": z,
                               "periodStart": K.stamp(now - timedelta(days=3)), "periodEnd": K.stamp(now)})
        if r is None or r.status_code != 200:
            K.log(f"  {name}: {err or r.status_code}")
            continue
        res = K.parse_series(r.text, want_generation=False)
        K.log(f"  {name}: " + (f"no data ({res[1][:80]})" if isinstance(res, tuple) else f"points={ {k: (v['points'], v['res']) for k, v in res.items()} }"))
        time.sleep(0.3)

    K.log("\n== Energy-Charts public_power ==")
    d1 = (now - timedelta(days=1)).strftime("%Y-%m-%d")
    d0 = (now - timedelta(days=8)).strftime("%Y-%m-%d")
    for cc in ("me", "rs", "ba", "mk", "al", "xk", "hr", "bg"):
        for _ in range(3):
            r = requests.get("https://api.energy-charts.info/public_power", params={"country": cc, "start": d0, "end": d1}, headers=K.UA, timeout=(10, 60))
            if r.status_code == 429:
                time.sleep(20)
                continue
            break
        if r.status_code == 200:
            d = r.json()
            names = [s["name"] if isinstance(s["name"], str) else s["name"].get("en") for s in d["production_types"]]
            K.log(f"  {cc}: 200 series={len(names)} points={len(d.get('unix_seconds', []))} {names}")
        else:
            K.log(f"  {cc}: {r.status_code} {r.text[:100]}")
        time.sleep(5)


if __name__ == "__main__":
    main()
