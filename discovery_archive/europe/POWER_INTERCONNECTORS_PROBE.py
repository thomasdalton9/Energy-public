"""Probe: electricity interconnector audit. (1) Elexon FUELHH per GB interconnector, daily GWh (+ = import into GB), 2021-
saved to discovery_archive/europe/data/gb_interconnectors_daily.csv; (2) ENTSO-E A11 for candidate zone pairs the flows pull
does not query (Malta, Norway-Russia, Latvia-Belarus); (3) Energinet per-border exchanges; (4) Elering cross-border dashboard."""
import os, sys, time
from datetime import datetime, timedelta, timezone
import pandas as pd, requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, "europe")); sys.path.insert(0, ROOT)
OUT = os.path.join(HERE, "data"); os.makedirs(OUT, exist_ok=True)
ELEXON = "https://data.elexon.co.uk/bmrs/api/v1"
INT = ["INTFR", "INTIFA2", "INTELEC", "INTNED", "INTNEM", "INTVKL", "INTNSL", "INTIRL", "INTEW", "INTGRNL"]


def elexon():
    S = requests.Session(); S.headers["Accept"] = "application/json"
    d0 = datetime(2020, 12, 31).date(); end = datetime.now(timezone.utc).date()
    frames = []
    while d0 < end:
        d1 = min(d0 + timedelta(days=30), end)
        for t in range(5):
            try:
                r = S.get(f"{ELEXON}/datasets/FUELHH/stream", params={"settlementDateFrom": d0.isoformat(), "settlementDateTo": d1.isoformat(), "format": "json", "fuelType": INT}, timeout=180)
                if r.status_code == 200:
                    break
            except requests.RequestException:
                pass
            time.sleep(3 * (t + 1))
        rows = r.json() if r.status_code == 200 else []
        if rows:
            d = pd.DataFrame(rows)[["startTime", "fuelType", "generation", "publishTime"]]
            d["startTime"] = pd.to_datetime(d["startTime"], utc=True)
            d = d.sort_values("publishTime").drop_duplicates(["startTime", "fuelType"], keep="last")
            d["gwh"] = pd.to_numeric(d["generation"], errors="coerce") * 0.5 / 1000
            d["date"] = d["startTime"].dt.tz_localize(None).dt.normalize()
            frames.append(d.groupby(["date", "fuelType"])["gwh"].agg(["sum", "count"]).reset_index())
        else:
            print("elexon empty", d0, d1, r.status_code, flush=True)
        d0 = d1 + timedelta(days=1)
        time.sleep(0.3)
    a = pd.concat(frames)
    a = a.groupby(["date", "fuelType"]).agg({"sum": "sum", "count": "sum"}).reset_index()
    a = a[a["count"] >= 46]
    p = a.pivot(index="date", columns="fuelType", values="sum").round(3)
    p.to_csv(os.path.join(OUT, "gb_interconnectors_daily.csv"))
    y = p.groupby(p.index.year).sum() / 1000
    n = p.groupby(p.index.year).count()
    print("ELEXON annual TWh (+ = import into GB)\n", y.round(2).to_string(), "\ndays\n", n.to_string(), flush=True)


def entsoe():
    import entsoe_common as C
    import ENTSOE_FLOWS_DAILY as F
    C.api_key()
    zones = {"IT-Sicily": "10Y1001A1001A75E", "MT": "10Y1001A1001A93C", "NO4": "10YNO-4--------9", "RU": "10Y1001A1001A49F",
             "LV": "10YLV-1001A00074", "BY": "10Y1001A1001A51S", "NO1": "10YNO-1--------2", "SE3": "10Y1001A1001A46L",
             "IT-Sardinia": "10Y1001A1001A74G", "FR": "10YFR-RTE------C", "IT-Centre-North": "10Y1001A1001A70O",
             "SE1": "10Y1001A1001A44P", "FI": "10YFI-1--------U", "CY": "10YCY-1001A0003J", "GR": "10YGR-HTSO-----Y",
             "ES": "10YES-REE------0", "MA": "10YMA-ONE------O", "TR": "10YTR-TEIAS----W", "IT-South": "10Y1001A1001A788",
             "AL": "10YAL-KESH-----5", "IT-Centre-South": "10Y1001A1001A71M", "ME": "10YCS-CG-TSO---S", "MK": "10YMK-MEPSO----8", "XK": "10Y1001C--00100H"}
    pairs = [("IT-Sicily", "MT"), ("NO4", "RU"), ("LV", "BY"), ("IT-Sardinia", "FR"), ("IT-Centre-North", "FR"), ("GR", "CY"), ("ES", "MA"),
             ("NO1", "SE3"), ("IT-South", "AL"), ("ME", "MK")]
    end = C.today_utc(); st = datetime(2021, 1, 1, tzinfo=timezone.utc)
    dl = time.time() + 3000
    for a, b in pairs:
        for x, y in ((a, b), (b, a)):
            try:
                g = F.flow_days(zones[x], zones[y], st, end, dl)
            except Exception as e:
                print(f"A11 {x}>{y} FAILED {type(e).__name__}: {str(e)[:100]}", flush=True); continue
            if not g:
                print(f"A11 {x}>{y}: no data", flush=True); continue
            s = pd.Series({pd.Timestamp(k): v for k, v in g.items()})
            print(f"A11 {x}>{y}: " + ", ".join(f"{yr} {v / 1000:.2f}TWh/{c}d" for (yr, v), c in zip(s.groupby(s.index.year).sum().items(), s.groupby(s.index.year).count())), flush=True)


def energinet():
    """Energinet ElectricityBalanceNonv: hourly exchange per area (MWh, + = import into DK?). Monthly requests to avoid 429."""
    ds = "ElectricityBalanceNonv"
    rows = []
    for yr in range(2023, 2026):
        for mo in range(1, 13):
            a = datetime(yr, mo, 1); b = datetime(yr + (mo == 12), mo % 12 + 1, 1)
            if a > datetime.now():
                break
            for t in range(3):
                r = requests.get(f"https://api.energidataservice.dk/dataset/{ds}", params={"start": a.strftime("%Y-%m-%dT00:00"), "end": b.strftime("%Y-%m-%dT00:00"), "limit": 0, "columns": "HourUTC,PriceArea,ExchangeContinent,ExchangeGreatBelt,ExchangeNordicCountries,ExchangeGreatBritain"}, timeout=300)
                if r.status_code == 200:
                    break
                time.sleep(15); print('energinet', a, r.status_code, flush=True)
            recs = r.json().get("records") or [] if r.status_code == 200 else []
            if recs:
                rows.append(pd.DataFrame(recs))
            time.sleep(1.5)
    d = pd.concat(rows)
    d["date"] = pd.to_datetime(d["HourUTC"]).dt.normalize()
    ex = ["ExchangeContinent", "ExchangeGreatBelt", "ExchangeNordicCountries", "ExchangeGreatBritain"]
    for c in ex:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    g = d.groupby(["date", "PriceArea"])[ex].sum().reset_index()
    g.to_csv(os.path.join(OUT, "energinet_exchanges_daily.csv"), index=False)
    d["y"] = d["date"].dt.year
    print("ENERGINET exchange TWh by year and area\n", (d.groupby(["y", "PriceArea"])[ex].sum() / 1e6).round(2).to_string(), flush=True)


def elering():
    """Elering dashboard cross-border flows, 5-minute MW per link (Estlink 1/2, Finland total, Latvia, Russia Narva/Pihkva)."""
    frames = []
    for yr in range(2021, 2027):
        for half in (0, 1):
            a = datetime(yr, 1 + 6 * half, 1); b = datetime(yr + half, 1 + 6 * (1 - half), 1)
            if a > datetime.now():
                break
            r = requests.get("https://dashboard.elering.ee/api/transmission/cross-border", params={"start": a.strftime("%Y-%m-%dT00:00:00.000Z"), "end": b.strftime("%Y-%m-%dT00:00:00.000Z")}, timeout=300)
            if r.status_code != 200:
                print("ELERING", yr, half, r.status_code, flush=True); continue
            d = pd.DataFrame(r.json()["data"])
            frames.append(d)
            time.sleep(1)
    d = pd.concat(frames)
    d["date"] = pd.to_datetime(d["timestamp"], unit="s").dt.normalize()
    cols = [c for c in d.columns if c not in ("timestamp", "date")]
    daily = d.groupby("date")[cols].sum() * 5 / 60 / 1000   # MW x 5 min -> GWh
    cnt = d.groupby("date")["timestamp"].count()
    daily = daily[cnt >= 280]
    daily.to_csv(os.path.join(OUT, "elering_electricity_daily_gwh.csv"))
    print("ELERING GWh by year (sign as published)\n", (daily.groupby(daily.index.year).sum() / 1000).round(2).to_string(), flush=True)


if __name__ == "__main__":
    for f in [globals()[n] for n in (sys.argv[1:] or ['elexon', 'energinet', 'elering', 'entsoe'])]:
        try:
            f()
        except Exception as e:
            print(f.__name__, "FAILED", type(e).__name__, e, flush=True)
