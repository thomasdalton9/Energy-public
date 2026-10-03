"""
Shared ENTSO-E Transparency Platform helpers for the Europe pulls
(ENTSOE_GENERATION_DAILY.py, ENTSOE_CAPACITY.py, ENTSOE_PRICES_DAILY.py).

  - API key from the ENTSOE_API_KEY environment variable (GitHub secret) or api_keys.py
  - request() with retry/backoff; "no matching data" is an empty result, not an error;
    "too much data" windows are split in half by the callers via fetch_split()
  - parsers: generation per type (A75) -> MWh per UTC day, installed capacity (A68) -> MW per year,
    day-ahead prices (A44) -> mean EUR/MWh per UTC day
  - COUNTRIES: the bidding zones behind each country, and the psr-type -> column mappings

Direction convention (checked against Energy-Charts/SMARD/RTE for a settled week): in A75 a TimeSeries tagged
inBiddingZone_Domain is generation and one tagged outBiddingZone_Domain is the unit's own consumption (auxiliary
load, pumped-storage pumping, battery charging). Fuel columns are net (in - out); pumped storage and batteries
keep generation and consumption in separate columns.
"""
import os
import re
import sys
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone

import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root (api_keys, xlsx_notes)
try:
    import api_keys  # type: ignore  # noqa: F401
except Exception:  # noqa: BLE001
    api_keys = None

URL = "https://web-api.tp.entsoe.eu/api"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def api_key():
    k = os.environ.get("ENTSOE_API_KEY") or (getattr(api_keys, "ENTSOE_API_KEY", "") if api_keys else "")
    if not k:
        raise SystemExit("ENTSOE_API_KEY is not set (GitHub secret, or ENTSOE_API_KEY in api_keys.py). "
                         "Register at https://transparency.entsoe.eu and request API access.")
    return k


def redact(text):
    return re.sub(r"securityToken=[^&\s]+", "securityToken=***", str(text))


from europe_countries import COUNTRIES  # noqa: E402,F401  (country -> name, file slug, bidding zones)


# ENTSO-E psrType -> workbook column (generation). B10 / B25 are kept apart because pumping / charging is real
# consumption, not netting.
GEN_COLUMN = {
    "B11": "Hydro", "B12": "Hydro", "B10": "PumpedStorage",
    "B04": "Gas", "B03": "Gas",
    "B02": "Coal", "B05": "Coal", "B07": "Coal", "B08": "Coal",
    "B06": "Oil", "B14": "Nuclear", "B18": "Wind", "B19": "Wind", "B16": "Solar",
    "B01": "Bioenergy",
    "B09": "Other", "B13": "Other", "B15": "Other", "B17": "Other", "B20": "Other",
    "B25": "Storage",
}
GEN_FUELS = ["Hydro", "PumpedStorage", "Gas", "Coal", "Oil", "Nuclear", "Wind", "Solar", "Bioenergy", "Other", "Storage"]
# capacity columns follow add_charts.CAPACITY_FUELS; pumped storage / batteries are extra columns
CAP_COLUMN = {
    "B11": "Hydro", "B12": "Hydro", "B10": "PumpedStorage",
    "B04": "Gas", "B03": "Gas",
    "B02": "Coal", "B05": "Coal", "B07": "Coal", "B08": "Coal",
    "B06": "Oil", "B14": "Nuclear", "B18": "Wind", "B19": "Wind", "B16": "Solar",
    "B01": "Bioenergy",
    "B09": "Other", "B13": "Other", "B15": "Other", "B17": "Other", "B20": "Other",
    "B25": "Storage",
}
CAP_FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other", "PumpedStorage", "Storage"]


def stamp(d):
    return d.strftime("%Y%m%d%H%M")


class NoData(Exception):
    pass


class TooMuch(Exception):
    pass


# ENTSO-E allows 400 requests per minute per API key, shared by every job using it (generation, flows, capacity and
# prices can run at once), so each process keeps to about 100 a minute.
MIN_INTERVAL = float(os.environ.get("ENTSOE_MIN_INTERVAL", "0.6"))
_last_request = [0.0]


def _throttle():
    wait = _last_request[0] + MIN_INTERVAL - time.time()
    if wait > 0:
        time.sleep(wait)
    _last_request[0] = time.time()


def request(params, key=None, tries=5):
    """GET with retry. Returns the XML text. Raises NoData for 'no matching data', TooMuch when the platform says
    the window is too large, RuntimeError for anything else after the retries."""
    key = key or api_key()
    last = ""
    for attempt in range(tries):
        _throttle()
        try:
            r = requests.get(URL, params=dict(params, securityToken=key), headers=UA, timeout=(10, 180))
        except requests.RequestException as e:
            last = redact(f"{type(e).__name__}: {e}")[:200]
            time.sleep(5 * (attempt + 1))
            continue
        if r.status_code == 429 or r.status_code >= 500:
            last = f"HTTP {r.status_code}"
            time.sleep(20 * (attempt + 1))
            continue
        text = r.text
        if "<Acknowledgement_MarketDocument" in text[:600] or r.status_code == 400:
            reason = " ".join(re.findall(r"<text>(.*?)</text>", text, re.S))[:300]
            if "No matching data" in reason or "no data" in reason.lower():
                raise NoData(reason)
            if re.search(r"exceed|too many|too large|limit|maximum", reason, re.I):
                raise TooMuch(reason)
            raise RuntimeError(f"HTTP {r.status_code}: {redact(reason or text[:200])}")
        if r.status_code == 401:
            raise RuntimeError("HTTP 401: ENTSO-E rejected the API key")
        if r.status_code != 200:
            raise RuntimeError(f"HTTP {r.status_code}: {redact(text[:200])}")
        return text
    raise RuntimeError(f"failed after {tries} tries: {last}")


def fetch_split(params, d0, d1, parse, merge, min_days=1):
    """Fetch [d0, d1) and parse; if the platform says it's too much, split in half and merge. NoData -> None."""
    try:
        text = request(dict(params, periodStart=stamp(d0), periodEnd=stamp(d1)))
    except NoData:
        return None
    except TooMuch:
        if (d1 - d0).days <= min_days:
            raise
        mid = d0 + (d1 - d0) / 2
        mid = mid.replace(hour=0, minute=0)
        a = fetch_split(params, d0, mid, parse, merge, min_days)
        b = fetch_split(params, mid, d1, parse, merge, min_days)
        return merge(a, b)
    return parse(text)


def _strip(tag):
    return tag.split("}")[-1]


def _periods(root):
    """Yield (direction, psr, start, step_minutes, n_slots, values) per Period; values is a list of n_slots floats,
    None where the slot has no data."""
    for ts in root:
        if _strip(ts.tag) != "TimeSeries":
            continue
        kids = {_strip(c.tag): c for c in ts}
        direction = "in" if "inBiddingZone_Domain.mRID" in kids else "out" if "outBiddingZone_Domain.mRID" in kids else ""
        curve = kids["curveType"].text if "curveType" in kids else "A01"
        psr = None
        for e in ts.iter():
            if _strip(e.tag) == "psrType":
                psr = e.text
        for per in ts:
            if _strip(per.tag) != "Period":
                continue
            ti = {_strip(c.tag): c.text for c in next(c for c in per if _strip(c.tag) == "timeInterval")}
            res = next(c.text for c in per if _strip(c.tag) == "resolution")
            m = re.match(r"PT(\d+)([MH])", res)
            step = (int(m.group(1)) * (60 if m.group(2) == "H" else 1)) if m else {"P1D": 1440, "P7D": 10080, "P1Y": 525600}.get(res, 60)
            start = datetime.strptime(ti["start"], "%Y-%m-%dT%H:%MZ").replace(tzinfo=timezone.utc)
            end = datetime.strptime(ti["end"], "%Y-%m-%dT%H:%MZ").replace(tzinfo=timezone.utc)
            n = max(1, int(round((end - start).total_seconds() / 60 / step)))
            pts = {}
            for p in per:
                if _strip(p.tag) == "Point":
                    vals = {_strip(c.tag): c.text for c in p}
                    q = vals.get("quantity", vals.get("price.amount"))
                    if q is not None:
                        pts[int(vals["position"])] = float(q)
            yield direction, psr, start, step, n, _values(n, pts, curve)


def _values(n, pts, curve):
    """Per-slot values. Curve A03 (variable-sized blocks) omits points that repeat the previous one, so fill
    forward; A01 (sequential fixed blocks) has every slot it reports, and a missing slot is missing data."""
    if curve == "A03":
        out, last = [], None
        for pos in range(1, n + 1):
            if pos in pts:
                last = pts[pos]
            out.append(last)
        return out
    return [pts.get(pos) for pos in range(1, n + 1)]


def _daily(slots):
    """{key: {slot start: (value, hours)}} -> energy and covered hours per UTC day. A slot reported by more than one
    series (revisions, overlapping periods) counts once - the last one wins."""
    e, h = {}, {}
    for key, ts in slots.items():
        de, dh = e.setdefault(key, {}), h.setdefault(key, {})
        for t, (v, hrs) in ts.items():
            d = t.date()
            de[d] = de.get(d, 0.0) + v * hrs
            dh[d] = dh.get(d, 0.0) + hrs
    return e, h


def parse_generation(text):
    """A75 XML -> {"e": {(psr, direction): {date: MWh}}, "h": {(psr, direction): {date: hours with data}}} with each
    slot (MW x hours) summed into its UTC day. A slot reported twice counts once."""
    root = ET.fromstring(text)
    slots = {}
    for direction, psr, start, step, n, vals in _periods(root):
        if psr is None:
            continue
        d = slots.setdefault((psr, direction), {})
        for i, v in enumerate(vals):
            if v is not None:
                d[start + timedelta(minutes=step * i)] = (v, step / 60.0)
    e, h = _daily(slots)
    return {"e": e, "h": h}


def merge_generation(a, b):
    if a is None:
        return b
    if b is None:
        return a
    out = {"e": {k: dict(v) for k, v in a["e"].items()}, "h": {k: dict(v) for k, v in a["h"].items()}}
    for part in ("e", "h"):
        for k, v in b[part].items():
            o = out[part].setdefault(k, {})
            for d, x in v.items():
                o[d] = o.get(d, 0.0) + x
    return out


def parse_energy(text):
    """Load (A65) or cross-border flow (A11) XML -> {"e": {date: MWh}, "h": {date: hours with data}} per UTC day
    (series without a psrType: one value per slot, MW). A slot reported twice counts once."""
    root = ET.fromstring(text)
    slots = {"x": {}}
    for direction, psr, start, step, n, vals in _periods(root):
        for i, v in enumerate(vals):
            if v is not None:
                slots["x"][start + timedelta(minutes=step * i)] = (v, step / 60.0)
    e, h = _daily(slots)
    return {"e": e["x"], "h": h["x"]}


def merge_energy(a, b):
    if a is None:
        return b
    if b is None:
        return a
    out = {"e": dict(a["e"]), "h": dict(a["h"])}
    for part in ("e", "h"):
        for d, x in b[part].items():
            out[part][d] = out[part].get(d, 0.0) + x
    return out


def parse_capacity(text):
    """A68 XML -> {psr: {year: MW}} (installed capacity at the year in the period)."""
    root = ET.fromstring(text)
    out = {}
    for direction, psr, start, step, n, vals in _periods(root):
        v = next((x for x in vals if x is not None), None)
        if psr is None or v is None:
            continue
        out.setdefault(psr, {})[start.year] = out.setdefault(psr, {}).get(start.year, 0.0) + v
    return out


def merge_capacity(a, b):
    if a is None:
        return b
    if b is None:
        return a
    out = {k: dict(v) for k, v in a.items()}
    for k, v in b.items():
        out.setdefault(k, {}).update(v)
    return out


def parse_prices(text):
    """A44 XML -> {date: [price EUR/MWh, ...]} using only the finest resolution present on each day."""
    root = ET.fromstring(text)
    per_day = {}
    for direction, psr, start, step, n, vals in _periods(root):
        for i, v in enumerate(vals):
            if v is None:
                continue
            d = (start + timedelta(minutes=step * i)).date()
            per_day.setdefault(d, {}).setdefault(step, []).append(v)
    return {d: v[min(v)] for d, v in per_day.items()}


def merge_prices(a, b):
    if a is None:
        return b
    if b is None:
        return a
    out = dict(a)
    out.update(b)
    return out


def windows(d0, d1, days=60):
    """[d0, d1) split into consecutive windows of at most `days` days."""
    out, cur = [], d0
    while cur < d1:
        nxt = min(cur + timedelta(days=days), d1)
        out.append((cur, nxt))
        cur = nxt
    return out


def utc_midnight(d):
    return datetime(d.year, d.month, d.day, tzinfo=timezone.utc)


def today_utc():
    return datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
