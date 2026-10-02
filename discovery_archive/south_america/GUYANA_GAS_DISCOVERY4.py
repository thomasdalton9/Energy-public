"""
Round 4: the gas data-chart arrays end 30-Sep-2023 (page text claims to Jul-2026). Look for later gas data:
the data-chart REST posts (ACF/meta fields, modified dates), the other data-chart pages' date ranges (are oil pages
current?), and media-library uploads / posts mentioning gas, production reports, flaring, Gas-to-Energy.
"""
import json
import re

import requests

S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                                "Chrome/124.0 Safari/537.36"})
B = "https://petroleum.gov.gy"


def get(url, **kw):
    try:
        r = S.get(url, timeout=90, **kw)
        return r
    except Exception as e:   # noqa: BLE001
        print(f"  {url}: {e}", flush=True)
        return None


print("===== data-chart posts (REST)", flush=True)
r = get(f"{B}/wp-json/wp/v2/data-chart", params={"per_page": 100})
if r is not None:
    print(f"  status {r.status_code}", flush=True)
    try:
        for p in r.json():
            print(f"  id {p.get('id')} {p.get('slug')}: modified {p.get('modified')}, keys {sorted(p)[:30]}", flush=True)
            for k in ("acf", "meta"):
                if p.get(k):
                    print(f"    {k}: {json.dumps(p[k])[:800]}", flush=True)
            c = (p.get("content") or {}).get("rendered", "")
            print(f"    content {len(c):,} chars: {re.sub(r'<[^>]+>', ' ', c)[:300]!r}", flush=True)
    except Exception as e:   # noqa: BLE001
        print(f"  not JSON: {e}; {r.text[:300]!r}", flush=True)

print("\n===== date range of every data-chart page", flush=True)
for slug in ["gas-produced", "gas-injected-flared-and-used", "oil-production", "liza-unity-gold-and-payara-gold-oil-production",
             "water-injected", "water-produced-and-injected", "oil-prices"]:
    r = get(f"{B}/data-chart/{slug}/")
    if r is None:
        continue
    t = r.text
    desc = re.search(r'<meta name="description" content="([^"]+)"', t)
    for m in re.finditer(r"var\s+(\w+)\s*=\s*'(\[.*?\])'\s*;", t, re.S):
        try:
            rows = json.loads(m.group(2))
        except Exception:   # noqa: BLE001
            continue
        ms = [x.get("Data") for x in rows if isinstance(x, dict) and isinstance(x.get("Data"), (int, float))]
        last = rows[-1] if rows else {}
        import datetime as dt
        rng = (f"{dt.datetime.utcfromtimestamp(min(ms) / 1000):%Y-%m-%d} to "
               f"{dt.datetime.utcfromtimestamp(max(ms) / 1000):%Y-%m-%d}") if ms else "no Data field"
        print(f"  {slug} var {m.group(1)}: {len(rows)} rows, {rng}; keys {list(last)[:12]}; last {json.dumps(last)[:300]}",
              flush=True)
    print(f"  {slug} description: {desc.group(1) if desc else None}", flush=True)

print("\n===== media library and posts mentioning gas / production", flush=True)
for kind in ("media", "posts", "pages"):
    for q in ("gas", "production report", "flared", "Gas to Energy", "monthly production", "Stabroek"):
        r = get(f"{B}/wp-json/wp/v2/{kind}", params={"search": q, "per_page": 50, "orderby": "date", "order": "desc"})
        if r is None or r.status_code != 200:
            print(f"  {kind} '{q}': {None if r is None else r.status_code}", flush=True)
            continue
        try:
            items = r.json()
        except Exception:   # noqa: BLE001
            continue
        print(f"  {kind} '{q}': {len(items)} (total {r.headers.get('X-WP-Total')})", flush=True)
        for it in items[:25]:
            title = re.sub(r"<[^>]+>", "", (it.get("title") or {}).get("rendered", ""))
            print(f"    {it.get('date', '')[:10]} {title[:90]!r} {it.get('source_url') or it.get('link')}", flush=True)
