"""
Probe 2: Spain REData physical exchanges (all borders incl. Morocco and Andorra, which ENTSO-E's Spanish flows lack), Elia open data annual totals
(load, generation by fuel, physical flows by border), SiStat / EMS / Litgrid link discovery. Prints only.
"""
import json
import re
import signal
import sys
import urllib.parse

import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "application/json, text/html, */*"}


def _timeout(*_):
    raise TimeoutError("hard timeout")


signal.signal(signal.SIGALRM, _timeout)


def get(url, show=0):
    try:
        signal.alarm(60)
        r = requests.get(url, headers=UA, timeout=(10, 30))
        signal.alarm(0)
        print(f"GET {url[:170]} -> {r.status_code} {len(r.content)}b", flush=True)
        if show:
            print("   ", r.text[:show].replace("\n", " ").replace("\r", " "), flush=True)
        return r
    except BaseException as e:  # noqa: BLE001
        signal.alarm(0)
        print(f"GET {url[:170]} -> ERR {str(e)[:100]}", flush=True)


def jget(url):
    r = get(url)
    try:
        return r.json() if r is not None and r.ok else None
    except Exception:  # noqa: BLE001
        return None


print("=== SPAIN REData exchanges (TWh per year)")
for ep in ("todas-fronteras-fisicos", "todas-fronteras-programados", "francia-frontera", "portugal-frontera", "marruecos-frontera", "andorra-frontera",
           "enlaces-francia", "enlaces-portugal", "enlaces-marruecos", "enlaces-andorra"):
    for y in (2023, 2024, 2025):
        j = jget(f"https://apidatos.ree.es/en/datos/intercambios/{ep}?start_date={y}-01-01T00:00&end_date={y}-12-31T23:59&time_trunc=year")
        if not j:
            continue
        for grp in j.get("included", []):
            vals = grp["attributes"].get("values") or []
            tot = sum(v["value"] for v in vals) if vals else None
            if tot is None and grp["attributes"].get("content"):
                for it in grp["attributes"]["content"]:
                    v = it["attributes"].get("values") or []
                    print(f"   {ep} {y} {grp['type']} / {it['type']}: {sum(x['value'] for x in v)/1e6:.3f} TWh")
            else:
                print(f"   {ep} {y} {grp['type']}: {tot/1e6 if tot is not None else None} TWh")
j = jget("https://apidatos.ree.es/en/datos/balance/balance-electrico?start_date=2025-01-01T00:00&end_date=2025-12-31T23:59&time_trunc=year")
if j:
    for grp in j["included"]:
        for it in grp["attributes"]["content"]:
            t = it["attributes"].get("total")
            if t is not None:
                print(f"   bal2025 {it['type'][:40]:42s} {t/1e6:.3f}")

print("\n=== ELIA annual totals")
B = "https://opendata.elia.be/api/explore/v2.1/catalog/datasets"
for ds, field in (("ods001", "totalload"), ("ods003", "eliagridload")):
    q = urllib.parse.urlencode({"select": f"year(datetime) as y, sum({field}) as s, count(*) as n", "group_by": "y", "order_by": "y", "limit": 20})
    j = jget(f"{B}/{ds}/records?{q}")
    print("  ", ds, field, json.dumps(j)[:700] if j else None)
for ds in ("ods201", "ods026", "ods124", "ods160"):
    j = jget(f"{B}/{ds}/records?limit=2")
    print("  ", ds, json.dumps(j)[:600] if j else None)
    if not j or not j.get("results"):
        continue
    rec = j["results"][0]
    strs = [k for k, v in rec.items() if isinstance(v, str) and k not in ("datetime", "resolutioncode")]
    nums = [k for k, v in rec.items() if isinstance(v, (int, float))]
    print("      string keys", strs, "numeric keys", nums)
    for num in nums[:2]:
        sel = f"year(datetime) as y, {', '.join(strs[:1])}, sum({num}) as s, count(*) as n" if strs else f"year(datetime) as y, sum({num}) as s, count(*) as n"
        gb = f"y,{strs[0]}" if strs else "y"
        q = urllib.parse.urlencode({"select": sel, "group_by": gb, "order_by": "y", "where": "datetime>='2023-01-01'", "limit": 100})
        jj = jget(f"{B}/{ds}/records?{q}")
        if jj:
            for row in jj["results"]:
                print("        ", num, row)

print("\n=== SiStat electricity tables")
r = get("https://pxweb.stat.si/SiStatData/api/v1/en/Data/")
if r is not None and r.ok:
    for t in r.json():
        if re.search(r"electric|electr", t.get("text", ""), re.I):
            print("  ", t["id"], t["text"][:110], t.get("updated"))

print("\n=== link discovery")
for u, pat in (("https://ems.rs/en/", r"bilans|balance|annual|report|potro|consum|data|statist|xls|trans"),
               ("https://www.litgrid.eu/", r"xls|csv|balans|load|sistem|energy|data|api|statist|suvart|gamyb"),
               ("https://www.eles.si/", r"xls|csv|poraba|podatki|data|transp|bilanc|izgub|letno")):
    r = get(u)
    if r is not None and r.ok:
        seen = set()
        for m in re.findall(r'href=["\']?([^\s"\'>]+)', r.text):
            if m not in seen and re.search(pat, m, re.I):
                seen.add(m)
                print("      link", m[:150])
sys.exit(0)
