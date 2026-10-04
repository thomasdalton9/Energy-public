"""
Probe 3: Elia annual totals (load, generation by fuel, physical border flows), SiStat electricity balance tables, TSO page scans
(EMS Serbia, Litgrid, ELES). Prints only.
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


def get(url, post=None):
    try:
        signal.alarm(70)
        r = requests.post(url, json=post, headers=UA, timeout=(10, 40)) if post is not None else requests.get(url, headers=UA, timeout=(10, 40))
        signal.alarm(0)
        print(f"{'POST' if post is not None else 'GET'} {url[:130]} -> {r.status_code} {len(r.content)}b" + ("" if r.ok else " " + r.text[:200].replace("\n", " ")), flush=True)
        return r
    except BaseException as e:  # noqa: BLE001
        signal.alarm(0)
        print(f"GET {url[:130]} -> ERR {str(e)[:100]}", flush=True)


B = "https://opendata.elia.be/api/explore/v2.1/catalog/datasets"
print("=== ELIA (MWh = MW x 0.25 for PT15M; TWh)")
for y in (2023, 2024, 2025):
    where = f"datetime>=date'{y}-01-01' and datetime<date'{y + 1}-01-01'"
    r = get(f"{B}/ods001/records?" + urllib.parse.urlencode({"select": "sum(totalload) as s, count(*) as n", "where": where}))
    if r is not None and r.ok:
        row = r.json()["results"][0]
        print(f"  {y} Elia total load (ods001): {row['s'] / 4e6:.3f} TWh, {row['n']} slots")
    r = get(f"{B}/ods201/records?" + urllib.parse.urlencode({"select": "fueltypeentsoe, sum(generatedpower) as s, count(*) as n", "group_by": "fueltypeentsoe", "where": where, "limit": 50}))
    if r is not None and r.ok:
        tot = 0
        for row in r.json()["results"]:
            print(f"  {y} gen {row['fueltypeentsoe']:28s} {row['s'] / 4e6:8.3f} TWh ({row['n']})")
            tot += row["s"]
        print(f"  {y} gen total {tot / 4e6:.3f} TWh")
    r = get(f"{B}/ods026/records?" + urllib.parse.urlencode({"select": "controlarea, sum(physicalflowatborder) as s, count(*) as n", "group_by": "controlarea", "where": where, "limit": 50}))
    if r is not None and r.ok:
        for row in r.json()["results"]:
            print(f"  {y} flow {row['controlarea']:20s} {(row['s'] or 0) / 4e6:8.3f} TWh ({row['n']})")

print("\n=== SiStat 1817602S (annual) / 1817601S (monthly) electricity balance")
for t in ("1817602S.px", "1817601S.px"):
    r = get(f"https://pxweb.stat.si/SiStatData/api/v1/en/Data/{t}")
    if r is not None and r.ok:
        try:
            meta = r.json()
            for v in meta["variables"]:
                print("   var", v["code"], v["text"], "|", "; ".join(f"{c}={x}" for c, x in zip(v["values"], v["valueTexts"]))[:700])
        except Exception as e:  # noqa: BLE001
            print("   meta parse", e, r.text[:300])
        r = get(f"https://pxweb.stat.si/SiStatData/api/v1/en/Data/{t}", post={"query": [], "response": {"format": "csv"}})
        if r is not None and r.ok:
            lines = r.text.splitlines()
            print("   rows", len(lines))
            for ln in lines[:2] + [l for l in lines[2:] if re.search(r"202[3-5]", l)][:80]:
                print("   ", ln[:220])

print("\n=== page scans")
for u in ("https://ems.rs/en/transparency-2/", "https://www.litgrid.eu/sistema/elektros-energetikos-sistema/elektros-gamybos-ir-vartojimo-balanso-duomenys",
          "https://www.litgrid.eu/sistemos-duomenys", "https://www.eles.si/trzni-podatki", "https://www.eles.si/osnovni-podatki",
          "https://www.eles.si/kljucni-podatki-o-poslovanju"):
    r = get(u)
    if r is not None and r.ok:
        seen = set()
        for m in re.findall(r'href=["\']?([^\s"\'>]+)', r.text):
            if m not in seen and re.search(r"xls|csv|\.pdf|api|json|download|Portals|uploads", m, re.I) and "icon" not in m and ".png" not in m:
                seen.add(m)
                print("      link", m[:160])
        txt = re.sub(r"<[^>]+>", " ", re.sub(r"<script.*?</script>|<style.*?</style>", " ", r.text, flags=re.S))
        for m in re.finditer(r"[^.]{0,80}(izgub|losses|nuostol|perdavimo nuostol|consumption|poraba|vartojim)[^.]{0,100}", re.sub(r"\s+", " ", txt), re.I):
            print("      txt", m.group(0)[:200])
            break
sys.exit(0)
