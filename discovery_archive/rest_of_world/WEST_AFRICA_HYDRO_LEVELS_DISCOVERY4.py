"""
Follow-up 3: Bui site-content.json water_level_m block; Hydroweb.next West African lake items
(span, download keyless?) and the file format of L_volta / L_kainji / L_shiroro / L_lagdo.
Writes west_africa_hydro_levels_output4.txt.
"""
import json
import re
import signal
import requests

OUT = open("west_africa_hydro_levels_output4.txt", "w", encoding="utf-8")


def log(*a):
    s = " ".join(str(x) for x in a)
    print(s, flush=True)
    OUT.write(s + "\n")
    OUT.flush()


class Hard(Exception):
    pass


def _alarm(*a):
    raise Hard("hard deadline")


signal.signal(signal.SIGALRM, _alarm)
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def get(url, **kw):
    try:
        signal.alarm(60)
        r = requests.get(url, headers=H, timeout=(10, 40), **kw)
        signal.alarm(0)
        return r
    except BaseException as e:
        signal.alarm(0)
        log("  ERR", url, type(e).__name__, str(e)[:150])
        return None


log("## Bui site-content.json water_level / plant blocks")
r = get("https://www.buipower.com/data/site-content.json")
if r is not None and r.ok:
    j = r.json()
    log(" top keys:", list(j.keys()))
    for k, v in j.items():
        if re.search(r"(?i)water|level|reservoir|live|plant|stat|generation", k):
            log("  KEY", k, json.dumps(v)[:1200])
    def walk(o, path=""):
        if isinstance(o, dict):
            for k, v in o.items():
                if re.search(r"(?i)water_level|updated_at|history|reservoir", k):
                    log("  found", path + "/" + k, json.dumps(v)[:500])
                walk(v, path + "/" + k)
        elif isinstance(o, list):
            for i, v in enumerate(o[:50]):
                walk(v, path + f"[{i}]")
    walk(j)

log("## Hydroweb lake items")
base = "https://hydroweb.next.theia-land.fr/api/v1/rs-catalog/stac/search"
for coll in ["HYDROWEB_LAKES_OPE", "HYDROWEB_LAKES_RESEARCH"]:
    r = get(f"{base}?collections={coll}&bbox=-14,4,15,17&limit=100")
    if r is None or not r.ok:
        continue
    fs = r.json().get("features", [])
    log(coll, "n =", len(fs))
    for f in fs:
        p = f["properties"]
        nm = f["id"].split("@")[0]
        log("  ", nm, p.get("start_datetime"), "->", p.get("end_datetime"))
    for f in fs:
        nm = f["id"].split("@")[0]
        if nm in ("L_volta", "L_kainji", "L_shiroro", "L_bui") and coll == "HYDROWEB_LAKES_OPE" or nm == "L_bui":
            for a in f["assets"].values():
                d = get(a["href"])
                if d is None:
                    continue
                log(f"  DOWNLOAD {nm}: {d.status_code} {d.headers.get('content-type')} {len(d.content)}")
                log("   head:", d.text[:900].replace("\n", " | "))
                log("   tail:", d.text[-300:].replace("\n", " | "))
log("Done")
OUT.close()
