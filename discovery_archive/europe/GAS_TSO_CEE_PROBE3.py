"""CEE probe 3: AGGM favourite structure, Amber Grid Power BI exploration/schema (raw), NET4GAS allocation scan. Writes probe_cee_out/*"""
import json, os, re
import requests

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "probe_cee_out")
os.makedirs(OUT, exist_ok=True)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
def w(n, t): open(os.path.join(OUT, n), "w").write(t)

def shape(o, depth=0):
    if isinstance(o, dict): return {k: shape(v, depth + 1) for k, v in o.items()}
    if isinstance(o, list):
        return [f"list[{len(o)}]"] + ([shape(o[0], depth + 1)] if o else [])
    return o if not isinstance(o, str) else o[:80]

def aggm():
    r = requests.get("https://platform.aggm.at/vis-service/api/ts/favorite/Qo1", headers={"User-Agent": UA}, timeout=60)
    j = r.json()
    w("aggm_fav_shape.json", json.dumps(shape(j), indent=1, ensure_ascii=False))
    w("aggm_fav_full_head.json", r.text[:200000])
    r = requests.get("https://platform.aggm.at/vis-service/api/ts/attributes", headers={"User-Agent": UA}, timeout=60)
    w("aggm_attributes.json", r.text)

def amber():
    base = "https://wabi-west-europe-d-primary-api.analysis.windows.net/public/reports"
    for rid in ("305e13f3-56da-4cce-813c-92a5ac0181e1", "8542baae-0515-423f-b88c-ceb11a2b056f"):
        h = {"User-Agent": UA, "X-PowerBI-ResourceKey": rid, "Accept": "application/json"}
        r = requests.get(f"{base}/{rid}/modelsAndExploration?preferReadOnlySession=true", headers=h, timeout=60)
        w(f"amber_{rid[:4]}_exploration.json", r.text)
        try:
            mid = r.json()["models"][0]["id"]
            r2 = requests.post(f"{base}/conceptualschema", json={"modelIds": [mid]}, headers=h, timeout=60)
            w(f"amber_{rid[:4]}_schema.json", r2.text)
        except Exception as e: w(f"amber_{rid[:4]}_err.txt", repr(e) + r.text[:300])

def net4gas():
    h = {"User-Agent": UA, "Content-Type": "application/json"}
    bal = "https://extranet.cams.net4gas.cz/usy-cams-balancingg01-public/11100000000000000000000000000010/publicApi/tsData/list"
    L = []
    for code, dirs in (("allocation_KwhD_day_ipoD", ("entry", "exit")),):
        for d in dirs:
            for ipo in range(1, 80):
                body = {"tsList": [{"tsId": {"tsDefinitionCode": code, "tsDimensionValueIdList": [str(ipo), d]}, "timeInterval": {"from": "2025-01-01T05:00:00.000Z", "to": "2025-01-08T05:00:00.000Z"},
                                    "tsValueListSampleType": "fixed", "loadLastPreviousFilledValue": "false"}], "pageInfo": {"pageIndex": 0, "pageSize": 1000}}
                r = requests.post(bal, json=body, headers=h, timeout=60)
                vals = re.findall(r'"value":\s*(-?[0-9.eE+]+)', r.text)
                if vals and any(float(v) != 0 for v in vals):
                    L.append(f"{code} ipo={ipo} {d}: n={len(vals)} sum={sum(float(v) for v in vals):.0f} first={vals[:3]}")
                elif r.status_code != 200: L.append(f"ipo={ipo} {d} HTTP {r.status_code} {r.text[:150]}")
    # time series definitions?
    for p in ("tsDefinition/list", "tsDefinition/get"):
        try:
            r = requests.get("https://extranet.cams.net4gas.cz/usy-camsg01-technicalmodule-public/10100000000000000000000000000010/publicApi/" + p, headers=h, timeout=30)
            L.append(f"{p}: {r.status_code} {r.text[:600]}")
        except Exception as e: L.append(f"{p}: {e}")
    w("net4gas_alloc.txt", "\n".join(L))

for fn in (aggm, amber, net4gas):
    try: fn()
    except Exception as e: w(f"err_{fn.__name__}.txt", repr(e))
