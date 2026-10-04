"""CEE probe 2: NET4GAS CAMS public API (meters, long-range gasFlowPublished), AGGM + Amber Grid data requests. Writes discovery_archive/europe/probe_cee_out/*.txt"""
import json, os, re, sys
import requests
from playwright.sync_api import sync_playwright

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "probe_cee_out")
os.makedirs(OUT, exist_ok=True)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"

def w(name, txt):
    with open(os.path.join(OUT, name), "w") as f: f.write(txt)

def net4gas():
    L = []
    base = "https://extranet.cams.net4gas.cz"
    h = {"User-Agent": UA, "Content-Type": "application/json", "Accept": "application/json"}
    r = requests.get(base + "/usy-cams-gridmanagementg01-public/10800000000000000000000000000010/publicApi/interconnectionPoint/list",
                     params={"day": "2026-10-03T04:00:00.000Z", "pageInfo.pageIndex": 0, "pageInfo.pageSize": 1000}, headers=h, timeout=60)
    L.append(f"ipoint list {r.status_code} {len(r.text)}")
    try:
        for it in r.json()["itemList"]:
            L.append(json.dumps({k: it.get(k) for k in ("id", "meterId", "eic", "name", "directionCode", "parentInterconnectionPointId", "order")}))
    except Exception as e: L.append(f"{e} {r.text[:300]}")
    tech = base + "/usy-camsg01-technicalmodule-public/10100000000000000000000000000010/publicApi/tsData/list"
    def q(code, dims, a, b):
        body = {"tsList": [{"tsId": {"tsDefinitionCode": code, "tsDimensionValueIdList": dims}, "timeInterval": {"from": a, "to": b},
                            "tsValueListSampleType": "fixed", "loadLastPreviousFilledValue": "false"}], "pageInfo": {"pageIndex": 0, "pageSize": 1000}}
        return requests.post(tech, json=body, headers=h, timeout=90)
    for a, b in (("2026-09-30T04:00:00.000Z", "2026-10-03T04:00:00.000Z"), ("2021-01-01T05:00:00.000Z", "2026-10-03T04:00:00.000Z")):
        for dims in (["2", "exit"], ["2", "entry"]):
            r = q("gasFlowPublished_KwhD_day_meterD", dims, a, b)
            L.append(f"tsData {dims} {a[:10]}..{b[:10]}: {r.status_code} len={len(r.text)} {r.text[:1500]}")
    # try other meter ids for exit
    for m in range(1, 40):
        r = q("gasFlowPublished_KwhD_day_meterD", [str(m), "exit"], "2026-09-30T04:00:00.000Z", "2026-10-02T04:00:00.000Z")
        t = r.text
        mm = re.findall(r'"value":\s*"?([-0-9.eE]+)', t)
        L.append(f"meter {m} exit: {r.status_code} len={len(t)} values={mm[:3]} err={t[:120] if r.status_code != 200 else ''}")
    w("net4gas.txt", "\n".join(L))

def pw(label, url, wait, filt, fname, clicks=()):
    L = []
    with sync_playwright() as p:
        b = p.chromium.launch(); pg = b.new_context(user_agent=UA, viewport={"width": 1500, "height": 1000}).new_page()
        def on(r):
            if not re.search(filt, r.url): return
            try: body = r.text()[:1200].replace("\n", " ")
            except Exception: body = ""
            L.append(f"[{r.status}] {r.request.method} {r.url[:300]}\n  hdr={ {k: v for k, v in r.request.headers.items() if k.lower().startswith('x-power') or k.lower() in ('authorization','content-type')} }\n  post={(r.request.post_data or '')[:6000]}\n  resp={body}")
        pg.on("response", on)
        try:
            pg.goto(url, wait_until="domcontentloaded", timeout=60000); pg.wait_for_timeout(wait)
            for c in clicks:
                try:
                    pg.locator(c).first.click(timeout=4000); pg.wait_for_timeout(8000); L.append(f"clicked {c}")
                except Exception as e: L.append(f"click fail {c} {type(e).__name__}")
        except Exception as e: L.append(f"load problem {e}")
        L.append("URL " + pg.url); 
        try: L.append("TEXT " + pg.inner_text("body")[:1500].replace("\n", " | "))
        except Exception: pass
        for fr in pg.frames:
            L.append("FRAME " + fr.url[:200])
            try: L.append("FRAMETEXT " + fr.inner_text("body")[:1500].replace("\n", " | "))
            except Exception: pass
        b.close()
    w(fname, "\n".join(L))

for fn in (net4gas,):
    try: fn()
    except Exception as e: w("net4gas_err.txt", repr(e))
pw("aggm", "https://platform.aggm.at/portal/visualisation/ts-publication?src=map&granularity=hour&startDate=2026-09-01&Consumption=EndConsumer", 35000,
   r"aggm\.at/(vis|api|ts|portal).*", "aggm.txt")
pw("aggm2", "https://platform.aggm.at/portal/visualisation/ts-publication?fav=Qo1", 35000, r"aggm\.at/(vis|api|ts).*", "aggm_fav.txt")
pw("amber", "https://ambergrid.lt/en/for-clients/open-data/650", 45000, r"querydata|conceptualschema|modelsAndExploration|exploration", "amber.txt")
pw("amber_home", "https://ambergrid.lt/en", 45000, r"querydata|conceptualschema|modelsAndExploration|exploration", "amber_home.txt")
