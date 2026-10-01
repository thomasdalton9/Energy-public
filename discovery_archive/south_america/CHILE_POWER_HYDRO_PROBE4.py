"""
Chile reservoirs, fourth probe: the JSON API behind DGA's Visualizador
Hidrometrico Nacional (https://vipnet.mop.gob.cl, Angular app; endpoints
read from its main-*.js bundle by CHILE_POWER_HYDRO_PROBE3.py):
  GET  /v1/general/system, /v1/vipnet/parametro/MAP-EST
  POST /v1/vipnet/mediciones/fecha-inicio {tipoEstacion}
  POST /v1/vipnet/estaciones/valor {tipoEstacion, mapStatistic, currentTabIndex, fetchHour, fetchDay, hoursRange}
  POST /v1/vipnet/estacion/valores {codigoEstacion, tipoEstacion, fetchHour, fetchDay, hoursRange}
tipoEstacion 2 = 'Embalse' (reservoir volume, Mm3). Tests how far back the
history goes and how long a window one request can return.
"""
import json
import os

import requests

OUT = os.environ.get("PROBE_OUT", "probe_out")
API = "https://vipnet.mop.gob.cl"
S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0", "Accept": "application/json",
                  "Origin": API, "Referer": API + "/"})
os.makedirs(OUT, exist_ok=True)


def call(tag, method, path, body=None):
    try:
        r = S.request(method, API + path, json=body, timeout=(10, 120))
    except Exception as e:
        print(f"{tag}: ERR {e}", flush=True)
        return None
    print(f"{tag}: {method} {path} {json.dumps(body)} -> {r.status_code} {len(r.content)}B {r.text[:300]!r}", flush=True)
    with open(os.path.join(OUT, f"{tag}.json"), "wb") as f:
        f.write(r.content)
    try:
        return r.json()
    except Exception:
        return None


call("system", "GET", "/v1/general/system")
call("param_mapest", "GET", "/v1/vipnet/parametro/MAP-EST")
call("oldest_2", "POST", "/v1/vipnet/mediciones/fecha-inicio", {"tipoEstacion": 2})
call("stations_2", "POST", "/v1/vipnet/estaciones", {"tipoEstacion": 2})
now = call("map_now", "POST", "/v1/vipnet/estaciones/valor",
           {"tipoEstacion": 2, "mapStatistic": 4, "currentTabIndex": 0, "fetchHour": 6, "fetchDay": "2026-09-30",
            "hoursRange": 3})
call("map_current", "POST", "/v1/vipnet/estaciones/valor", {"tipoEstacion": 2, "currentTabIndex": 1, "hour": 3})
for day in ["2025-06-15", "2023-01-15", "2021-01-15", "2018-01-15"]:
    call(f"map_{day}", "POST", "/v1/vipnet/estaciones/valor",
         {"tipoEstacion": 2, "mapStatistic": 4, "currentTabIndex": 0, "fetchHour": 6, "fetchDay": day, "hoursRange": 3})
data = (now or {}).get("data") or []
print("stations:", len(data), flush=True)
for d in data:
    print("  ", {k: d.get(k) for k in ("codigoEstacion", "nombre", "value", "regionEstacion", "fecha")}, flush=True)
picks = [d for d in data if any(k in str(d.get("nombre", "")).upper() for k in ("COLBUN", "RAPEL", "LAJA", "RALCO"))][:3]
for d in picks or data[:2]:
    code = d["codigoEstacion"]
    for hr in (72, 720, 8760):
        for day in ("2026-09-30", "2021-03-01"):
            j = call(f"series_{code}_{hr}_{day}", "POST", "/v1/vipnet/estacion/valores",
                     {"codigoEstacion": code, "tipoEstacion": 2, "fetchHour": 23, "fetchDay": day, "hoursRange": hr})
            rows = (j or {}).get("data") or []
            if rows:
                print(f"    {code} {hr}h {day}: {len(rows)} rows {rows[0]} .. {rows[-1]}", flush=True)
