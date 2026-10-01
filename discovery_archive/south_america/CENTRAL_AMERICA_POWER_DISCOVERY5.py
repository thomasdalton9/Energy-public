"""
Round 5 of CENTRAL_AMERICA_POWER_DISCOVERY.py: check the Nicaragua unit -> technology
mapping against CNDC's own by-type figures.

Round 4 found: CNDC post-dispatch gives MWh per unit (GGD) only; agent names come from
/SectorElectrico (e.g. GEOSA, PENSA = Polaris geothermal, PMN-PHC = ALBA Generacion,
NFE = Nicaragua Development Partners, which has no units in post-dispatch 2021-2026).
CNDC's by-type series (graficos/consultarGeneracionPorTipo, Inicio/ConsultarTipoGeneracion)
only cover the current day, and mapageneracion/ConsultarGeneracionMapa gives real-time
MW per unit. Taking both at the same moment, a few times, lets the unit groups be
matched to CNDC's TERMICA / HIDROELECTRICA / GEOTERMICA / EOLICA / BIOMASA / SOLAR.
"""
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import CENTRAL_AMERICA_POWER_DISCOVERY as d1  # noqa: E402
from CENTRAL_AMERICA_POWER_DISCOVERY2 import Tee  # noqa: E402


def nicaragua():
    b = "https://www.cndc.org.ni/"
    h = {"X-Requested-With": "XMLHttpRequest"}
    snaps = []
    for i in range(4):
        snap = {}
        for ep in ["Inicio/ConsultarTipoGeneracion", "mapageneracion/ConsultarGeneracionMapa"]:
            r = d1.get(b + ep, headers=h, save=False)
            try:
                snap[ep] = r.json()
            except Exception as e:  # noqa: BLE001
                print("  parse failed", e, r.text[:300] if r is not None else "")
        snaps.append(snap)
        if i < 3:
            time.sleep(240)
    with open(os.path.join(d1.RAW, "ni_snapshots.json"), "w") as f:
        json.dump(snaps, f, ensure_ascii=False)
    d1.get(b + "graficos/consultarGeneracionPorTipo", headers=h)
    d1.get(b + "mapageneracion", save=True)


if __name__ == "__main__":
    os.makedirs(d1.RAW, exist_ok=True)
    log = open(os.path.join(d1.RAW, "log_round5.txt"), "w")
    sys.stdout = Tee(sys.__stdout__, log)
    try:
        nicaragua()
    except Exception as e:  # noqa: BLE001
        print(f"FAILED: {type(e).__name__}: {e}", flush=True)
    sys.stdout.flush()
    sys.stdout = sys.__stdout__
    log.close()
