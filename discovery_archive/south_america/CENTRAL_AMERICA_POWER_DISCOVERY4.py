"""
Round 4 of CENTRAL_AMERICA_POWER_DISCOVERY.py.

Round 3 found (all three are automatable with plain requests):
  Costa Rica: data/sen/json/EnergiaHorariaFuente?anno=&mes=&dia= -> one day of hourly
    MWh by source (Hidro, Eólico, Geotérmico, Solar, Bagazo, Búnker, Diesel,
    Intercambio); history back to 2021-01-01 at least. ~13 KB per day.
  Nicaragua: MEMNDiarios/consultarPosdespachoEnergia?fecha=DD/MM/YYYY -> hourly
    post-dispatch MWh per generating unit (GGD) plus posdespachoTotal per unit,
    back to 2021. No fuel per unit in that response -> need CNDC's own plant
    metadata (mapa de generacion) for the unit -> technology mapping.
  Panama: sitioprivado.cnd.com.pa Informe/GetListOperativosComerciales,
    tipo 110 "Reporte Diario" (REPORTE DIARIO YYYYMMDD.xls) has "ENTREGADO AL
    SISTEMA" MWh by type (Hidro, Bunker, Diesel, Autogeneradores Hidros/Térmicos/
    Solares, Eólicos, Carbón, Solares, Biogas, Gas Natural, Intercambio), daily
    from 2021-01-01.

This round: CNDC plant metadata + every unit code that appears in 2021-2026;
two older Panama daily reports to check the layout is stable.
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import CENTRAL_AMERICA_POWER_DISCOVERY as d1  # noqa: E402
from CENTRAL_AMERICA_POWER_DISCOVERY2 import Tee, page_and_scripts  # noqa: E402
from CENTRAL_AMERICA_POWER_DISCOVERY3 import PA, PA_KEY  # noqa: E402


def nicaragua():
    b = "https://www.cndc.org.ni/"
    h = {"X-Requested-With": "XMLHttpRequest"}
    for p in ["js/customScript/MapaGeneracion/mapaGeneracion.js", "js/customScript/MapaGeneracion/sidebarControl.js",
              "js/customScript/Inicio/informacionTecnica.js"]:
        d1.get(b + p)
    for p in ["MapaGeneracion/mapaGeneracion", "MapaGeneracion/sidebarControl"]:
        d1.get(b + p, headers=h)
    page_and_scripts(b + "SectorElectrico")
    units = {}
    for y in range(2021, 2027):
        for m in range(1, 13):
            if (y, m) > (2026, 9):
                break
            r = d1.get(b + "MEMNDiarios/consultarPosdespachoEnergia", params={"fecha": f"15/{m:02d}/{y}"},
                       headers=h, save=False)
            try:
                for x in r.json()["posdespachoTotal"]:
                    units.setdefault(f'{x["agente"]}|{x["ggd"]}', {})[f"{y}-{m:02d}"] = x["total"]
            except Exception as e:  # noqa: BLE001
                print("   parse failed", e)
    with open(os.path.join(d1.RAW, "ni_units.json"), "w") as f:
        json.dump(units, f, ensure_ascii=False, indent=0)
    for k in sorted(units):
        print("UNIT", k, len(units[k]), "months", min(units[k]), "->", max(units[k]))


def panama():
    for anio, mes in [(2021, 1), (2023, 6), (2025, 3)]:
        r = d1.get(PA + "Informe/GetListOperativosComerciales",
                   params={"page": 0, "publico": 1, "key": PA_KEY, "categoria": 6, "tipo": 110,
                           "anio": anio, "mes": mes, "semana": 0, "dia": 0}, save=False)
        rows = r.json()
        print(f"   {anio}-{mes}: {len(rows)} reportes diarios")
        for x in rows:
            print("     ", x.get("id"), x.get("fechaPublicaFull"), (x.get("adjunto") or {}).get("path", "").split("\\")[-1])
        if rows:
            d1.get(PA + f"Informe/Download/{rows[-1]['id']}", params={"key": PA_KEY})


if __name__ == "__main__":
    os.makedirs(d1.RAW, exist_ok=True)
    log = open(os.path.join(d1.RAW, "log_round4.txt"), "w")
    sys.stdout = Tee(sys.__stdout__, log)
    for w in sys.argv[1:] or ["ni", "pa"]:
        print(f"\n\n######## {w} ########", flush=True)
        try:
            {"ni": nicaragua, "pa": panama}[w]()
        except Exception as e:  # noqa: BLE001
            print(f"FAILED {w}: {type(e).__name__}: {e}", flush=True)
    sys.stdout.flush()
    sys.stdout = sys.__stdout__
    log.close()
