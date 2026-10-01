"""
Round 3 of CENTRAL_AMERICA_POWER_DISCOVERY.py: try the real endpoints found in round 2.

Round 2 found:
  Costa Rica: paginas/GeneracionReal.html calls
    data/sen/json/EnergiaHorariaFuentePlanta?anno=YYYY&mes=M&dia=D (also /csv/, /xml/);
    ServiciosWeb.html lists services via data/servicios/json/WebServices/.
  Nicaragua: MEMNDiarios pages call /MEMNDiarios/consultarPosdespachoEnergia,
    ConsultarPosResumenEnergia, ConsultarCurvaDemandaGeneracion,
    consultarPosIdentificacionGGDEnergia with fecha=DD/MM/YYYY (hourly by unit/GGD);
    /graficos/consultarGeneracionPorTipo (today, by type).
  Panama: cnd.com.pa report lists come from sitioprivado.cnd.com.pa
    Informe/GetListOperativosComerciales (categoria 6 = operaciones, 14 = mercado)
    and TipoInforme/GetList, with the public key embedded in the site JS;
    files download from /Informe/Download/{id}?key=...; Estadistica/GetList.
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import CENTRAL_AMERICA_POWER_DISCOVERY as d1  # noqa: E402
from CENTRAL_AMERICA_POWER_DISCOVERY2 import Tee  # noqa: E402

PA_KEY = "VXd9e23Z9JRA5aIUR21R-P8gocoGOMqdvSo79FduN"  # public, from sitiopublico mod_logic_informescnd/app.js
PA = "https://sitioprivado.cnd.com.pa/"


def costa_rica():
    b = "https://apps.grupoice.com/CenceWeb/data/"
    d1.get(b + "servicios/json/WebServices/")
    for q in ["EnergiaHorariaFuente", "EnergiaHorariaFuentePlanta"]:
        d1.get(b + "servicios/json/ParametrosServicioWeb", params={"nombreConsulta": q})
    for p in [{"anno": 2021, "mes": 1, "dia": 1}, {"anno": 2021, "mes": 1}, {"anno": 2023, "mes": 6, "dia": 15}]:
        d1.get(b + "sen/json/EnergiaHorariaFuente", params=p)
    d1.get(b + "sen/json/EnergiaHorariaFuentePlanta", params={"anno": 2021, "mes": 1, "dia": 1})
    d1.get(b + "sen/csv/EnergiaHorariaFuentePlanta", params={"anno": 2026, "mes": 9, "dia": 1})


def nicaragua():
    b = "https://www.cndc.org.ni/"
    d1.SESSION.get(b + "MEMNDiarios/posdespachoEnergia", timeout=d1.TIMEOUT)  # cookies
    h = {"X-Requested-With": "XMLHttpRequest", "Referer": b + "MEMNDiarios/posdespachoEnergia"}
    for f in ["05/01/2021", "15/06/2023", "28/09/2026"]:
        for ep in ["consultarPosdespachoEnergia", "ConsultarPosResumenEnergia", "ConsultarCurvaDemandaGeneracion",
                   "consultarPosIdentificacionGGDEnergia"]:
            d1.get(b + "MEMNDiarios/" + ep, params={"fecha": f}, headers=h)
    d1.get(b + "graficos/consultarGeneracionPorTipo", headers=h)
    d1.get(b + "graficos/consultarGeneracionPorTipo", params={"fecha": "05/01/2025"}, headers=h)
    d1.get(b + "graficos/consultarGGDS", headers=h)


def panama():
    r = d1.get(PA + "TipoInforme/GetList", params={"page": 0, "nombre": "", "aplica": 1, "publico": 1, "key": PA_KEY})
    try:
        for t in r.json():
            print("   TIPO", t.get("id"), t.get("name"), "cat", (t.get("categoria") or {}).get("id"),
                  (t.get("categoria") or {}).get("name"), "y/m/w/d", t.get("year"), t.get("month"),
                  t.get("weekend"), t.get("day"))
    except Exception as e:  # noqa: BLE001
        print("   tipos parse failed", e)
    d1.get(PA + "Estadistica/GetList", params={"page": 1, "nombre": "", "key": PA_KEY})
    picked = {}
    for cat, anio, mes in [(6, 2026, 9), (14, 2026, 9), (6, 2021, 1), (17, 0, 0), (8, 0, 0)]:
        r = d1.get(PA + "Informe/GetListOperativosComerciales",
                   params={"page": 0, "publico": 1, "key": PA_KEY, "categoria": cat, "tipo": 0,
                           "anio": anio, "mes": mes, "semana": 0, "dia": 0})
        try:
            rows = r.json()
        except Exception as e:  # noqa: BLE001
            print("   list parse failed", e)
            continue
        print(f"   cat {cat} {anio}-{mes}: {len(rows)} informes")
        for x in rows[:400]:
            t = (x.get("tipo") or {}).get("name")
            path = (x.get("adjunto") or {}).get("path", "")
            print("     ", x.get("id"), t, x.get("fechaPublicaFull"), path.split("\\")[-1])
            if path and (cat, t) not in picked:
                picked[(cat, t)] = x.get("id")
    print("PICKED", json.dumps({f"{k[0]}|{k[1]}": v for k, v in picked.items()}, ensure_ascii=False))
    for (cat, t), i in picked.items():
        if cat in (6, 14, 17):
            d1.get(PA + f"Informe/Download/{i}", params={"key": PA_KEY})


if __name__ == "__main__":
    os.makedirs(d1.RAW, exist_ok=True)
    log = open(os.path.join(d1.RAW, "log_round3.txt"), "w")
    sys.stdout = Tee(sys.__stdout__, log)
    for w in sys.argv[1:] or ["cr", "ni", "pa"]:
        print(f"\n\n######## {w} ########", flush=True)
        try:
            {"cr": costa_rica, "ni": nicaragua, "pa": panama}[w]()
        except Exception as e:  # noqa: BLE001
            print(f"FAILED {w}: {type(e).__name__}: {e}", flush=True)
    sys.stdout.flush()
    sys.stdout = sys.__stdout__
    log.close()
