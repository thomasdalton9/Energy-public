"""
Round 2 of CENTRAL_AMERICA_POWER_DISCOVERY.py.

Round 1 found:
  Costa Rica: CenceWeb data/sen/json/EnergiaHorariaFuentePlanta and
    EnergiaHorariaFuente answer JSON, but ignore inicio/fin (always the
    latest day). Pages paginas/GeneracionReal.html, ServiciosWeb.html and
    GuiaServiciosWeb.html should show the real parameter names.
  Nicaragua: cndc.org.ni was rebuilt (Laravel); the old .php report URLs are
    404. New pages /graficos (generacion por recursos) and
    /MEMNDiarios/posdespachoEnergia load data from JS.
  Panama: cnd.com.pa (Joomla) loads reports and statistics through
    sitiopublico.cnd.com.pa/index.php/api?task=... from module JS files
    (mod_logic_estadisticas, mod_logic_informescnd, ...). sitr.cnd.com.pa
    returned Cloudflare 522.

This round fetches those pages plus every custom script they load, so the
AJAX endpoints and parameter names can be read offline.
"""
import os
import re
import sys
import urllib.parse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import CENTRAL_AMERICA_POWER_DISCOVERY as d1  # noqa: E402

LIBS = re.compile(r"jquery|bootstrap|highcharts|highstock|leaflet|fontawesome|moment|popper|owl\.|three|aos\.|"
                  r"pdf|sweetalert|intro|bowser|cloudflare|chart\.js|chartjs|caption|html5|tempus|camera|easing|wrld|"
                  r"datatables|select2|toastr|amcharts", re.I)


class Tee:
    def __init__(self, *s):
        self.s = s

    def write(self, x):
        for s in self.s:
            s.write(x)

    def flush(self):
        for s in self.s:
            s.flush()


def page_and_scripts(url, **kw):
    r = d1.get(url, "PAGE", **kw)
    if r is None or "html" not in r.headers.get("content-type", ""):
        return r
    for u in d1.links(r):
        print("    link " + u)
    host = urllib.parse.urlparse(url).netloc.split(".", 1)[-1]
    for src in re.findall(r"""<script[^>]*src=["']([^"']+)""", r.text, re.I):
        full = urllib.parse.urljoin(r.url, src)
        if host not in full or LIBS.search(full) or full in d1.SEEN:
            continue
        d1.SEEN.add(full)
        d1.get(full, "SCRIPT")
    for inline in re.findall(r"<script[^>]*>(.*?)</script>", r.text, re.S | re.I):
        for m in set(re.findall(r"""["'`]([^"'`\s]*(?:/api|task=|json|\.php|Consultar|consultar|obtener|Obtener)[^"'`\s]*)["'`]""", inline)):
            print("    inline-url " + m)
    return r


def costa_rica():
    b = "https://apps.grupoice.com/CenceWeb/"
    for p in ["paginas/GeneracionReal.html", "paginas/ServiciosWeb.html", "paginas/GuiaServiciosWeb.html",
              "paginas/PosdespachosDiarios.html", "paginas/CurvaDemanda.html"]:
        page_and_scripts(b + p)
    # monthly reports page (list of files)
    page_and_scripts(b + "CenceDescargaArchivos.jsf?init=true&categoria=3&codigoTipoArchivo=3007&fecha_inic=ante")


def nicaragua():
    b = "https://www.cndc.org.ni/"
    for p in ["graficos", "MEMNDiarios/posdespachoEnergia", "MEMNDiarios/posdespachoPotencia"]:
        page_and_scripts(b + p)
    for p in ["js/customScript/graficos/generacionPorTipo.js",
              "js/customScript/MEMNDiarios/PrePos_Energia/posdespacho/postdespacho.js",
              "js/customScript/MEMNDiarios/PrePos_Energia/posdespacho/resumen.js",
              "js/customScript/MEMNDiarios/PrePos_Energia/posdespacho/curvaDespacho.js",
              "js/customScript/MEMNDiarios/PrePos_Energia/posdespacho/identificacionGGD.js",
              "js/customScript/graficos/MEMNDiarios/curvademandaGeneracion.js",
              "Inicio/ConsultarTipoGeneracion"]:
        if b + p not in d1.SEEN:
            d1.SEEN.add(b + p)
            d1.get(b + p)


def panama():
    s = "https://sitiopublico.cnd.com.pa/modules/"
    for m in ["mod_logic_estadisticas", "mod_logic_informescnd", "mod_logic_entregachart",
              "mod_logic_inforenovables", "mod_logic_boundariechart", "mod_logic_costomarginalchart"]:
        d1.get(s + m + "/assets/js/app.js")
    for p in ["https://www.cnd.com.pa/index.php/estadisticas",
              "https://www.cnd.com.pa/index.php/informes/categoria/informes-de-operaciones",
              "https://www.cnd.com.pa/index.php/informes/categoria/informe-historicos-cnd",
              "https://sitr.cnd.com.pa/m/"]:
        page_and_scripts(p)


if __name__ == "__main__":
    os.makedirs(d1.RAW, exist_ok=True)
    log = open(os.path.join(d1.RAW, "log_round2.txt"), "w")
    sys.stdout = Tee(sys.__stdout__, log)
    for w in sys.argv[1:] or ["cr", "ni", "pa"]:
        print(f"\n\n######## {w} ########", flush=True)
        try:
            {"cr": costa_rica, "ni": nicaragua, "pa": panama}[w]()
        except Exception as e:  # noqa: BLE001
            print(f"FAILED {w}: {type(e).__name__}: {e}", flush=True)
    log.close()
