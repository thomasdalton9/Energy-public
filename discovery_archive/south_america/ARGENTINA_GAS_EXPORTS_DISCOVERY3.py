"""
Round 3 for Argentina gas exports by destination.

Round 2 found ENARGAS's monthly Exportaciones.xlsx covers only exports
"dentro del sistema de transporte" and leaves out "Brasil por Bolivia";
the ENARGAS daily export reports (Partes diarios de exportacion) have
two tables, both in thousand m3/day, with an export-point column each:
  dentro del sistema: Chile GasAndes, Chile NorAndino, Chile Methanex YPF,
    Uruguay PetroUruguay, Brasil TGM, Uruguay Cruz del Sur,
    Chile Methanex EGS, Brasil por Bolivia, Total
  fuera del sistema (producers' own pipelines): Chile Methanex PAE,
    Chile Methanex SIP, Chile Atacama, Chile Methanex PTB, Chile Pacifico,
    Brasil por Bolivia, Total
The page loads 01/01/<year> to today by default via a JS function
PD_RefrescaListadoImportacionExportacion(tipo). This finds the AJAX
endpoint / xls download link and checks how far back the reports go.
"""
import re
from urllib.parse import urljoin

import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 90)
PAGE = "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/dod-partes-exp-imp-consulta.php"


def main():
    s = requests.Session()
    s.headers.update(HEADERS)
    r = s.get(PAGE, params={"tipo": "exp_dentro"}, timeout=TIMEOUT)
    html = r.text
    print(f"page status={r.status_code} bytes={len(html)}")
    for m in re.finditer(r"<script[^>]*src=[\"']([^\"']+)[\"']", html, re.I):
        print("script src:", m.group(1))
    for m in re.finditer(r"<a[^>]+href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", html, re.I | re.S):
        t = re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()
        if any(k in (m.group(1) + t).lower() for k in ("xls", "pdf", "descargar", "javascript")):
            print("link:", repr(t[:40]), m.group(1)[:200])
    for m in re.finditer(r"(onclick|href)=[\"']([^\"']*(?:xls|Descarga|Exporta|PD_)[^\"']*)[\"']", html, re.I):
        print("handler:", m.group(2)[:300])
    i = html.find("<table")
    print("first table html:", re.sub(r"\s+", " ", html[i:i + 2500]))
    inline = re.findall(r"<script[^>]*>(.*?)</script>", html, re.S | re.I)
    for js in inline:
        if "PD_" in js or "ajax" in js.lower() or ".php" in js:
            print("inline js:", re.sub(r"\s+", " ", js)[:2500])
    # look in external js for the function
    for m in re.finditer(r"<script[^>]*src=[\"']([^\"']+)[\"']", html, re.I):
        src = urljoin(PAGE, m.group(1))
        if "jquery" in src.lower() or "bootstrap" in src.lower():
            continue
        try:
            js = s.get(src, timeout=TIMEOUT).text
        except requests.RequestException as e:
            print("js fetch error", src, e)
            continue
        for fn in ("PD_RefrescaListadoImportacionExportacion", "PD_Descarga", "ImportacionExportacion"):
            for mm in re.finditer(fn, js):
                print(f"--- {src} @ {mm.start()} ({fn}):")
                print(re.sub(r"\s+", " ", js[max(0, mm.start() - 200):mm.start() + 1800]))
                break

    # guess: same page with date params
    for params in ({"tipo": "exp_dentro", "fechadesde": "01/01/2021", "fechahasta": "31/01/2021"},
                   {"tipo": "exp_dentro", "desde": "01/01/2021", "hasta": "31/01/2021"},
                   {"tipo": "exp_fuera", "fechadesde": "01/01/2021", "fechahasta": "31/01/2021"}):
        r = s.get(PAGE, params=params, timeout=TIMEOUT)
        dates = re.findall(r"\d{2}/\d{2}/\d{4}", r.text)
        print(f"GET {params}: status={r.status_code} first dates {dates[:4]} last {dates[-2:]}")


if __name__ == "__main__":
    main()
