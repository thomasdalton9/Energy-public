"""
Round 4 for Argentina gas exports by destination (see rounds 1-3).

Round 3: the ENARGAS daily export report page refreshes its table through
PD_RefrescaListadoImportacionExportacion(tipo) in
secciones/transporte-y-distribucion/funciones/js/funciones.js (max 365 days
per request), and links a monthly report,
/secciones/reportes/dod-reporte-mensual-exportaciones.php. This prints the
whole JS function (AJAX URL + parameters), the xls export function, the
monthly report page, then calls the AJAX endpoint for Jan-2021 to check
history depth.
"""
import re

import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 90)
BASE = "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/"
JS = BASE + "funciones/js/funciones.js"
MONTHLY = "https://www.enargas.gob.ar/secciones/reportes/dod-reporte-mensual-exportaciones.php"


def flat(s):
    return re.sub(r"\s+", " ", s)


def main():
    s = requests.Session()
    s.headers.update(HEADERS)
    js = s.get(JS, timeout=TIMEOUT).text
    for fn in ("function PD_RefrescaListadoImportacionExportacion", "function PD_ExportarListadoImportacionExportacion"):
        i = js.find(fn)
        j = js.find("\nfunction ", i + 10)
        print(f"=== {fn} ===")
        print(flat(js[i:j if j > 0 else i + 5000])[:5000])

    r = s.get(MONTHLY, timeout=TIMEOUT)
    print(f"\n=== monthly report status={r.status_code} bytes={len(r.content)} ct={r.headers.get('content-type')}")
    text = flat(re.sub(r"<[^>]+>", " ", re.sub(r"<script.*?</script>|<style.*?</style>", " ", r.text, flags=re.S)))
    print(text[:4000])
    for m in re.finditer(r"<(form|select|input|option)[^>]*>", r.text, re.I):
        print("  ", flat(m.group(0))[:200])
    for m in re.finditer(r"<script[^>]*>(.*?)</script>", r.text, re.S | re.I):
        if "ajax" in m.group(1).lower() or ".php" in m.group(1):
            print("  inline js:", flat(m.group(1))[:2000])

    # AJAX endpoints: try the url pattern found in the JS
    urls = sorted(set(re.findall(r"url\s*:\s*[\"']([^\"']+)[\"']", js)))
    print("\nall ajax urls in funciones.js:", urls)


if __name__ == "__main__":
    main()
