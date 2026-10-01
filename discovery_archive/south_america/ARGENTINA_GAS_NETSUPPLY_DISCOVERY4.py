"""
Round 4 for Argentina gas net supply (see rounds 1-3).

Round 3 found:
  - No daily injection table: the 'Inyeccion Nacional por Gasoducto' PDFs (ING_yyyymmdd)
    are Power BI charts of a 15-day window, only for recent days (2021-2025 files 404).
  - 'Partes de Distribucion y Transporte' (POST partes-diarios-listado.php,
    fecha_desde/fecha_hasta as yyyymmdd) lists one row per day back to 2021 with
    three PDF links each (Requerido / Real / Transporte), via DecargarPDF(...).
  - SE Sesco gas balance pivot: producer-declared destinations by company/area
    (Entregado a TGN / TGS / Distribuidoras / Generadores / Industrias / Otros
    Productores, Recibido de Terceros, ...); transfers between producers double-count,
    so it is not used for net supply. ENARGAS GRT (receipts metered by the
    transporters) is.
This prints the listing HTML and the DecargarPDF JS, then downloads the three PDFs for
two days (2021 and 2026) and prints their text, to see whether the daily 'Real' report
has injection by basin / receipt point.
"""
import io
import re

import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 180)
BASE = "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/"
S = requests.Session()
S.headers.update(HEADERS)


def flat(s):
    return re.sub(r"\s+", " ", s)


def pdf_text(content):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        print(f"   pages={len(pdf.pages)}")
        for i, p in enumerate(pdf.pages[:4]):
            print(f"   --- page {i + 1} text:")
            print((p.extract_text() or "")[:6000])
            tables = p.extract_tables()
            print(f"   --- page {i + 1}: {len(tables)} tables; first rows: {[t[:4] for t in tables[:3]]}")


def main():
    S.get(BASE + "dod-partes-dist-trans.php", timeout=TIMEOUT)
    js = S.get(BASE + "funciones/js/funciones.js", timeout=TIMEOUT).text
    i = js.find("function DecargarPDF")
    print("JS:", flat(js[i:js.find("\nfunction ", i + 10)])[:2500])
    for d in ("20210601", "20260901"):
        r = S.post(BASE + "partes-diarios-listado.php", data={"fecha_desde": d, "fecha_hasta": d}, timeout=TIMEOUT)
        print(f"\n===== {d}: listing html ({len(r.text)} bytes):\n{flat(r.text)[:3000]}")
        calls = re.findall(r"DecargarPDF\(([^)]*)\)", r.text)
        print("calls:", calls)
        params_js = re.search(r"function DecargarPDF\s*\(([^)]*)\)", js)
        names = [a.strip() for a in params_js.group(1).split(",")] if params_js else []
        body = js[i:js.find("\nfunction ", i + 10)]
        exprs = re.findall(r"(?:window\.open|location\.href\s*=|url\s*[:=])\s*\(?\s*([^;\n]+)", body)
        print("param names:", names, "url expressions:", exprs)
        for c in calls[:3]:
            args = [a.strip().strip("'\"") for a in c.split(",")]
            env = dict(zip(names, args))
            for e in exprs:
                e = re.split(r",(?=(?:[^\"']*[\"'][^\"']*[\"'])*[^\"']*$)", e)[0].rstrip(")")
                parts = re.findall(r"\"([^\"]*)\"|'([^']*)'|([A-Za-z_]\w*)", e)
                url = "".join(a or b or env.get(v, "") for a, b, v in parts)
                u = url if url.startswith("http") else BASE + url.lstrip("/")
                try:
                    f = S.get(u, timeout=TIMEOUT)
                except requests.RequestException as ex:
                    print("   ERROR", u, ex)
                    continue
                ok = f.content[:4] == b"%PDF"
                print(f"   {u}: status={f.status_code} bytes={len(f.content)} pdf={ok} ct={f.headers.get('content-type')}")
                if ok:
                    pdf_text(f.content)
                    break


if __name__ == "__main__":
    main()
