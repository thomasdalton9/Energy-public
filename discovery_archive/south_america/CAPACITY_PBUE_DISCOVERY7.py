"""
Discovery round 7 (after CAPACITY_PBUE_DISCOVERY6.py).

Round 6 found:
  Peru: COES folders list sub-folders per year via openBlob('<path>', 'D'|'F', ...); 'Potencia Efectiva' has
        2009-2026 folders, 'Boletines' 2011-2026. Chapter 2 'BD' sheet also has an INGRESO/RETIRO table
        (unit, MW, date) below the unit list.
  Ecuador: BNEE_<mes>_<anio>.xls: 'Potencia Nominal en Generacion (2)' MW, Total and S.N.I.: Hidraulica, Eolica,
        Fotovoltaica, Biomasa, Biogas, MCI, Turbogas, Turbovapor, Importacion. Only the latest month is linked;
        WP media has one PNG per month (abril 2024 .. junio 2026), no download-monitor API, directory listing 403.

This round: COES year folders (file names); ARCONEL older BNEE file URLs (guessed from the PNG upload months)
and the hrefs of ARCONEL's statistics publications page.
"""
import html
import itertools
import re
import sys

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from CAPACITY_PBUE_DISCOVERY import get  # noqa: E402

COES = "https://www.coes.org.pe/Portal/"


def coes_items(path):
    r = get(COES + "browser/vistadatos", method="POST",
            data={"baseDirectory": path, "url": path, "indicador": "", "initialLink": "", "orderFolder": ""})
    if r is None or r.status_code != 200:
        return []
    items = []
    for m in re.finditer(r"openBlob\('([^']+)',\s*'(\w)'", r.text):
        it = (html.unescape(m.group(1)), m.group(2))
        if it not in items:
            items.append(it)
    sizes = re.findall(r"<td[^>]*text-align:right[^>]*>([^<]*)</td>", r.text)
    print(f"     {path}: {len(items)} items")
    for it in items[:40]:
        print(f"       {it[1]} {it[0]}")
    return items


def peru():
    print("\n================ PERU")
    for p in ["Publicaciones/Boletines/2026/", "Publicaciones/Boletines/2021/",
              "Operación/Estudios/Potencia Efectiva/2026/", "Publicaciones/Estadisticas Anuales/2021/Excel/",
              "Publicaciones/Estadisticas Anuales/2022/Excel/", "Publicaciones/Estadisticas Anuales/2023/Excel/",
              "Publicaciones/Estadisticas Anuales/2024/Excel/"]:
        items = coes_items(p)
        for sub, kind in items[:2]:
            if kind == "D":
                coes_items(sub)


MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
         "noviembre", "diciembre"]


def ecuador():
    print("\n================ ECUADOR")
    r = get("https://arconel.gob.ec/publicaciones-estadistica-del-sector-electrico-2/")
    if r is not None:
        hrefs = sorted(set(re.findall(r'href=["\']([^"\']+)["\']', r.text)))
        for h in hrefs:
            if not re.search(r"\.(css|js|png|jpg|svg|woff2?)(\?|$)|/feed/|wp-json|xmlrpc|#", h):
                print("      href:", h)
    # (month, year, upload year/month of the PNG) from round 5's media listing
    tests = [("mayo", 2026, "2026/07"), ("enero", 2026, "2026/04"), ("diciembre", 2025, "2026/03"),
             ("junio", 2025, "2025/10"), ("diciembre", 2024, "2025/04"), ("abril", 2024, "2024/07")]
    found = 0
    for mes, year, up in tests:
        y, m = map(int, up.split("/"))
        months = [f"{y}/{m:02d}", f"{y + (m == 12)}/{m % 12 + 1:02d}", f"{y - (m == 1)}/{(m - 2) % 12 + 1:02d}"]
        hit = None
        for folder, suffix, ext in itertools.product(months, ["", "_revACH", "-1", "_rev", "_v2", "_final",
                                                              "_revisado", "_rev1"], [".xls", ".xlsx"]):
            for base in ("https://arconel.gob.ec/wp-content/uploads/downloads/", "https://arconel.gob.ec/wp-content/uploads/"):
                for nm in (f"BNEE_{mes}_{year}", f"BNEE_{mes.capitalize()}_{year}"):
                    u = f"{base}{folder}/{nm}{suffix}{ext}"
                    try:
                        import requests
                        rr = requests.head(u, timeout=20, allow_redirects=True, headers={"User-Agent": "Mozilla/5.0"})
                    except Exception:  # noqa: BLE001
                        continue
                    if rr.status_code == 200:
                        hit = u
                        break
                if hit:
                    break
            if hit:
                break
        print(f"     {mes} {year}: {hit}")
        found += bool(hit)
    print(f"     found {found}/{len(tests)}")


if __name__ == "__main__":
    for w in sys.argv[1:] or ["peru", "ecuador"]:
        try:
            globals()[w]()
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            print(f"!! {w} failed: {type(e).__name__}: {e}")
        sys.stdout.flush()
