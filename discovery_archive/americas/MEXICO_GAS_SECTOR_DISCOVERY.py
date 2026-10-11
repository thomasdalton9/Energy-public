"""
One-off probe: where can Mexico natural gas demand by sector be downloaded (monthly)?

Written because the editing sandbox cannot reach any Mexican source (proxy 403 on sie.energia.gob.mx,
datos.gob.mx, pemex.com, ri.pemex.com, gob.mx/cenagas, cre.gob.mx). Run it in GitHub Actions
(discovery_archive/workflows/mexico_gas_sector_discovery.yml); it logs status/content-type of every candidate
and saves pages and data files under mx_raw/ for offline inspection. Then pin the confirmed file URLs in
americas/MEXICO_GAS_SECTOR.py (PINNED) and adjust parse_table() to the real layout.

Candidates:
  - datos.gob.mx CKAN package_search for Pemex / SENER gas-by-sector datasets (all resources listed with format,
    url, last_modified; csv/xls/xlsx saved)
  - SENER SIE (sie.energia.gob.mx): the Balance Nacional de Gas Natural and 'Ventas internas' cuadros
  - Pemex Base de Datos Institucional (bdi.pemex.com) and Indicadores Petroleros
  - CENAGAS Boletin / SISTRANGAS statistics (gob.mx/cenagas)
  - SENER Prospectiva de Gas Natural (PDF, annual demand by sector)
"""
import os
import re
import requests

OUT = "mx_raw"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept-Language": "es-MX,es;q=0.9"}
PAGES = [
    "https://sie.energia.gob.mx/",
    "https://sie.energia.gob.mx/bdiController.do?action=temas&tema=Gas%20natural",
    "https://sie.energia.gob.mx/bdiController.do?action=cuadro&cvecuadro=PMXE1C03&locale=es",
    "https://bdi.pemex.com/",
    "https://www.pemex.com/ri/Publicaciones/Paginas/IndicadoresPetroleros.aspx",
    "https://www.gob.mx/cenagas/acciones-y-programas/boletines-estadisticos",
    "https://www.gob.mx/cenagas/documentos/sistrangas",
    "https://www.gob.mx/sener/documentos/prospectiva-de-gas-natural",
    "https://www.gob.mx/cre/documentos/estadisticas",
]
QUERIES = ["gas natural", "ventas internas gas natural", "balance nacional gas natural", "demanda gas natural sector",
           "consumo gas natural", "sistrangas", "pemex gas natural", "prospectiva gas natural"]


def save(name, content):
    os.makedirs(OUT, exist_ok=True)
    with open(os.path.join(OUT, re.sub(r"[^A-Za-z0-9._-]+", "_", name)[-120:]), "wb") as f:
        f.write(content)


def main():
    for u in PAGES:
        try:
            r = requests.get(u, headers=H, timeout=(20, 90))
            print(f"{r.status_code} {len(r.content):>8} {r.headers.get('content-type', '')[:40]} {u}")
            if r.status_code == 200:
                save(u, r.content)
                for l in sorted(set(re.findall(r'href="([^"]+\.(?:xlsx?|csv|pdf|zip))"', r.text, re.I)))[:80]:
                    print("     link:", l)
        except requests.RequestException as e:
            print(f"ERR {type(e).__name__} {u}")
    seen = set()
    for q in QUERIES:
        try:
            r = requests.get("https://datos.gob.mx/api/3/action/package_search", params={"q": q, "rows": 40},
                             headers=H, timeout=(20, 90))
            res = r.json()["result"]["results"] if r.status_code == 200 else []
            print(f"CKAN '{q}': HTTP {r.status_code}, {len(res)} datasets")
        except (requests.RequestException, ValueError, KeyError) as e:
            print(f"CKAN '{q}': {type(e).__name__}")
            continue
        for p in res:
            for x in p.get("resources", []):
                if x.get("id") in seen:
                    continue
                seen.add(x.get("id"))
                print(f"  [{p.get('organization', {}).get('title', '')}] {p.get('title')} | {x.get('format')} | "
                      f"{x.get('last_modified')} | {x.get('url')}")
                if re.search(r"\.(csv|xlsx?)($|\?)", x.get("url") or "", re.I) and "gas" in p.get("title", "").lower():
                    try:
                        d = requests.get(x["url"], headers=H, timeout=(20, 120))
                        if d.status_code == 200:
                            save(f"{p.get('name')}_{x.get('id')}_{os.path.basename(x['url'])}", d.content)
                    except requests.RequestException:
                        pass


if __name__ == "__main__":
    main()
