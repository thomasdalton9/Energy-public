"""
Round 5 (Oct-2026): fetch the candidate gas / annual files so they can be
inspected offline (saved under probe_files/).
  do5   SIE "Consumo Combustible 2015-2026" (datos.gob.do, monthly fuel use of the
        generators), CNE data repository "Generacion y Consumo por Central"
        (datacne.gob.do Dropbox listing API: POST /api/dropbox/list {path})
  jm5   MSET "Jamaica Energy Statistics" 2025 / 2024 / 2023 PDFs (annual tables:
        JPS electricity generation by source, natural gas consumption)
Usage: python3 CARIBBEAN_POWER_GAS_PROBE5.py do5|jm5
"""

print("STARTING", flush=True)

import json
import os
import re
import sys

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "es-ES,es;q=0.9,en;q=0.8"}
S = requests.Session()
S.headers.update(H)
OUT = "probe_files"
os.makedirs(OUT, exist_ok=True)


def get(url, save=None, method="GET", **kw):
    try:
        r = S.request(method, url, timeout=120, **kw)
        print(f"[{r.status_code}] {method} {r.url[:200]} {r.headers.get('content-type')} {len(r.content):,} B", flush=True)
        if save and r.ok:
            with open(os.path.join(OUT, save), "wb") as f:
                f.write(r.content)
        return r
    except requests.RequestException as e:
        print(f"ERR {url}: {type(e).__name__}: {str(e)[:200]}", flush=True)
        return None


def cne_list(path, depth=0, maxdepth=3):
    r = get("https://datacne.gob.do/api/dropbox/list", method="POST", json={"path": path})
    if r is None or not r.ok:
        return
    try:
        files = r.json().get("files", [])
    except ValueError:
        print(r.text[:500])
        return
    for f in files:
        print("  " * depth + f"- {f.get('type')} {f.get('path_display')} {f.get('size')} {f.get('modifiedAt')}", flush=True)
        if f.get("type") == "folder" and depth < maxdepth:
            cne_list(f.get("path_lower"), depth + 1, maxdepth)
    return files


def do5():
    r = get("https://datos.gob.do/api/3/action/package_show", params={"id": "energia-y-potencia-facturadas-ede"})
    if r is not None and r.ok:
        p = r.json()["result"]
        print(json.dumps({k: p.get(k) for k in ("title", "notes", "metadata_modified")}, ensure_ascii=False)[:1500])
        for res in p.get("resources", []):
            print("   RES", res.get("format"), res.get("url"))
    get("https://sie.gob.do/wp-content/uploads/2025/07/Copia-de-Consumo-Combustible_Mensual_2026-xls-1.xlsx",
        save="sie_consumo_combustible.xlsx")
    get("https://sie.gob.do/wp-content/uploads/2025/07/Copia-de-Consumo-Combustible_Mensual_2026-xls.csv.csv",
        save="sie_consumo_combustible.csv")
    cne_list("/generación y consumo por central")
    # try to fetch one file through the site's own preview / proxy endpoints
    r = get("https://datacne.gob.do/api/dropbox/list", method="POST", json={"path": "/generación y consumo por central"})
    first = None
    try:
        stack = [f for f in r.json()["files"]]
        while stack and first is None:
            f = stack.pop(0)
            if f["type"] == "folder":
                rr = get("https://datacne.gob.do/api/dropbox/list", method="POST", json={"path": f["path_lower"]})
                stack = rr.json().get("files", []) + stack
            else:
                first = f
    except Exception as e:
        print("walk error", e)
    if first:
        print("FIRST FILE", first)
        for method, url, kw in [
            ("POST", "https://datacne.gob.do/api/dropbox/preview", {"json": {"path": first["path_lower"]}}),
            ("GET", "https://datacne.gob.do/api/dropbox/preview", {"params": {"path": first["path_lower"]}}),
            ("GET", "https://datacne.gob.do/api/dropbox/pdf-proxy", {"params": {"path": first["path_lower"]}}),
        ]:
            rr = get(url, method=method, **kw)
            if rr is not None:
                ct = rr.headers.get("content-type") or ""
                print("   ", ct, rr.text[:600] if "json" in ct or "text" in ct else rr.content[:8])
                if rr.ok and "json" in ct:
                    for m in re.findall(r"https://[^\"']+", rr.text)[:3]:
                        get(m, save="cne_first_" + re.sub(r"\W", "_", first["name"])[-60:])


def jm5():
    for name in ["JAMAICA-ENERGY-STATISTICS-2025.pdf", "JAMAICA-ENERGY-STATISTICS-2024.pdf",
                 "JAMAICA-ENERGY-STATISTICS-2023-Revised-August-2024.pdf"]:
        get(f"https://www.mset.gov.jm/wp-content/uploads/2021/07/{name}", save=name)


if __name__ == "__main__":
    {"do5": do5, "jm5": jm5}[sys.argv[1]]()
    print("DONE", flush=True)
