"""
Round 3 of the capacity probes.

argentina: CAMMESA's nemo API has no public installed-capacity publication
  (round 1: INFORME_MENSUAL, POTENCIA_INSTALADA ... 'no es publica'), but
  cammesaweb has WordPress Download Manager pages /download/potencia-instalada/
  and /download/informe-mensual_YYYY-MM/, and the Secretaria de Energia CKAN
  (datos.energia.gob.ar) has a 'publicaciones-cammesa' dataset with CSVs.
  This lists that dataset's resources, the WPDM packages (wp-json) and the
  download links on those pages, and opens what they point to.
colombia: 48 small solar plants (~19.9 MW each, ~580 MW) drop out of XM's
  CapEfecNeta between end-Feb and end-Mar 2025 and never return. Checks
  whether they are still in ListadoRecursos (state) and still generating.

    python3 POWER_CAPACITY_PROBE3.py argentina|colombia
"""
import datetime as dt
import io
import re
import sys
import zipfile

import requests

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"}


def get(url, **kw):
    for attempt in range(3):
        try:
            r = requests.get(url, headers=kw.get("headers", UA), timeout=kw.get("timeout", (20, 180)),
                             params=kw.get("params"), allow_redirects=True)
            print(f"GET {r.url[:220]} -> {r.status_code} {r.headers.get('content-type')} {len(r.content):,} B "
                  f"{r.headers.get('content-disposition', '')}", flush=True)
            return r
        except requests.RequestException as e:
            print(f"GET {url} attempt {attempt + 1} FAILED {type(e).__name__}", flush=True)
    return None


def hdr(s):
    print(f"\n{'=' * 78}\n{s}\n{'=' * 78}", flush=True)


def show_file(content, label):
    import pandas as pd
    head = content[:8]
    if head[:2] == b"PK":
        try:
            zf = zipfile.ZipFile(io.BytesIO(content))
            names = zf.namelist()
            if any(n.startswith("xl/") for n in names):
                xl = pd.ExcelFile(io.BytesIO(content))
                print(f"  {label}: xlsx sheets {xl.sheet_names}")
                for s in xl.sheet_names[:12]:
                    d = xl.parse(s, header=None, nrows=40)
                    print(f"  --- sheet {s} {d.shape}")
                    print(d.dropna(how="all").head(30).to_string(max_colwidth=28)[:5000])
            else:
                print(f"  {label}: zip {names[:30]}")
                for n in names[:3]:
                    show_file(zf.read(n), n)
        except Exception as e:  # noqa: BLE001
            print("  open failed", e)
    elif head[:4] == b"%PDF":
        print(f"  {label}: PDF")
    elif head[:4] == b"\xd0\xcf\x11\xe0":
        try:
            xl = pd.ExcelFile(io.BytesIO(content))
            print(f"  {label}: xls sheets {xl.sheet_names}")
            for s in xl.sheet_names[:8]:
                print(xl.parse(s, header=None, nrows=30).dropna(how="all").to_string(max_colwidth=28)[:4000])
        except Exception as e:  # noqa: BLE001
            print("  xls open failed", e)
    else:
        txt = content[:3000].decode("utf-8", errors="replace")
        print(f"  {label}: text\n{txt}")


def argentina():
    hdr("datos.energia.gob.ar publicaciones-cammesa resources")
    r = get("http://datos.energia.gob.ar/api/3/action/package_show", params={"id": "publicaciones-cammesa"})
    res = r.json()["result"]["resources"] if r is not None and r.ok else []
    for x in res:
        print(f"  {x.get('name')!r} fmt={x.get('format')} modified={x.get('last_modified')} url={x.get('url')}")
    for x in res:
        if re.search(r"potencia|oferta|generador|instalad|parque|central", (x.get("name") or "") + (x.get("url") or ""), re.I) \
                and x.get("format", "").upper() in ("CSV", "XLSX", "XLS", "ZIP"):
            rr = get(x["url"], timeout=(20, 300))
            if rr is not None and rr.ok:
                show_file(rr.content, x.get("name"))
    for q in ["potencia", "oferta", "parque generador", "centrales"]:
        r = get("http://datos.energia.gob.ar/api/3/action/package_search", params={"q": q, "rows": 10})
        if r is not None and r.ok:
            for p in r.json()["result"]["results"]:
                print(f"  [{q}] PKG {p.get('name')}: {p.get('title')}")
                for x in p.get("resources", []):
                    if re.search(r"potencia|instalad|oferta|parque", x.get("name") or "", re.I):
                        print(f"       res {x.get('name')} fmt={x.get('format')} url={x.get('url')}")

    hdr("cammesaweb WPDM packages via wp-json")
    for q in ["potencia", "instalada", "informe mensual", "informe-mensual"]:
        for typ in ("wpdmpro", "wpdm"):
            r = get(f"https://cammesaweb.cammesa.com/wp-json/wp/v2/{typ}", params={"search": q, "per_page": 50},
                    timeout=(20, 90))
            if r is not None and r.ok:
                try:
                    for it in r.json():
                        print(f"  [{typ}:{q}] {it.get('id')} {it.get('date')} {it.get('slug')} {it.get('link')}")
                except ValueError:
                    pass
    hdr("cammesaweb download pages")
    for page in ["https://cammesaweb.cammesa.com/download/potencia-instalada/",
                 "https://cammesaweb.cammesa.com/potencia-instalada/",
                 "https://cammesaweb.cammesa.com/download/informe-mensual_2026-07/",
                 "https://cammesaweb.cammesa.com/download/informe-mensual_2021-09/"]:
        r = get(page, timeout=(20, 90))
        if r is None or not r.ok:
            continue
        t = r.text
        links = sorted(set(re.findall(r"""(?:href|data-downloadurl|src)=["']([^"']+)["']""", t)))
        keep = [u for u in links if re.search(r"wpdmdl|download|\.xls|\.zip|\.csv|iframe|powerbi|tableau|app\.|embed",
                                               u, re.I) and "wp-content/themes" not in u and "plugins" not in u]
        print("  links:", keep[:40])
        for m in re.findall(r"<iframe[^>]+src=[\"']([^\"']+)", t):
            print("  iframe:", m)
        title = re.findall(r"<title>([^<]+)", t)
        print("  title:", title[:1])
        for kw in ("Potencia", "potencia", "xlsx", "Fecha", "Tama"):
            for m in re.finditer(kw, t):
                print("   ctx:", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t[max(0, m.start() - 150): m.start() + 200]))[:300])
                break
        dl = [u for u in keep if "wpdmdl" in u]
        for u in dl[:2]:
            rr = get(u.replace("&#038;", "&"), timeout=(20, 300))
            if rr is not None and rr.ok:
                show_file(rr.content, u)


def colombia():
    gone = ["3HF9", "3GVI", "3HF7", "3GVK", "3HBP", "3MS2", "TR3G", "2ZP9", "3C3Z", "TR2G"]
    hdr("XM ListadoRecursos entries for plants dropping out of CapEfecNeta in Mar 2025")
    r = requests.post("https://servapibi.xm.com.co/lists", json={"MetricId": "ListadoRecursos", "Entity": "Sistema"},
                      timeout=(20, 300))
    ents = {e["Values"].get("Code"): e["Values"] for it in r.json().get("Items", []) for e in it.get("ListEntities", [])}
    for c in gone:
        print(" ", c, ents.get(c))
    hdr("CapEfecNeta presence by day, 2025-02-25..2025-03-08")
    for d in range(12):
        day = dt.date(2025, 2, 25) + dt.timedelta(days=d)
        try:
            j = requests.post("https://servapibi.xm.com.co/daily", json={"MetricId": "CapEfecNeta", "Entity": "Recurso",
                                                                        "StartDate": day.isoformat(),
                                                                        "EndDate": day.isoformat()},
                              timeout=(20, 300)).json()
        except Exception as e:  # noqa: BLE001
            print(day, "failed", e)
            continue
        codes = {e["Code"]: e["Value"] for it in j.get("Items", []) for e in it.get("DailyEntities", [])}
        tot = sum(float(v) for v in codes.values()) / 1000
        print(day, len(codes), f"{tot:,.0f} MW", {c: codes.get(c) for c in gone[:4]})
    hdr("Generation (Gene) of those plants, 2026-08-01..03")
    j = requests.post("https://servapibi.xm.com.co/hourly", json={"MetricId": "Gene", "Entity": "Recurso",
                                                                  "StartDate": "2026-08-01", "EndDate": "2026-08-03"},
                      timeout=(20, 300)).json()
    for it in j.get("Items", []):
        for e in it.get("HourlyEntities", []):
            v = e.get("Values", {})
            if v.get("code") in gone:
                kwh = sum(float(x) for k, x in v.items() if k.startswith("Hour") and x not in (None, ""))
                print(" ", it.get("Date"), v.get("code"), f"{kwh / 1000:.1f} MWh")
    hdr("Other XM metrics listing capacity per plant")
    r = requests.post("https://servapibi.xm.com.co/lists", json={"MetricId": "ListadoMetricas"}, timeout=(20, 120))
    for it in r.json().get("Items", []):
        for e in it.get("ListEntities", []):
            v = e.get("Values", {})
            if re.search(r"capac|instal|potencia", v.get("MetricName", ""), re.I):
                print("  ", v.get("MetricId"), v.get("MetricName"), v.get("Entity"), v.get("MetricUnits"))


if __name__ == "__main__":
    {"argentina": argentina, "colombia": colombia}[sys.argv[1]]()
