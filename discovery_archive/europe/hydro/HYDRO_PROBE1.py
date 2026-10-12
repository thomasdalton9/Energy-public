"""Probe 1 for European hydro reservoir filling: ENTSO-E A72/A16 per zone, plus national sources (NVE, Energiforetagen,
BFE, MITECO, ODRE, E-Control, REN/SNIRH). Prints status and the first bytes; nothing is saved except a summary CSV."""
import os, re, sys, json, time
from datetime import datetime, timezone
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "..", "europe"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", ".."))
import requests
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}


def show(url, n=500, **kw):
    try:
        r = requests.get(url, headers=UA, timeout=60, **kw)
        print(f"\n== {url}\n   HTTP {r.status_code} {r.headers.get('content-type')} len={len(r.content)}")
        print("   " + r.text[:n].replace("\n", " ")[:n])
        return r
    except Exception as e:  # noqa
        print(f"\n== {url}\n   ERR {type(e).__name__}: {str(e)[:150]}")


def links(r, pat):
    if r is None:
        return
    for m in sorted(set(re.findall(r'href="([^"]+)"', r.text))):
        if re.search(pat, m, re.I):
            print("   link:", m)


def entsoe():
    import entsoe_common as ec
    from europe_countries import COUNTRIES
    import xml.etree.ElementTree as ET
    zones = []
    for iso, (name, slug, zs) in COUNTRIES.items():
        zones.append((iso, name, iso, zs[0][1]))
        if len(zs) > 1:
            for z in zs:
                zones.append((z[0], name, z[0], z[1]))
    for lab, name, _, eic in zones:
        try:
            txt = ec.request({"documentType": "A72", "processType": "A16", "in_Domain": eic,
                              "periodStart": "201501010000", "periodEnd": "201601010000"})
        except ec.NoData:
            print(f"ENTSOE {lab:10s} {name}: NO DATA 2015"); continue
        except Exception as e:  # noqa
            print(f"ENTSOE {lab:10s} {name}: ERR {str(e)[:120]}"); continue
        res = sorted(set(re.findall(r"<resolution>(.*?)</resolution>", txt)))
        pts = re.findall(r"<quantity>(.*?)</quantity>", txt)
        ns = len(re.findall(r"<TimeSeries>", txt))
        print(f"ENTSOE {lab:10s} {name}: 2015 ok, series={ns} res={res} points={len(pts)} first={pts[:3]} "
              f"unit={set(re.findall('quantity_Measure_Unit.name>(.*?)<', txt))} psr={set(re.findall('psrType>(.*?)<', txt))}")
        if lab in ("NO", "NO1", "FR"):
            print(txt[:1800])
    # a recent window for every zone (earliest start / latest)
    for lab, name, _, eic in zones:
        try:
            txt = ec.request({"documentType": "A72", "processType": "A16", "in_Domain": eic,
                              "periodStart": "202509010000", "periodEnd": "202610120000"})
            starts = re.findall(r"<start>(.*?)</start>", txt)
            print(f"ENTSOE recent {lab:10s}: starts {starts[:2]} .. {starts[-2:]} n={len(re.findall('<Point>', txt))}")
        except ec.NoData:
            print(f"ENTSOE recent {lab:10s}: NO DATA")
        except Exception as e:  # noqa
            print(f"ENTSOE recent {lab:10s}: ERR {str(e)[:100]}")


def national():
    r = show("https://biapi.nve.no/magasinstatistikk/api/Magasinstatistikk/HentOffentligData", 800)
    if r is not None and r.ok:
        try:
            j = r.json(); print("   NVE rows", len(j), "first", j[0], "last", j[-1])
            print("   keys omrType", sorted({(x['omrType'], x['omrnr']) for x in j}))
            print("   years", min(x['iso_aar'] for x in j), max(x['iso_aar'] for x in j))
        except Exception as e:  # noqa
            print("   parse", e)
    show("https://biapi.nve.no/magasinstatistikk/api/Magasinstatistikk/HentOmråder", 400)
    r = show("https://www.energiforetagen.se/statistik/kraftlaget/", 300); links(r, r"magasin|fyllnad|json|csv|xls")
    show("https://www.energiforetagen.se/api/hydrologi", 200)
    # Opendata.swiss
    for q in ("Speicherseen", "Speicherinhalt", "Speicherkraftwerke Füllungsgrad", "Wasserkraft Speicher wöchentlich"):
        r = show(f"https://ckan.opendata.swiss/api/3/action/package_search?q={q}&rows=8", 100)
        if r is not None and r.ok:
            for p in r.json().get("result", {}).get("results", []):
                print("   CH:", p["name"], "|", (p.get("title") or {}).get("de"), [x.get("url") for x in p.get("resources", [])][:4])
    for u in ("https://www.bfe.admin.ch/bfe/de/home/versorgung/statistik-und-geodaten/energiestatistiken/wasserkraft.html",
              "https://www.bfe.admin.ch/bfe/de/home/versorgung/statistik-und-geodaten/energiestatistiken/elektrizitaetsstatistik.html"):
        r = show(u, 100); links(r, r"speicher|wasser|\.csv|\.xls")
    r = show("https://www.miteco.gob.es/es/agua/temas/evaluacion-de-los-recursos-hidricos/boletin-hidrologico.html", 200)
    links(r, r"boletin|embalse|reserva|xls|csv|json")
    show("https://www.miteco.gob.es/app/BoletinHidrologico/", 200)
    show("https://www.embalses.net/", 200)
    for q in ("stock", "reservoir", "hydraulique", "lac"):
        r = show(f"https://opendata.reseaux-energies.fr/api/explore/v2.1/catalog/datasets?where=search(title,%22{q}%22)&limit=10", 50)
        if r is not None and r.ok:
            for d in r.json().get("results", []):
                print("   ODRE:", d["dataset_id"], "|", d.get("metas", {}).get("default", {}).get("title"))
    r = show("https://www.e-control.at/statistik/e-statistik/archiv/betriebsstatistik", 100); links(r, r"speicher|fuell|xls|csv")
    show("https://www.apg.at/transparency/", 100)
    show("https://transparency.apg.at/api/v1/", 100)
    show("https://oesterreichsenergie.at/", 100)
    r = show("https://oesterreichsenergie.at/energiestatistik/speicherfuellstand", 100); links(r, r"speicher|fuell")
    r = show("https://snirh.apambiente.pt/index.php?idMain=2&idItem=1.3", 100); links(r, r"albufeira|csv|xls")
    show("https://datahub.ren.pt/en/electricity/hydrological-balance/", 200)
    show("https://datahub.ren.pt/service/download/csv/1449?mode=preview", 200)
    show("https://api.ember-energy.org/", 50)


if __name__ == "__main__":
    for f in (entsoe, national):
        try:
            f()
        except SystemExit as e:
            print("EXIT", e)
        except Exception as e:  # noqa
            print("FAIL", f.__name__, type(e).__name__, e)
