"""Probe 2 for European hydro reservoir filling: BFE ogd17 CSV, MITECO BD-Embalses zip and weekly bulletin, ENTSO-E
earliest dates for zones that did not answer for 2015, multi-year window limit, Energiforetagen PDF, REN DataHub."""
import io, os, re, sys, zipfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "..", "europe"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", ".."))
import requests
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}


def get(u, **kw):
    try:
        r = requests.get(u, headers=UA, timeout=120, **kw)
        print(f"\n== {u}\n   HTTP {r.status_code} {r.headers.get('content-type')} len={len(r.content)}")
        return r
    except Exception as e:  # noqa
        print(f"\n== {u}\n   ERR {type(e).__name__} {str(e)[:150]}")


def swiss():
    r = get("https://www.bfe-ogd.ch/ogd17/ogd17_fuellungsgrad_speicherseen.csv")
    if r is not None and r.ok:
        t = r.content.decode("utf-8-sig", "replace").splitlines()
        print("   rows", len(t)); print("\n".join(t[:6])); print("...\n" + "\n".join(t[-4:]))
    r = get("https://ckan.opendata.swiss/api/3/action/package_show?id=enstat-fullungsgrad-der-speicherseen-sonntag-24h-wochenbericht-speicherinhalt")
    if r is not None and r.ok:
        d = r.json()["result"]; print("   notes:", (d.get("description") or {}).get("de", "")[:1200])


def spain():
    r = get("https://www.miteco.gob.es/content/dam/miteco/es/agua/temas/evaluacion-de-los-recursos-hidricos/boletin-hidrologico/Historico-de-embalses/BD-Embalses.zip")
    if r is not None and r.ok:
        z = zipfile.ZipFile(io.BytesIO(r.content))
        for i in z.infolist():
            print("   zip:", i.filename, i.file_size)
        for i in z.infolist()[:2]:
            b = z.read(i)
            print("   ---", i.filename); print(b[:1500].decode("latin-1", "replace"))
            print("   ...tail:", b[-600:].decode("latin-1", "replace"))
    r = get("https://www.miteco.gob.es/es/agua/temas/evaluacion-de-los-recursos-hidricos/boletin-hidrologico.html")
    if r is not None and r.ok:
        t = re.sub(r"<[^>]+>", " ", r.text); t = re.sub(r"\s+", " ", t)
        for m in re.finditer(r"hidroel|Hidroel|BD-Embalses|actualiz", t):
            print("   ctx:", t[max(0, m.start() - 150): m.start() + 250]); 
        for m in sorted(set(re.findall(r'(?:src|href|data-[a-z-]+)="([^"]*(?:json|csv|xls|api|arcgis|embalse)[^"]*)"', r.text, re.I))):
            print("   attr:", m)


def sweden():
    r = get("https://www.energiforetagen.se/globalassets/energiforetagen/statistik/kraftlaget/aktuellt-magasinslage-sverige-veckorapport.pdf")
    if r is None:
        r = get("https://www.energiforetagen.se/4a8157/globalassets/energiforetagen/statistik/kraftlaget/aktuellt-magasinslage-sverige-veckorapport.pdf")
    if r is not None and r.ok:
        open("/tmp/se.pdf", "wb").write(r.content)
        try:
            import pdfplumber
            with pdfplumber.open("/tmp/se.pdf") as p:
                print("   pages", len(p.pages)); print(p.pages[0].extract_text()[:1500])
        except Exception as e:  # noqa
            print("   pdf", e)
    r = get("https://www.energiforetagen.se/statistik/kraftlaget/")
    if r is not None and r.ok:
        for m in sorted(set(re.findall(r'(?:href|src)="([^"]*(?:xls|csv|json|api|pdf|magasin)[^"]*)"', r.text, re.I))):
            print("   attr:", m)


def entsoe():
    import entsoe_common as ec
    from europe_countries import COUNTRIES
    # 2-year window allowed?
    try:
        t = ec.request({"documentType": "A72", "processType": "A16", "in_Domain": "10YFR-RTE------C",
                        "periodStart": "201601010000", "periodEnd": "201901010000"})
        print("FR 3-year window ok, points", len(re.findall("<Point>", t)))
    except Exception as e:  # noqa
        print("FR 3-year window:", str(e)[:150])
    first = {}
    for iso in ("GR", "RS", "MK", "AL", "LT", "SI", "BA", "EE", "DK", "DE", "NL", "CZ", "HU", "PL", "SK", "BE", "IE", "XK"):
        eic = COUNTRIES[iso][2][0][1]
        found = None
        for y in range(2016, 2027):
            try:
                t = ec.request({"documentType": "A72", "processType": "A16", "in_Domain": eic,
                                "periodStart": f"{y}01010000", "periodEnd": f"{y+1}01010000"})
                pts = len(re.findall("<Point>", t)); st = re.findall(r"<start>(.*?)</start>", t)[:1]
                found = (y, st, pts); break
            except ec.NoData:
                continue
            except Exception as e:  # noqa
                print(iso, y, "ERR", str(e)[:100]); break
        print("EARLIEST", iso, found)
    # points per year for sparse ones
    for iso in ("SI", "LT", "GR"):
        eic = COUNTRIES[iso][2][0][1]
        for y in (2020, 2022, 2024, 2025, 2026):
            try:
                t = ec.request({"documentType": "A72", "processType": "A16", "in_Domain": eic,
                                "periodStart": f"{y}01010000", "periodEnd": f"{y+1}01010000"})
                print("YEAR", iso, y, len(re.findall("<Point>", t)), re.findall(r"<resolution>(.*?)</resolution>", t)[:3])
            except Exception as e:  # noqa
                print("YEAR", iso, y, str(e)[:60])
    # latest week of FR / NO1 / CH: when does the last point fall
    for iso in ("FR", "NO", "CH", "ES", "AT", "PT", "IT"):
        eic = COUNTRIES[iso][2][0][1]
        t = ec.request({"documentType": "A72", "processType": "A16", "in_Domain": eic,
                        "periodStart": "202609010000", "periodEnd": "202610130000"})
        print("LATEST", iso, re.findall(r"<start>(.*?)</start>", t)[:2], re.findall(r"<end>(.*?)</end>", t)[:2],
              re.findall(r"<quantity>(.*?)</quantity>", t)[-3:], len(re.findall("<Point>", t)), re.findall(r"<resolution>(.*?)</resolution>", t)[:3])


def ren():
    r = get("https://datahub.ren.pt/en/electricity/")
    if r is not None and r.ok:
        for m in sorted(set(re.findall(r'href="([^"]+)"', r.text))):
            if re.search(r"hidr|storage|reserv|balance|water", m, re.I):
                print("   link:", m)


if __name__ == "__main__":
    for f in (swiss, spain, sweden, ren, entsoe):
        try:
            f()
        except SystemExit as e:
            print("EXIT", e)
        except Exception as e:  # noqa
            print("FAIL", f.__name__, type(e).__name__, e)
