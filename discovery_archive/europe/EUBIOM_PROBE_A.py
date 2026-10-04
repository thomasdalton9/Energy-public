"""EU biomethane probe A (GB IT SE AT EBA) (prints only): AGGM attributes, National Gas portal, GOV.UK/Ofgem, GSE/Snam, Energimyndigheten/SCB, EBA, GIE map, Eurostat."""
import json
import re
import sys

import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}


def get(url, **kw):
    try:
        r = requests.get(url, headers=kw.pop("headers", H), timeout=40, **kw)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"  ERR {url[:100]}: {type(e).__name__} {str(e)[:80]}")
        return None


def show(url, n=300, pat=None):
    r = get(url)
    if r is None:
        return None
    print(f"GET {url[:130]} -> {r.status_code} {r.headers.get('content-type','')[:40]} len={len(r.content)}")
    if pat:
        for m in list(re.finditer(pat, r.text, re.I))[:12]:
            s = max(0, m.start() - 120)
            print("    ...", r.text[s:m.end() + 160].replace("\n", " "))
    elif n:
        print("   ", r.text[:n].replace("\n", " "))
    return r


def aggm():
    print("== AGGM")
    r = get("https://platform.aggm.at/vis-service/api/ts/attributes")
    if r is None or not r.ok:
        return
    txt = r.text
    print(len(txt))
    for m in list(re.finditer(r"bio|gr[uü]n|green|wasserstoff|hydrogen|einspeis", txt, re.I))[:40]:
        print("   ", txt[max(0, m.start() - 100):m.end() + 100].replace("\n", " "))
    j = r.json()
    print(list(j.keys()))
    a = j.get("attributes", {})
    for k, v in a.items():
        print("  attr", k, str(v)[:300])


def natgas():
    print("== National Gas")
    for u in ["https://data.nationalgas.com/api/find-gas-data", "https://data.nationalgas.com/api/find-gas-data-items",
              "https://data.nationalgas.com/api/get-gas-data-items", "https://data.nationalgas.com/api/find-gas-data/items",
              "https://data.nationalgas.com/api/find-gas-data-menu", "https://data.nationalgas.com/api/data-items",
              "https://data.nationalgas.com/find-gas-data", "https://data.nationalgas.com/api/find-gas-data/categories"]:
        show(u, 200)
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(user_agent=UA)
        seen = []

        def on(resp):
            try:
                if "nationalgas" in resp.url and "json" in resp.headers.get("content-type", ""):
                    t = resp.text()
                    seen.append((resp.url, t))
            except Exception:  # noqa: BLE001
                pass
        pg.on("response", on)
        pg.goto("https://data.nationalgas.com/find-gas-data", wait_until="networkidle", timeout=60000)
        pg.wait_for_timeout(5000)
        for u, t in seen:
            print("  XHR", u[:140], len(t), "bio-hits", len(re.findall("bio", t, re.I)))
            for m in list(re.finditer(r"bio|embedded|green gas", t, re.I))[:15]:
                print("     ", t[max(0, m.start() - 100):m.end() + 140].replace("\n", " "))
        print("  page text bio:", [m.group(0) for m in re.finditer(r".{40}biomethane.{40}", pg.inner_text("body"), re.I | re.S)][:5])
        b.close()


def gb():
    print("== GB gov.uk / ofgem")
    for q in ["biomethane injection", "Green Gas Support Scheme statistics", "biomethane Energy Trends", "RHI biomethane injection monthly"]:
        r = get("https://www.gov.uk/api/search.json?count=12&fields=title&fields=link&fields=public_timestamp&q=" + requests.utils.quote(q))
        if r is not None and r.ok:
            for x in r.json().get("results", []):
                print("  ", q, "|", x.get("title"), x.get("link"), x.get("public_timestamp"))
    show("https://www.ofgem.gov.uk/search?keyword=biomethane%20injection", 0, r"href=\"[^\"]*(?:green-gas|biomethane)[^\"]*\"")
    show("https://www.ofgem.gov.uk/environmental-and-social-schemes/green-gas-support-scheme-ggss", 0, r"href=\"[^\"]*(?:xlsx|csv|data|report)[^\"]*\"")
    show("https://www.ofgem.gov.uk/environmental-and-social-schemes/non-domestic-renewable-heat-incentive-rhi/data", 0, r"href=\"[^\"]*(?:xlsx|csv|biomethane)[^\"]*\"")
    show("https://www.gov.uk/government/statistics/energy-trends-section-4-gas", 0, r"href=\"[^\"]*(?:xlsx|ods)[^\"]*\"")
    show("https://www.gov.uk/government/statistics/gas-section-4-energy-trends", 0, r"href=\"[^\"]*(?:xlsx|ods)[^\"]*\"")
    show("https://www.renewableenergyassociation.org/", 0, r"biomethane")
    show("https://www.biomethane.org.uk/", 200)
    show("https://www.greengas.org.uk/", 200)
    show("https://www.energy-uk.org.uk/", 100)
    show("https://assets.publishing.service.gov.uk/media/", 100)


def it():
    print("== Italy")
    for u in ["https://www.gse.it/dati-e-scenari/statistiche", "https://www.gse.it/servizi-per-te/fonti-rinnovabili/biometano",
              "https://www.snam.it/it/energia-e-gas/biometano/", "https://www.snam.it/it/trasporto/dati-biometano",
              "https://www.snam.it/en/energy_transition/biomethane/", "https://www.consorziobiogas.it/", "https://dati.mase.gov.it/",
              "https://www.arera.it/dati-e-statistiche", "https://www.gse.it/documenti_site/Documenti%20GSE/Studi%20e%20scenari/"]:
        show(u, 0, r"href=\"[^\"]*(?:biometano|xlsx|statisti)[^\"]*\"")


def se():
    print("== Sweden")
    for u in ["https://www.energimyndigheten.se/statistik/den-officiella-statistiken/statistikprodukter/produktion-och-anvandning-av-biogas-och-rotrest/",
              "https://www.energimyndigheten.se/en/news/?type=statistics",
              "https://pxexternal.energimyndigheten.se/api/v1/sv/Energimyndigheten/", "https://pxexternal.energimyndigheten.se/api/v1/en/",
              "https://api.scb.se/OV0104/v1/doris/en/ssd/EN/EN0105", "https://api.scb.se/OV0104/v1/doris/en/ssd/EN/EN0107",
              "https://www.energigas.se/fakta-om-energigas/statistik/", "https://www.energigas.se/"]:
        show(u, 300, r"biogas|biometan")


def eba():
    print("== EBA / GIE")
    for u in ["https://www.europeanbiogas.eu/eba-statistical-report-2024/", "https://www.europeanbiogas.eu/eba-statistical-report-2025/",
              "https://www.europeanbiogas.eu/eba-statistical-report-2023/", "https://www.europeanbiogas.eu/category/publications/",
              "https://www.europeanbiogas.eu/european-biogas-association-statistical-report/"]:
        show(u, 0, r"href=\"[^\"]*(?:xlsx|pdf|statistical)[^\"]*\"")
    show("https://www.gie.eu/transparency-platform/biomethane-map/", 0, r"href=\"[^\"]*(?:xlsx|csv|api|biomethane)[^\"]*\"|api[^ ]{0,60}")
    show("https://www.gie.eu/transparency-platform/biomethane-map/#/", 200)
    for u in ["https://bm.gie.eu/", "https://biomethane.gie.eu/", "https://www.gie.eu/biomethane/"]:
        show(u, 300)


def eurostat():
    print("== Eurostat")
    base = "https://ec.europa.eu/eurostat/api/dissemination/sdmx/2.1/"
    r = get(base + "codelist/ESTAT/SIEC?format=JSON") or None
    show(base + "dataflow/ESTAT/all?detail=allstubs", 0, r"[^\n]{0,80}biomethane[^\n]{0,80}|nrg_cb_[a-z]+")
    for ds in ["nrg_bal_c", "nrg_cb_bm", "nrg_cb_gas", "nrg_cb_gasm", "nrg_ind_pehcf", "nrg_bio"]:
        r = get(f"https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/{ds}?format=JSON&lang=EN&geo=SE&time=2022&sinceTimePeriod=2022&freq=A")
        print("  ", ds, r.status_code if r is not None else None, (r.text[:200] if r is not None else ""))
    r = get("https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_bal_c?format=JSON&lang=EN&geo=SE&time=2022&freq=A&siec=R5300")
    if r is not None and r.ok:
        j = r.json()
        print("  nrg_bal_c dims", {k: len(v['category']['index']) for k, v in j['dimension'].items()})
        print("  nrg_bal ids", [k for k in j['dimension']['nrg_bal']['category']['label'].items() if 'iogas' in k[1] or 'ethane' in k[1] or 'GIC' in k[0]][:30])
    r = get("https://ec.europa.eu/eurostat/api/dissemination/sdmx/2.1/codelist/ESTAT/SIEC/?format=TSV")
    if r is not None:
        for m in re.finditer(r"[^\n]*(?:iogas|iomethane|R53)[^\n]*", r.text):
            print("   SIEC", m.group(0)[:150])


if __name__ == "__main__":
    for f in (aggm, natgas, gb, it, se, eba, eurostat):
        try:
            f()
        except Exception as e:  # noqa: BLE001
            print("FAILED", f.__name__, type(e).__name__, e)
        sys.stdout.flush()
