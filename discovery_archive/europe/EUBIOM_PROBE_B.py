"""Biomethane probe 2 (prints only): National Gas biomethane catalogue + values, AGGM bio series history, Eurostat nrg_bal_c biogases,
Snam / ET / Ofgem / GIE / EBA / Swedish PX-web tables."""
import io
import json
import re
import sys
from datetime import date, timedelta

import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}


def get(url, **kw):
    try:
        return requests.get(url, headers=kw.pop("headers", H), timeout=kw.pop("timeout", 60), **kw)
    except Exception as e:  # noqa: BLE001
        print(f"  ERR {url[:100]}: {type(e).__name__} {str(e)[:80]}")


def links(url, pat, maxn=40, ctx=0):
    r = get(url)
    if r is None:
        return None
    print(f"GET {url[:120]} -> {r.status_code} len={len(r.text)}")
    seen = set()
    for m in re.finditer(r'href="([^"]+)"[^>]*>([^<]{0,100})', r.text):
        h, tx = m.group(1), m.group(2).strip()
        if re.search(pat, h + " " + tx, re.I) and h not in seen:
            seen.add(h)
            print("   LINK", h[:160], "|", tx[:80])
            if len(seen) >= maxn:
                break
    return r


def natgas():
    print("== National Gas catalogue")
    r = get("https://data.nationalgas.com/api/search-everywhere", headers={**H, "Referer": "https://data.nationalgas.com/find-gas-data"}, timeout=120)
    ids = {}
    if r is not None and r.ok:
        j = r.json()

        def walk(n, path):
            if isinstance(n, dict):
                nm = n.get("name", "")
                p = path + [nm]
                meta = " ".join(n.get("meta", []) or []) if isinstance(n.get("meta"), list) else ""
                m = re.match(r"(PUBOB\w+)", meta.strip())
                if m and any("iomethane" in x or "iogas" in x or "mbedded" in x for x in p):
                    ids[m.group(1)] = " > ".join(p)[-170:]
                for v in n.values():
                    walk(v, p)
            elif isinstance(n, list):
                for x in n:
                    walk(x, path)
        walk(j, [])
        print(len(ids), "biomethane-ish ids")
        for k, v in sorted(ids.items()):
            print("  ", k, v)
    # values
    bodyids = [k for k, v in ids.items() if "Energy" in v]
    print("energy ids", bodyids)
    if bodyids:
        d1 = date.today()
        for dt, d0 in (("GASDAY", d1 - timedelta(days=20)),):
            body = {"latestFlag": "Y", "applicableFor": "Y", "dateFrom": d0.isoformat(), "dateTo": d1.isoformat(), "dateType": dt, "ids": ",".join(bodyids)}
            rr = requests.post("https://data.nationalgas.com/api/find-gas-data", json=body, timeout=120,
                               headers={"Content-Type": "application/json", "Accept": "application/json", "User-Agent": UA, "Referer": "https://data.nationalgas.com/find-gas-data/view"})
            print("  values", rr.status_code, rr.text[:1500])
        body["dateFrom"], body["dateTo"] = "2023-01-01", "2023-01-10"
        rr = requests.post("https://data.nationalgas.com/api/find-gas-data", json=body, timeout=120,
                           headers={"Content-Type": "application/json", "Accept": "application/json", "User-Agent": UA, "Referer": "https://data.nationalgas.com/find-gas-data/view"})
        print("  values 2023", rr.status_code, rr.text[:800])


def aggm():
    print("== AGGM biomethane series")
    url = "https://platform.aggm.at/vis-service/api/ts/values"
    names = ["EntryBiogasOst_MGM-Allokationen", "EntryBiogasTirol_MGM-Allokationen", "EntryBiogasVorarlberg_MGM-Allokationen",
             "EntryBiogasproduktionOst_Kapazitaetgenutzt"]
    for frm, to in (("2019-01-01", "2019-02-01"), ("2021-01-01", "2021-02-01"), ("2023-01-01", "2023-02-01"), ("2026-08-01", "2026-09-01")):
        body = {"rangeType": "individual", "from": f"{frm}T06:00:00", "to": f"{to}T06:00:00", "granularity": "day", "timeseries": names}
        r = requests.post(url, json=body, headers={"User-Agent": UA, "Content-Type": "application/json", "Accept": "application/json"}, timeout=90)
        print(frm, r.status_code, len(r.text))
        try:
            for cd in r.json()["timeSeriesData"]["chartData"]:
                ys = [p["y"] for p in cd["dataSet"] if p.get("y") is not None]
                print("   ", cd["header"]["name"], len(ys), "sum GWh", round(sum(ys) / 1e6, 1), "unit", cd["header"].get("unit"))
        except Exception as e:  # noqa: BLE001
            print("   ", e, r.text[:200])
    # other aggm attributes with bio / EntryBio
    r = get("https://platform.aggm.at/vis-service/api/ts/attributes")
    t = r.text
    print("names with Bio:", sorted(set(re.findall(r'"name":"([^"]*[Bb]io[^"]*)"', t))))


def eurostat():
    print("== Eurostat")
    B = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/"
    r = get(B + "nrg_bal_c?format=JSON&lang=EN&geo=EU27_2020&siec=R5300&unit=TJ_GCV&freq=A&sinceTimePeriod=2018", timeout=90)
    print(" EU27 R5300 TJ_GCV", r.status_code if r is not None else None)
    if r is not None and r.ok:
        j = r.json()
        print("  dims", {k: len(v["category"]["index"]) for k, v in j["dimension"].items()}, "units", j["dimension"]["unit"]["category"]["label"])
        labs = j["dimension"]["nrg_bal"]["category"]["label"]
        for k in list(labs):
            print("   ", k, labs[k])
    for siec in ("R5300", "R5301SB", "R5300P", "R5300B", "R5300S", "R5300SP"):
        r = get(B + f"nrg_bal_c?format=JSON&lang=EN&geo=EU27_2020&geo=SE&geo=DE&geo=FR&geo=IT&geo=DK&siec={siec}&nrg_bal=IPRD&nrg_bal=GIC&nrg_bal=NRG_BIOG_E&nrg_bal=FEC&nrg_bal=TI_EHG_MAPE_E&nrg_bal=TI_EHG_MAPCHP_E&nrg_bal=TO&unit=GWH&freq=A&sinceTimePeriod=2020", timeout=90)
        if r is None or not r.ok:
            print(" ", siec, r.status_code if r is not None else None, (r.text[:200] if r is not None else ""))
            continue
        j = r.json()
        dims = [k for k in j["id"]]
        sz = j["size"]
        print(" ", siec, dims, sz, "n values", len(j.get("value", {})))
        geo = list(j["dimension"]["geo"]["category"]["index"])
        nb = list(j["dimension"]["nrg_bal"]["category"]["index"])
        tm = list(j["dimension"]["time"]["category"]["index"])
        # index: freq, nrg_bal, siec, unit, geo, time
        import itertools
        for (a, b, c) in itertools.product(range(len(nb)), range(len(geo)), range(len(tm))):
            idx = ((a * 1 * 1) * len(geo) + b) * len(tm) + c
            v = j["value"].get(str(idx))
            if v is not None and nb[a] in ("IPRD", "GIC", "FEC", "TO"):
                print("    ", nb[a], geo[b], tm[c], v)
    # dataflows with biogas / biomethane in the name
    r = get("https://ec.europa.eu/eurostat/api/dissemination/sdmx/2.1/dataflow/ESTAT/all?detail=allstubs", timeout=120)
    if r is not None:
        for m in re.finditer(r'id="([A-Z0-9_]+)"[^>]*>(?:<c:Name xml:lang="[a-z]+">[^<]*</c:Name>)*?<c:Name xml:lang="en">([^<]*)</c:Name>', r.text):
            if re.search(r"biogas|biomethane|gas.*(?:balance|supply)|gaseous", m.group(2), re.I):
                print("  DF", m.group(1), m.group(2)[:120])


def snam_gb_misc():
    print("== Snam / GB / misc")
    links("https://www.snam.it/en/energy_transition/biomethane/", r"xls|csv|biomet|dati|data|pdf", 40)
    links("https://www.snam.it/it/trasporto/dati-biometano", r"xls|csv|biomet|dati|pdf", 40)
    r = get("https://www.snam.it/en/energy_transition/biomethane/")
    if r is not None:
        txt = re.sub(r"<[^>]+>", " ", r.text)
        for m in list(re.finditer(r"(?:GWh|Smc|mcm|million|MWh|TWh|plants|impiant)", txt))[:12]:
            print("   TXT", re.sub(r"\s+", " ", txt[max(0, m.start() - 150):m.end() + 100]))
    # DESNZ ET 4.x - look for biomethane row
    for u in ["https://assets.publishing.service.gov.uk/media/6aba6b15fceb6fb3a65012ba/ET_4.1_SEP_26.xlsx",
              "https://assets.publishing.service.gov.uk/media/6aba6b21fe72ed1e2b02f149/ET_4.2_SEP_26.xlsx"]:
        r = get(u)
        if r is not None and r.ok:
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(r.content), data_only=True)
            for ws in wb:
                for row in ws.iter_rows(values_only=True):
                    s = " ".join(str(c) for c in row if c is not None)
                    if re.search(r"bio", s, re.I):
                        print("   ", u[-16:], ws.title, s[:300])
    links("https://www.ofgem.gov.uk/environmental-and-social-schemes/green-gas-support-scheme-and-green-gas-levy", r"xls|csv|data|report|statist|public", 40)
    links("https://www.ofgem.gov.uk/environmental-and-social-schemes/renewable-heat-incentive-rhi", r"xls|csv|data|report|statist|public", 30)
    links("https://www.gov.uk/government/publications/green-gas-support-scheme-ggss-expenditure-forecast-statements-and-tariff-change-notices", r"xls|ods|csv|pdf|forecast", 20)
    links("https://www.gov.uk/government/publications/green-gas-support-scheme-budget-management", r"xls|ods|csv|pdf", 20)
    links("https://www.biomethane.org.uk/", r"xls|csv|data|stat|report|plants", 20)
    links("https://www.renewableenergyassociation.org/", r"biomethane|green-gas|stat", 20)


def gie_eba():
    print("== GIE / EBA")
    links("https://www.gie.eu/publications/maps/european-biomethane-map/", r"xls|csv|api|biomethane|map|data|pdf", 40)
    links("https://www.gie.eu/biomethane/", r"xls|csv|api|biomethane|map|data|pdf", 40)
    links("https://www.europeanbiogas.eu/eba-statistical-report-2025/", r"xls|csv|pdf|release|press|biomethane|statistical", 30)
    links("https://www.europeanbiogas.eu/eba-statistical-report-2024/", r"xls|csv|pdf|release|press|biomethane|statistical", 30)
    r = get("https://www.europeanbiogas.eu/wp-content/uploads/2024/12/EBA_stats_report_complete_241204_preview.pdf", timeout=120)
    print("EBA preview pdf", r.status_code if r is not None else None, len(r.content) if r is not None else 0)
    if r is not None and r.ok:
        open("eba_preview.pdf", "wb").write(r.content)
        try:
            import subprocess
            subprocess.run([sys.executable, "-m", "pip", "install", "-q", "pypdf"], check=False)
            from pypdf import PdfReader
            rd = PdfReader("eba_preview.pdf")
            print(" pages", len(rd.pages))
            for i, pg in enumerate(rd.pages[:40]):
                tx = pg.extract_text() or ""
                if re.search(r"biomethane", tx, re.I):
                    print(f" --- page {i+1}:", re.sub(r"\s+", " ", tx)[:1200])
        except Exception as e:  # noqa: BLE001
            print(" pdf err", e)


def sweden():
    print("== Sweden PX")
    for base in ("https://pxexternal.energimyndigheten.se/api/v1/en/", "https://api.scb.se/OV0104/v1/doris/en/ssd/EN/EN0105", "https://api.scb.se/OV0104/v1/doris/en/ssd/EN/EN0107",
                 "https://api.scb.se/OV0104/v1/doris/en/ssd/EN/"):
        r = get(base)
        print(base, r.status_code if r is not None else None, (r.text[:600] if r is not None else ""))
    for base in ("https://pxexternal.energimyndigheten.se/api/v1/en/Energimyndigheten/", "https://pxexternal.energimyndigheten.se/api/v1/en/energimyndigheten/",
                 "https://pxexternal.energimyndigheten.se/api/v1/sv/Energimyndigheten/"):
        r = get(base)
        print(base, r.status_code if r is not None else None, (r.text[:600] if r is not None else ""))
    links("https://www.energimyndigheten.se/statistik/den-officiella-statistiken/statistikprodukter/produktion-och-anvandning-av-biogas-och-rotrest/", r"xls|csv|biogas|pxweb|statistik|px", 30)
    links("https://www.energigas.se/om-oss/nyhetsrum/oekad-anvandning-och-produktion-av-biometan-2025/", r"xls|pdf|biometan|statistik", 20)
    r = get("https://www.energigas.se/om-oss/nyhetsrum/oekad-anvandning-och-produktion-av-biometan-2025/")
    if r is not None:
        print(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))[1500:5000])


if __name__ == "__main__":
    for f in (aggm, natgas, eurostat, snam_gb_misc, gie_eba, sweden):
        try:
            f()
        except Exception as e:  # noqa: BLE001
            print("FAILED", f.__name__, type(e).__name__, e)
        sys.stdout.flush()
