"""
Probe for middle_east/EGYPT_GAS.py (Oct 2026): (1) JODI jodidb.org viewer, Egypt only, every balance row, million m3,
all months, in one download; (2) CAPMAS IndicatorFilter JSON for the gas indicators (full payload, to see unit fields);
(3) MOEE ST_fuel.aspx tables. Workflow (manual): discovery_archive/workflows/egypt_gas_pull_probe.yml.

RUN 1 (2026-10-12, Actions run 38190691497):
- MOEE ST_fuel.aspx: one table (42 x 9): Chart | Year ('2023/2022' = FY2022/23 ... '1982/1981') | H.F.O Ktons |
  N.G. Million m3 (34,142 in FY2022/23) | L.F.O tons | Coke ktons | Nafatia ktons | L.F.O ktons ('17٫3', Arabic decimal
  sign) | Total ktoe. No Last-Modified header.
- CAPMAS: IndicatorDetails?indicatorId=&publicationId= gives measureUnit (EN): energy balance 3112/3102/3108/3121/3124
  'Million cubic meter', 3155 '1000 tons of equivalent oil', petroleum-sector 9543/9544/9545/5937/5939 'Quantity in 1000
  tons' (indicator-wide); yearTypeId=2, periodic 'Annually'. 3112 production 2014-2022 (2016 = 168, 2018/2020 missing),
  3102 to 2019, 9543 filters total/Houses/Other/Industry/Governmental/Car service..., 5937/5939 natural gas rows.
- JODI viewer: one download with Country=[h], BALANCE all, Unit 0, TIME all works (CSV link in OnDownload). Country
  handles follow the country listing order (handle 28 = Egypt; 29 = Equatorial Guinea). Item counts '95,1,14,3,211'
  = Country, Product, BALANCE, Unit, TIME.
"""
import html
import json
import re
import sys
import time
from io import StringIO

import pandas as pd
import requests


def out(*a):
    print(*a, flush=True)


H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data/1.0"}
B = "http://www.jodidb.org"
REPORT = "38673"
s = requests.Session(); s.headers.update(H)


def inputs(t):
    o = {}
    for m in re.finditer(r"<(input|select|textarea)\b((?:[^>\"]|\"[^\"]*\")*)>", t, flags=re.I | re.S):
        a = dict(re.findall(r'(\w+)\s*=\s*"([^"]*)"', m.group(2)))
        if "name" in a:
            o[a["name"]] = html.unescape(a.get("value", ""))
    return o


def jodi():
    out("#### JODI")
    ref = {"Referer": f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}"}
    r = s.get(f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}", timeout=300)
    base = inputs(r.text); rv0 = base["sWD_ReportView"]
    out("counts", base.get("sWD_SelectedItemsCount"))
    out("RV head", rv0[:3000])

    def build(sel):
        rv = rv0
        for name, hs in sel.items():
            if hs is None:
                continue
            d = '<Definition><Items type="handles"><String value="%s"/></Items></Definition>' % ",".join(map(str, hs))
            rv, n = re.subn(r'(<Dim name="%s">.*?)<Definition>.*?</Definition>' % re.escape(name), lambda m: m.group(1) + d, rv, count=1, flags=re.S)
            assert n == 1
        return rv

    def dl(sel):
        data = dict(base); data["sWD_ReportView"] = build(sel); data["WD_Command"] = ""
        r = s.post(f"{B}/TableViewer/tableView.aspx", data=data, headers=ref, timeout=300)
        f = inputs(r.text)
        data = dict(f); data.update({"WD_DownloadFormat": "CSV", "sCS_SpawnWindow": "True", "WD_PageBeforeSummary": "/TableViewer/tableView.aspx"})
        r = s.post(f"{B}/TableViewer/downloadPrompt.aspx", data=data, headers=ref, timeout=300)
        pt = r.text
        urls = re.findall(r"OnDownload\(\s*['\"]([^'\"]+)['\"]\s*\)", pt)
        ivt = re.findall(r"OnDownloadIvt\(\s*['\"]([^'\"]*)['\"]\s*,\s*['\"]([^'\"]*)['\"]\s*,\s*['\"]([^'\"]*)['\"]\s*\)", pt)
        out("  prompt urls", urls[:2], "ivt", ivt[:2])
        if urls:
            u = urls[0]; u = u if u.startswith("http") else B + (u if u.startswith("/") else "/TableViewer/" + u)
            r = s.get(u, headers={"Referer": f"{B}/TableViewer/downloadPrompt.aspx"}, timeout=300)
        elif ivt:
            mime, name, path = ivt[0]
            pf = inputs(pt); pf.update({"WD_Command": "Download", "WD_DownloadMimeType": mime, "WD_DownloadFilename": name, "WD_DownloadFilePath": path})
            r = s.post(f"{B}/TableViewer/download.aspx", data=pf, headers={"Referer": f"{B}/TableViewer/downloadPrompt.aspx"}, timeout=300)
        else:
            return None
        return r.content.decode("utf-8-sig", "replace")

    c = dl({"Product": [0], "BALANCE": [0], "Unit": [0], "TIME": [0]})
    out("country list:", (c or "")[:200])
    c = dl({"Country": [29], "Product": [0], "BALANCE": None, "Unit": [0], "TIME": None})
    out("=== EGYPT handle 29, all balance, mcm, all time:")
    out(c)
    c = dl({"Country": [29], "Product": [0], "BALANCE": list(range(14)), "Unit": [0], "TIME": list(range(200, 215))})
    out("=== EGYPT explicit handles 0-13, time 200-214:")
    out(c)


def capmas():
    out("#### CAPMAS")
    r = requests.get("https://www.capmas.gov.eg:8080/api/Subject/SubSubjectWithIndicator/30", headers=H, timeout=60)
    pubs = (r.json().get("data") or {}).get("publicationWithIndicators") or []
    for pb in pubs:
        out("PUB", {k: v for k, v in pb.items() if not isinstance(v, (list, dict))})
        for k, v in pb.items():
            if isinstance(v, list):
                for it in v:
                    if isinstance(it, dict):
                        flat = {kk: vv for kk, vv in it.items() if not isinstance(vv, (list, dict))}
                        tr = it.get("indicatorTranslations") or []
                        en = [t for t in tr if isinstance(t, dict) and t.get("locale") == "en"]
                        if re.search(r"غاز|gas", json.dumps(it, ensure_ascii=False), re.I):
                            out("  ITEM", json.dumps(flat, ensure_ascii=False)[:600], "| EN", json.dumps(en, ensure_ascii=False)[:400])
    for iid in [9543, 9544, 9545, 5937, 5939, 3112, 3102, 3108, 3155, 3121, 3124]:
        r = requests.get(f"https://www.capmas.gov.eg:8080/api/Indicator/IndicatorFilter?IndicatorId={iid}&SubSubjectId=30", headers=H, timeout=60)
        out(f"\n=== IndicatorFilter {iid}: {r.status_code} {len(r.content)}B")
        out(r.text[:6000])
        for pid in (57, 77):
            r2 = requests.get(f"https://www.capmas.gov.eg:8080/api/Indicator/IndicatorDetails?indicatorId={iid}&publicationId={pid}", headers=H, timeout=60)
            out(f"  details pub{pid}: {r2.status_code} {r2.text[:1500]}")


def moee():
    out("#### MOEE")
    r = requests.get("http://www.moee.gov.eg/english_new/ST_fuel.aspx", headers=H, timeout=60)
    out(r.status_code, len(r.content), r.headers.get("last-modified"))
    for i, t in enumerate(pd.read_html(StringIO(r.text))):
        out(f"table {i} {t.shape}")
        out(t.to_string(max_colwidth=40)[:8000])


for f in (moee, capmas, jodi):
    try:
        f()
    except Exception as e:  # noqa: BLE001
        out("ERROR", f.__name__, type(e).__name__, str(e)[:300])
