"""
Indonesia raw power data, discovery round 8. Round 7 found:
  - ESDM One Map (geoportal.esdm.go.id, ArcGIS portal "monaresia") publishes MapServers incl.
    gis5/rest/services/Sebaran_Pembangkit_Listrik (power plants), PLTP, PLTS/PLTM under 10 MW, Offgrid APBN,
    gis2/.../DBP/Pembangkit_Berbasis_Bioenergi, and a web map "Sub Sektor Ketenagalistrikan".
  - EBTKE /api/api/konten/data-angka: latest-month renewable capacity by type (MW), history routes need login.
  - Gatrik API is a CMS only (news, documents) - no data series.
This round: layer metadata + fields + record counts + capacity totals by type for the power-plant services
(is it a usable capacity dataset?), the web map's layer list, EBTKE data-angka in full, SIM EBTKE portal,
and the esdm.go.id main.js highlight AJAX.
"""
import json
import re
import sys

import requests
import urllib3

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from INDONESIA_DISCOVERY1 import H, T, out, probe  # noqa: E402

urllib3.disable_warnings()
G = "https://geoportal.esdm.go.id"
SERVICES = ["gis5/rest/services/Sebaran_Pembangkit_Listrik/MapServer", "gis5/rest/services/PLTP/MapServer",
            "gis5/rest/services/PLTS_under_10_MW/MapServer", "gis5/rest/services/PLTM_PLTMH_Under_10MW/MapServer",
            "gis5/rest/services/Pembangkit_Offgrid_APBN/MapServer",
            "gis2/rest/services/DBP/Pembangkit_Berbasis_Bioenergi/MapServer"]


def gj(url, **params):
    params.setdefault("f", "json")
    r = requests.get(url, params=params, headers=H, timeout=90, verify=False)
    return r.json()


def service(path):
    base = f"{G}/{path}"
    try:
        meta = gj(base)
    except Exception as e:
        out(f"\n#### {path}: ERR {e!r}")
        return
    out(f"\n#### {path}: {meta.get('serviceDescription', '')[:300]!r} desc={meta.get('description', '')[:300]!r}")
    for lyr in meta.get("layers", []):
        lid = lyr["id"]
        try:
            lm = gj(f"{base}/{lid}")
            cnt = gj(f"{base}/{lid}/query", where="1=1", returnCountOnly="true")
        except Exception as e:
            out(f"  layer {lid} ERR {e!r}")
            continue
        fields = [(f["name"], f.get("type", "")[13:]) for f in lm.get("fields", [])]
        out(f"  layer {lid} {lyr.get('name')!r} geom={lm.get('geometryType')} count={cnt.get('count')} "
            f"maxRecords={lm.get('maxRecordCount')} editInfo={lm.get('editingInfo')}")
        out(f"    fields {fields}")
        try:
            q = gj(f"{base}/{lid}/query", where="1=1", outFields="*", returnGeometry="false", resultRecordCount=5)
            for ft in q.get("features", [])[:5]:
                out("    REC " + json.dumps(ft.get("attributes"), ensure_ascii=False)[:900])
        except Exception as e:
            out(f"    sample ERR {e!r}")
        # capacity totals by a type-like field
        names = [f[0] for f in fields]
        tf = next((n for n in names if re.search(r"jenis|tipe|type|pembangkit|kategori|energi", n, re.I)
                   and not re.search(r"kap|mw|daya", n, re.I)), None)
        cf = next((n for n in names if re.search(r"kapasitas|kap_|mw|daya", n, re.I)), None)
        out(f"    type field {tf}, capacity field {cf}")
        if tf:
            stats = [{"statisticType": "count", "onStatisticField": names[0], "outStatisticFieldName": "n"}]
            if cf:
                stats.append({"statisticType": "sum", "onStatisticField": cf, "outStatisticFieldName": "mw"})
            try:
                q = gj(f"{base}/{lid}/query", where="1=1", groupByFieldsForStatistics=tf,
                       outStatistics=json.dumps(stats), returnGeometry="false")
                for ft in q.get("features", []):
                    out("    BYTYPE " + json.dumps(ft.get("attributes"), ensure_ascii=False))
                if "error" in q:
                    out(f"    stats error {q['error']}")
            except Exception as e:
                out(f"    stats ERR {e!r}")


def webmap():
    try:
        d = gj(f"{G}/monaresia/sharing/rest/content/items/81e64b6c88d342f993f121bb85ac8d90/data")
        for l in d.get("operationalLayers", []):
            out(f"  WEBMAP layer {l.get('title')} | {l.get('url')}")
    except Exception as e:
        out(f"  webmap ERR {e!r}")
    try:
        d = gj(f"{G}/monaresia/sharing/rest/search", q="owner:* AND (type:\"Feature Service\" OR type:\"Map Service\")",
               num=100, sortField="modified", sortOrder="desc")
        out(f"  portal total services {d.get('total')}")
        for it in d.get("results", []):
            out(f"  SVC {it.get('title')} | {it.get('url')}")
    except Exception as e:
        out(f"  portal ERR {e!r}")


def ebtke():
    r = requests.get("https://ebtke.esdm.go.id/api/api/konten/data-angka", headers=H, timeout=T, verify=False)
    out("\n#### EBTKE data-angka\n" + r.text[:4000])
    for p in ("konten/data-angka?tahun=2025", "konten/data-angka?bulan=Desember&tahun=2025",
              "konten/data-angka-kapasitas-pembangkit", "konten/kapasitas-pembangkit"):
        try:
            x = requests.get(f"https://ebtke.esdm.go.id/api/api/{p}", headers=H, timeout=T, verify=False)
            out(f"  GET {p} -> {x.status_code} {x.text[:400]}")
        except Exception as e:
            out(f"  GET {p} ERR {e!r}")
    probe("SIM EBTKE", "https://simebtke.esdm.go.id/", show=0, links=60)


def highlight_js():
    s = requests.get("https://www.esdm.go.id/themes/v2/js/main.js", headers=H, timeout=T, verify=False).text
    out(f"\n#### main.js {len(s)}")
    for m in re.finditer(r"highlight|ajax|url\s*:", s, re.I):
        out("  ..." + s[max(0, m.start() - 200):m.start() + 400].replace("\n", " "))


if __name__ == "__main__":
    for p in SERVICES:
        service(p)
    for f in (webmap, ebtke, highlight_js):
        try:
            f()
        except Exception as e:
            out(f"!! {e!r}")
