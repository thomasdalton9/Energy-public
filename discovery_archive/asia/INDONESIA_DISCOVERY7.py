"""
Indonesia raw power data, discovery round 7 (owner: "look for ESDM"). Digs into ESDM's own systems:
  1. Gatrik SPA: fetch every lazy-loaded chunk and list all API routes it calls (query/get/post strings),
     then GET the public-looking ones on https://gatrik.esdm.go.id/gatrik/gatrik-api/api.
  2. EBTKE SPA: same, on https://ebtke.esdm.go.id/api/api/.
  3. ESDM /id/highlight page: find the AJAX endpoints behind the "Time series" box and the electricity
     status table (does it keep history / MW?).
  4. geoportal.esdm.go.id ArcGIS portal ("monaresia"): search items for power plants (pembangkit).
  5. Hostname sweep for ESDM data portals (http + https).
"""
import json
import re
import sys
from urllib.parse import urljoin

import requests
import urllib3

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from INDONESIA_DISCOVERY1 import H, T, out, probe  # noqa: E402

urllib3.disable_warnings()
KEY = re.compile(r"statist|realisasi|produksi|kapasitas|pembangkit|beban|neraca|sistem|rasio|kinerja|listrik|"
                 r"energi|capaian|data|dashboard|grafik|chart|indikator|bauran|ebt|monitor|bulan", re.I)


def spa_routes(site, api):
    idx = requests.get(site, headers=H, timeout=T, verify=False).text
    main = re.findall(r"""src=["'](/assets/index\.[\w]+\.js)["']""", idx)
    if not main:
        out(f"!! no main bundle at {site}")
        return
    s = requests.get(site.rstrip("/") + main[0], headers=H, timeout=T, verify=False).text
    chunks = sorted(set(re.findall(r"""assets/([\w.-]+\.js)""", s)))
    out(f"\n#### {site}: main {main[0]}, {len(chunks)} chunks")
    routes = {}
    for c in [main[0].split("/")[-1]] + chunks:
        try:
            cs = s if c in main[0] else requests.get(f"{site.rstrip('/')}/assets/{c}", headers=H, timeout=T,
                                                     verify=False).text
        except Exception as e:
            out(f"  {c} ERR {e!r}")
            continue
        for m in re.finditer(r"""\.(?:query|get|post|put|request)\(\s*[`"']([^`"']{2,160})[`"']""", cs):
            routes.setdefault(m.group(1), set()).add(c)
        for m in re.finditer(r"""[`"'](/[a-z][\w-]*(?:/[\w${}.-]*){0,5})[`"']""", cs):
            p = m.group(1)
            if KEY.search(p) and not p.startswith("/assets"):
                routes.setdefault(p, set()).add(c)
    for r in sorted(routes):
        out(f"  ROUTE {r}   <- {sorted(routes[r])[:3]}")
    tried = 0
    for r in sorted(routes):
        if "${" in r or ("store" in r) or ("update" in r) or ("delete" in r) or tried > 150:
            continue
        rr = r.strip("/")
        u = f"{api.rstrip('/')}/{rr}"
        tried += 1
        try:
            x = requests.get(u, headers=dict(H, Accept="application/json"), timeout=20, verify=False)
            ct = x.headers.get("content-type", "")
            sn = re.sub(r"\s+", " ", x.text[:500]) if "json" in ct else ""
            out(f"  GET {u} -> {x.status_code} {len(x.content)} {sn}")
        except Exception as e:
            out(f"  GET {u} ERR {e!r}")


def highlight():
    r = requests.get("https://www.esdm.go.id/id/highlight", headers=H, timeout=T, verify=False)
    t = r.text
    out(f"\n#### highlight page {len(t)}")
    for m in re.finditer(r"<script[^>]*>(.*?)</script>", t, re.S | re.I):
        js = m.group(1)
        if re.search(r"ajax|fetch|url|\$\.get|\$\.post", js, re.I) and len(js) < 20000:
            out("  SCRIPT: " + re.sub(r"\s+", " ", js)[:4000])
    for m in re.finditer(r"""<script[^>]+src=["']([^"']+)["']""", t):
        out("  SRC " + urljoin(r.url, m.group(1)))
    for m in re.finditer(r"""<form[^>]*>""", t):
        out("  FORM " + m.group(0))
    i = t.find("Status Ketenagalistrikan")
    out("  KEL-HTML: " + re.sub(r"\s+", " ", t[i - 500:i + 4000]))


def geoportal():
    for u in ("https://geoportal.esdm.go.id/monaresia/sharing/rest/search?q=pembangkit&f=json&num=50",
              "https://geoportal.esdm.go.id/monaresia/sharing/rest/search?q=listrik&f=json&num=50",
              "https://geoportal.esdm.go.id/monaresia/sharing/rest/search?q=ketenagalistrikan&f=json&num=50",
              "https://geoportal.esdm.go.id/server/rest/services?f=json",
              "https://geoportal.esdm.go.id/gis/rest/services?f=json",
              "https://geoportal.esdm.go.id/arcgis/rest/services?f=json",
              "https://geoportal.esdm.go.id/monaresia/rest/services?f=json"):
        try:
            x = requests.get(u, headers=H, timeout=T, verify=False)
            out(f"\n  GEO {u} -> {x.status_code} {x.headers.get('content-type')} {len(x.content)}")
            try:
                j = x.json()
                if "results" in j:
                    out(f"   total {j.get('total')}")
                    for it in j["results"][:50]:
                        out(f"   ITEM {it.get('type')} | {it.get('title')} | {it.get('url')} | {it.get('id')} | "
                            f"{it.get('modified')}")
                else:
                    out("   " + json.dumps(j)[:3000])
            except Exception:
                out("   " + re.sub(r"\s+", " ", x.text[:600]))
        except Exception as e:
            out(f"  GEO {u} ERR {e!r}")


def hosts():
    names = ["satudata", "data", "dataesdm", "pusdatin", "statistik", "opendata", "ppid", "dashboard-moms",
             "dashboard-gatrik", "gatrik-dashboard", "simon", "simontok", "rptka", "listrik", "ketenagalistrikan",
             "djk", "djlpe", "sigap", "e-ketenagalistrikan", "ruptl", "ampere", "monitoring", "simebtke",
             "ebtkeconex", "dashboard-ebtke", "neraca", "heesi", "web-statistik", "dataenergi", "energi",
             "one", "onemap", "portal"]
    for n in names:
        for scheme in ("https", "http"):
            u = f"{scheme}://{n}.esdm.go.id/"
            try:
                x = requests.get(u, headers=H, timeout=15, verify=False)
                title = re.search(r"<title>(.*?)</title>", x.text, re.S | re.I)
                out(f"  HOST {u} -> {x.status_code} {len(x.content)} final={x.url} title={title.group(1).strip()[:80] if title else ''}")
                break
            except requests.exceptions.ConnectionError as e:
                if "resolve" in str(e) or "Name or service" in str(e):
                    out(f"  HOST {u} -> no DNS")
                    break
                out(f"  HOST {u} -> conn err {str(e)[:120]}")
            except Exception as e:
                out(f"  HOST {u} -> {e!r}"[:200])


if __name__ == "__main__":
    for f in (lambda: spa_routes("https://gatrik.esdm.go.id/", "https://gatrik.esdm.go.id/gatrik/gatrik-api/api"),
              lambda: spa_routes("https://ebtke.esdm.go.id/", "https://ebtke.esdm.go.id/api/api"),
              highlight, geoportal, hosts):
        try:
            f()
        except Exception as e:
            out(f"!! {e!r}")
