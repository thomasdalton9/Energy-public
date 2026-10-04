"""
Indonesia raw power data, discovery round 3. Round 2 found the Gatrik SPA's API base
https://gatrik.esdm.go.id/gatrik/gatrik-api/api (EBTKE: https://ebtke.esdm.go.id/api/api/) and that
dashboard.esdm.go.id is a Tableau Server (2023.3). This round lists the API paths the Gatrik/EBTKE bundles
call, tries the statistics / download ones, and asks Tableau whether guest access is on.
"""
import json
import re
import sys

import requests
import urllib3

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from INDONESIA_DISCOVERY1 import H, T, out, probe  # noqa: E402

urllib3.disable_warnings()
GAPI = "https://gatrik.esdm.go.id/gatrik/gatrik-api/api"


def api_paths(site, js):
    s = requests.get(site + js, headers=H, timeout=T, verify=False).text
    paths = set()
    for m in re.finditer(r"""\.(?:get|post|query|put)\(\s*[`"']([^`"']{2,160})[`"']""", s):
        paths.add(m.group(1))
    for m in re.finditer(r"""[`"'](/?[a-z][\w-]*(?:/[\w${}.-]+){0,5})[`"']""", s):
        p = m.group(1)
        if any(k in p.lower() for k in ("statist", "download", "data", "beban", "listrik", "kondisi", "sistem",
                                         "produksi", "dokumen", "publikasi", "infografis", "pembangkit", "kapasitas")):
            paths.add(p)
    out(f"\n#### {site}{js}: {len(paths)} paths")
    for p in sorted(paths):
        out("  P " + p)
    # context around the statistik route component
    for kw in ("statistik", "download_index", "downloadIndex", "dokumenPenting"):
        for m in list(re.finditer(kw, s))[:12]:
            out(f"  [{kw}] ..." + s[max(0, m.start() - 200):m.start() + 300].replace("\n", " "))
    return paths


def gatrik_try(paths):
    tries = [p for p in paths if not p.startswith("http")][:0]
    base_guesses = ["download-index", "download_index", "download-index?kode_category=statistik",
                    "frontend/download_index", "statistik", "data-statistik", "dokumen-penting", "publikasi",
                    "download-index/category", "kategori-download"]
    for p in sorted(set(base_guesses) | {p.lstrip("/") for p in paths if not p.startswith("http") and "$" not in p
                                          and "{" not in p and len(p) < 60}):
        u = f"{GAPI}/{p}"
        try:
            r = requests.get(u, headers=dict(H, Accept="application/json"), timeout=20, verify=False)
        except Exception as e:
            out(f"  {u} ERR {e!r}")
            continue
        ct = r.headers.get("content-type", "")
        snippet = re.sub(r"\s+", " ", r.text[:400]) if "json" in ct else ""
        out(f"  GET {u} -> {r.status_code} {ct} {len(r.content)} {snippet}")


def tableau():
    s = requests.Session()
    s.headers.update(H)
    for ep in ("getSessionInfo", "getServerSettingsUnauthenticated", "getAuthenticationSettings"):
        try:
            r = s.post(f"https://dashboard.esdm.go.id/vizportal/api/web/v1/{ep}", json={"method": ep, "params": {}},
                       headers={"Content-Type": "application/json;charset=UTF-8", "Accept": "application/json"},
                       timeout=T, verify=False)
            out(f"  tableau {ep} -> {r.status_code} {r.text[:1200]}")
        except Exception as e:
            out(f"  tableau {ep} ERR {e!r}")
    probe("tableau serverinfo", "https://dashboard.esdm.go.id/api/3.4/serverinfo", show=600, kw=False)
    probe("tableau signin page", "https://dashboard.esdm.go.id/#/signin", show=0, kw=False)


if __name__ == "__main__":
    try:
        ps = api_paths("https://gatrik.esdm.go.id", "/assets/index.a6066b35.js")
        gatrik_try(ps)
    except Exception as e:
        out(f"!! gatrik {e!r}")
    try:
        api_paths("https://ebtke.esdm.go.id", "/assets/index.975a9c6c.js")
    except Exception as e:
        out(f"!! ebtke {e!r}")
    tableau()
