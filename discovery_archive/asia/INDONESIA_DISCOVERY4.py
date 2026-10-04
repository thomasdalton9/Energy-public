"""
Indonesia raw power data, discovery round 4. Round 3: the Gatrik API's /download-index returns all 626
published documents as JSON (no auth); /dokumen-penting gives headline indicators. dashboard.esdm.go.id
(Tableau Server) has guestEnabled=true. This round lists every download-index document (title, category,
date, file), finds the file base URL in the DownloadIndex chunk, and tries listing Tableau views as guest.
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


def gatrik():
    j = requests.get(f"{GAPI}/download-index", headers=H, timeout=60, verify=False).json()
    rows = j["data"]
    out(f"#### download-index: {len(rows)} rows; keys {list(rows[0].keys())}")
    out(json.dumps(rows[0], ensure_ascii=False)[:1500])
    cats = {}
    for r in rows:
        k = r.get("kategori_download_index_id")
        cats.setdefault(k, []).append(r)
    for k, rs in sorted(cats.items(), key=lambda kv: str(kv[0])):
        catname = ""
        for key in ("kategori", "kategori_download_index", "category"):
            if isinstance(rs[0].get(key), dict):
                catname = rs[0][key].get("nama") or rs[0][key].get("judul") or str(rs[0][key])[:80]
        out(f"\n== category {k} {catname} ({len(rs)})")
        for r in rs:
            out(f"  {str(r.get('tgl_upload'))[:10]} | {r.get('judul')} | {r.get('deskripsi')} | {r.get('dokumen')} | "
                f"{r.get('nama_dokumen')} | {r.get('slug')}")
    for p in ("kategori-download-index", "download-index/kategori", "kategori-download", "statistic", "statistik",
              "download-index/detail/statistik-ketenagalistrikan", "infografis", "data-ketenagalistrikan"):
        try:
            r = requests.get(f"{GAPI}/{p}", headers=H, timeout=30, verify=False)
            out(f"  GET {p} -> {r.status_code} {r.text[:600]}")
        except Exception as e:
            out(f"  GET {p} ERR {e!r}")
    for js in ("DownloadIndex.ccc01027.js", "DetailDownloadIndex.fd55955c.js"):
        s = requests.get(f"https://gatrik.esdm.go.id/assets/{js}", headers=H, timeout=T, verify=False).text
        out(f"\n#### {js} ({len(s)})")
        for m in re.finditer(r"""[`"']([^`"']*(?:dokumen|storage|uploads|download|kategori|http)[^`"']*)[`"']""", s):
            out("  S " + m.group(1)[:200])


def tableau():
    s = requests.Session()
    s.headers.update(H)
    hdr = {"Content-Type": "application/json;charset=UTF-8", "Accept": "application/json"}
    base = "https://dashboard.esdm.go.id/vizportal/api/web/v1/"
    for ep, params in (("getViews", {"page": {"startIndex": 0, "maxItems": 200}}),
                       ("getWorkbooks", {"page": {"startIndex": 0, "maxItems": 200}}),
                       ("getProjects", {"page": {"startIndex": 0, "maxItems": 200}}),
                       ("getSiteNamesAcrossAllPods", {"page": {"startIndex": 0, "maxItems": 50}}),
                       ("getSessionInfo", {})):
        try:
            r = s.post(base + ep, json={"method": ep, "params": params}, headers=hdr, timeout=T, verify=False)
            out(f"  tableau {ep} -> {r.status_code} {r.text[:2500]}")
        except Exception as e:
            out(f"  tableau {ep} ERR {e!r}")
    out(f"  cookies {s.cookies.get_dict()}")
    # REST API: guest has no sign-in; try a sign-in with empty creds just to see the error
    for u in ("https://dashboard.esdm.go.id/api/3.21/sites", "https://dashboard.esdm.go.id/api/3.21/serverinfo"):
        probe("tableau rest", u, show=800, kw=False, links=0)


if __name__ == "__main__":
    for f in (gatrik, tableau):
        try:
            f()
        except Exception as e:
            out(f"!! {e!r}")
