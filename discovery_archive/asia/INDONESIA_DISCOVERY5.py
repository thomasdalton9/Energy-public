"""
Indonesia raw power data, discovery round 5. Round 4: Gatrik's download index holds only the annual
Statistik Ketenagalistrikan books (PDF, 2006-2024; 2024 edition published 2025-09-15) - no monthly series.
dashboard.esdm.go.id (Tableau) said guestEnabled but vizportal calls returned 401 without the XSRF header.
This round: (1) Tableau guest calls with X-XSRF-TOKEN; (2) open the 2024 statistics book and look for any
monthly tables (Januari..Desember, beban puncak, produksi) and the generation-by-plant-type tables.
"""
import io
import re
import sys

import requests
import urllib3

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from INDONESIA_DISCOVERY1 import H, T, out  # noqa: E402

urllib3.disable_warnings()
STAT = "https://gatrik.esdm.go.id/gatrik/gatrik-api//storage/migrasi/download_index/files/"


def tableau():
    s = requests.Session()
    s.headers.update(H)
    base = "https://dashboard.esdm.go.id/vizportal/api/web/v1/"
    s.post(base + "getServerSettingsUnauthenticated", json={"method": "getServerSettingsUnauthenticated", "params": {}},
           timeout=T, verify=False)
    tok = s.cookies.get("XSRF-TOKEN")
    hdr = {"Content-Type": "application/json;charset=UTF-8", "Accept": "application/json", "X-XSRF-TOKEN": tok or ""}
    for ep, params in (("getSessionInfo", {}),
                       ("getViews", {"page": {"startIndex": 0, "maxItems": 200}, "order": [{"field": "name", "ascending": True}]}),
                       ("getWorkbooks", {"page": {"startIndex": 0, "maxItems": 200}}),
                       ("getProjects", {"page": {"startIndex": 0, "maxItems": 200}})):
        try:
            r = s.post(base + ep, json={"method": ep, "params": params}, headers=hdr, timeout=T, verify=False)
            out(f"  tableau {ep} -> {r.status_code} {r.text[:4000]}")
        except Exception as e:
            out(f"  tableau {ep} ERR {e!r}")
    for u in ("https://dashboard.esdm.go.id/views/Ketenagalistrikan/Dashboard", "https://dashboard.esdm.go.id/#/explore"):
        try:
            r = s.get(u, timeout=T, verify=False)
            out(f"  GET {u} -> {r.status_code} {len(r.content)} {r.text[:300]!r}")
        except Exception as e:
            out(f"  GET {u} ERR {e!r}")


def book(fname):
    import pdfplumber
    r = requests.get(STAT + fname, headers=H, timeout=180, verify=False)
    out(f"\n#### {fname}: {r.status_code} {len(r.content)} bytes")
    if not r.ok:
        return
    pdf = pdfplumber.open(io.BytesIO(r.content))
    out(f"  pages {len(pdf.pages)}")
    months = re.compile(r"Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember|"
                        r"January|February|March|June|July|August|October|December", re.I)
    shown = 0
    for i, p in enumerate(pdf.pages):
        t = p.extract_text() or ""
        head = " / ".join(t.splitlines()[:3])[:220]
        nm = len(months.findall(t))
        flag = ""
        if nm >= 6:
            flag += f" MONTHS={nm}"
        if re.search(r"jenis pembangkit|type of power plant", t, re.I) and re.search(r"produksi|production", t, re.I):
            flag += " GEN-BY-TYPE"
        if re.search(r"beban puncak|peak load", t, re.I):
            flag += " PEAK"
        out(f"  p{i + 1}: {head}{flag}")
        if ("MONTHS" in flag or "GEN-BY-TYPE" in flag) and shown < 14:
            shown += 1
            out("    ----- page text -----\n    " + t[:3500].replace("\n", "\n    "))


if __name__ == "__main__":
    try:
        tableau()
    except Exception as e:
        out(f"!! tableau {e!r}")
    try:
        book("91fa8-buku-statistik-ketenagalistrikan-2024.pdf")
    except Exception as e:
        out(f"!! book {e!r}")
