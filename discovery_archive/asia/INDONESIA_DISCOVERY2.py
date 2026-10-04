"""
Indonesia raw power data, discovery round 2. Round 1 found:
  - esdm.go.id homepage "ESDM Highlight": daily oil/gas cards plus an "Electricity Grid Status" card (systems
    normal/alert/deficit per region) with a Detail link - does the detail give daya mampu / beban puncak?
  - gatrik.esdm.go.id and ebtke.esdm.go.id are Vite SPAs: find their JSON API in the JS bundle.
  - dashboard.esdm.go.id is a Tableau Server.
  - BPS pages are behind Cloudflare (403); webapi needs a key. PLN statistics page moved.
"""
import re
import sys
from urllib.parse import urljoin

import requests
import urllib3

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from INDONESIA_DISCOVERY1 import H, T, out, probe  # noqa: E402

urllib3.disable_warnings()


def esdm_highlight():
    r = requests.get("https://www.esdm.go.id/id", headers=H, timeout=T, verify=False)
    t = r.text
    i = t.find("Highlight #1")
    out(f"\n#### ESDM highlight raw HTML ({len(t)} chars, at {i})")
    out(t[i - 2000:i + 14000])
    for m in re.finditer(r"""href=["']([^"']*(highlight|status|listrik|kelistrikan|sistem)[^"']*)["']""", t, re.I):
        out("  HL-LINK " + urljoin(r.url, m.group(1)))


def spa(base):
    r = requests.get(base, headers=H, timeout=T, verify=False)
    js = re.findall(r"""src=["']([^"']+\.js)["']""", r.text)
    out(f"\n#### SPA {base}: js {js}")
    for j in js:
        u = urljoin(base, j)
        try:
            s = requests.get(u, headers=H, timeout=T, verify=False).text
        except Exception as e:
            out(f"  ERROR {e!r}")
            continue
        out(f"  {u}: {len(s)} chars")
        found = set()
        for m in re.finditer(r"""["'`](https?://[^"'`\s]{6,200}|/api/[^"'`\s]{1,200}|api/[^"'`\s]{1,200})["'`]""", s):
            found.add(m.group(1))
        for f in sorted(found)[:300]:
            out("   URL " + f)
        for kw in ("baseURL", "VITE_", "download_index", "statistik", "beban", "kondisi"):
            for m in list(re.finditer(kw, s))[:6]:
                out(f"   [{kw}] ..." + s[max(0, m.start() - 150):m.start() + 250].replace("\n", " "))
        # chunked imports
        for c in sorted(set(re.findall(r"""["'](\./|/assets/)([\w.-]+\.js)["']""", s)))[:400]:
            pass


def tableau():
    for u in ("https://dashboard.esdm.go.id/", "https://dashboard.esdm.go.id/#/views",
              "https://dashboard.esdm.go.id/api/2.8/serverinfo", "https://dashboard.esdm.go.id/api/3.4/serverinfo",
              "https://dashboard.esdm.go.id/vizportal/api/web/v1/getSessionInfo",
              "https://dashboard.esdm.go.id/views", "https://dashboard.esdm.go.id/t/"):
        probe("tableau", u, show=1500, links=20, kw=False)


def pln():
    for u in ("https://www.pln.co.id/stakeholder/statistik", "https://www.pln.co.id/id/stakeholder/laporan-statistik",
              "https://www.pln.co.id/investor-relations/statistik", "https://www.pln.co.id/webapi/",
              "https://www.pln.co.id/sitemap.xml", "https://www.pln.co.id/id"):
        r = probe("PLN", u, show=300, links=0, kw=False)
        if r is not None and r.ok:
            for m in sorted(set(re.findall(r"""(/[\w/-]*(?:statisti|report|laporan)[\w/.-]*)""", r.text, re.I)))[:60]:
                out("  PATH " + m)


if __name__ == "__main__":
    for f in (esdm_highlight, lambda: spa("https://gatrik.esdm.go.id/"), lambda: spa("https://ebtke.esdm.go.id/"),
              tableau, pln):
        try:
            f()
        except Exception as e:
            out(f"!! {e!r}")
