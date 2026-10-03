"""
Pakistan discovery: ISMO (Independent System and Market Operator, formerly NPCC/CPPA-G) runs a React app at
https://ismo.gov.pk/ whose bundle (static/js/main.*.js) calls '/api/v1/...' (SEA_DISCOVERY_SOUTH_ASIA2.py).
FFD (dam levels) and NEPRA returned 403 to GitHub runners. This lists every API path in the bundle and tries the
ones that look like generation / demand / report data, plus the NTDC and WAPDA pages for downloadable files.
"""
import os
import re
import sys

import requests

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "asia"))
from SEA_DISCOVERY_TH_VN_PH import H, out  # noqa: E402

T = (20, 60)


def main():
    r = requests.get("https://ismo.gov.pk/", headers=H, timeout=T)
    js = [u if u.startswith("http") else "https://ismo.gov.pk" + ("" if u.startswith("/") else "/") + u
          for u in re.findall(r'src="([^"]*main[^"]*\.js)"', r.text)]
    paths = set()
    for u in js:
        t = requests.get(u, headers=H, timeout=T).text
        paths |= set(re.findall(r'["\'`]((?:/api/v1|api/v1)?/[A-Za-z0-9_\-/]{3,80})["\'`]', t))
        base = re.findall(r'(https?://[a-z0-9.\-]+(?::\d+)?/api[^"\'`]*)', t)
        out(f"bundle {u}: {len(t)} bytes; base urls {sorted(set(base))[:10]}")
        for m in re.finditer(r'.{80}api/v1.{160}', t):
            out("  ctx: " + m.group(0).replace("\n", " "))
    api = sorted(p for p in paths if "api" in p or re.search(r"gen|demand|load|report|dispatch|energy|price|data", p, re.I))
    out(f"{len(api)} candidate paths: {api[:150]}")
    for p in api[:40]:
        u = "https://ismo.gov.pk" + (p if p.startswith("/") else "/" + p)
        try:
            x = requests.get(u, headers=dict(H, Accept="application/json"), timeout=T)
            out(f"GET {u} -> {x.status_code} {x.headers.get('content-type')} {len(x.content)} {x.text[:200]!r}")
        except Exception as e:
            out(f"GET {u} ERROR {e!r}")
    for u in ("https://ntdc.gov.pk/power-system-statistics", "https://ntdc.gov.pk/downloads",
              "https://www.wapda.gov.pk/download-category/khabarnama/"):
        try:
            x = requests.get(u, headers=H, timeout=T)
            out(f"GET {u} -> {x.status_code} {len(x.content)}; files " +
                str(sorted(set(re.findall(r'href="([^"]+\.(?:pdf|xlsx?|csv))"', x.text)))[:20]))
        except Exception as e:
            out(f"GET {u} ERROR {e!r}")


if __name__ == "__main__":
    main()
