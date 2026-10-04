"""
Indonesia discovery round 9: (1) test asia/INDONESIA_EBTKE_CAPACITY.py end to end (--out to a temp path, run
twice to check the incremental merge, print the sheets; add_charts' generic capacity chart is tried too);
(2) last tries at dashboard.esdm.go.id (Tableau, guest enabled): guessed public view URLs with
?:showVizHome=no and .csv exports, and the vizportal guest calls after loading the home page.
"""
import os
import subprocess
import sys
import tempfile

import pandas as pd
import requests
import urllib3

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from INDONESIA_DISCOVERY1 import H, T, out  # noqa: E402

urllib3.disable_warnings()
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def test_pull():
    tmp = os.path.join(tempfile.mkdtemp(), "indonesia_renewable_capacity.xlsx")
    for i in range(2):
        r = subprocess.run([sys.executable, os.path.join(ROOT, "asia", "INDONESIA_EBTKE_CAPACITY.py"), "--out", tmp],
                           capture_output=True, text=True, timeout=600)
        out(f"---- run {i + 1}: exit {r.returncode}\n{r.stdout[-3000:]}\n{r.stderr[-3000:]}")
    if os.path.exists(tmp):
        for sh in ("Units", "Monthly", "Detail"):
            df = pd.read_excel(tmp, sheet_name=sh)
            out(f"---- sheet {sh} {df.shape}\n{df.to_string()[:4000]}")
        sys.path.insert(0, ROOT)
        try:
            import add_charts
            spec = add_charts.power_capacity("Indonesia renewable capacity (ESDM EBTKE)")(tmp)
            out(f"add_charts.power_capacity -> {len(spec)} chart spec(s) OK")
        except Exception as e:
            out(f"add_charts check: {e!r}")


def tableau():
    s = requests.Session()
    s.headers.update(H)
    r = s.get("https://dashboard.esdm.go.id/", timeout=T, verify=False)
    out(f"home {r.status_code} cookies {list(s.cookies.get_dict())}")
    tok = s.cookies.get("XSRF-TOKEN", "")
    hdr = {"Content-Type": "application/json;charset=UTF-8", "Accept": "application/json", "X-XSRF-TOKEN": tok}
    for ep in ("getSessionInfo", "getViews"):
        x = s.post(f"https://dashboard.esdm.go.id/vizportal/api/web/v1/{ep}",
                   json={"method": ep, "params": {"page": {"startIndex": 0, "maxItems": 100}}}, headers=hdr,
                   timeout=T, verify=False)
        out(f"  {ep} -> {x.status_code} {x.text[:500]}")
    guesses = ["Ketenagalistrikan/Dashboard", "DashboardKetenagalistrikan/Dashboard", "Listrik/Dashboard",
               "Kelistrikan/Dashboard", "DashboardESDM/Dashboard", "EBTKE/Dashboard", "Pembangkit/Dashboard",
               "DashboardPNBP/Dashboard", "PNBP/Dashboard", "Migas/Dashboard", "Minerba/Dashboard",
               "DashboardKESDM/Dashboard1", "Gatrik/Dashboard1", "Ketenagalistrikan/Sheet1"]
    for g in guesses:
        for suf in ("?:showVizHome=no&:embed=y", ".csv"):
            u = f"https://dashboard.esdm.go.id/views/{g}{suf}"
            try:
                x = s.get(u, timeout=T, verify=False, allow_redirects=False)
                out(f"  {u} -> {x.status_code} {x.headers.get('content-type')} {len(x.content)} "
                    f"loc={x.headers.get('location')}")
            except Exception as e:
                out(f"  {u} ERR {e!r}")


if __name__ == "__main__":
    for f in (test_pull, tableau):
        try:
            f()
        except Exception as e:
            out(f"!! {e!r}")
