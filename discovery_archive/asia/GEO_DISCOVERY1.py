"""
Geo-blocking test (owner's option 3): which regions can reach the grid-operator sites that time out or reset from
GitHub's US runners? Uses check-host.net's free API (HTTP checks from many countries, no account):
  GET https://check-host.net/check-http?host=<url>&max_nodes=60   (Accept: application/json) -> request_id, nodes
  GET https://check-host.net/check-result/<request_id>             -> per node: [[ok, seconds, message, code, ip]]
Targets: EVN reservoir feed (Vietnam), NSMO (Vietnam), Grid-India / MERIT (India), PLN (Indonesia), BPS (Indonesia),
CEB GenSum (Sri Lanka), plus one control site.
"""
import time

import requests

TARGETS = [
    "https://hochuathuydien.evn.com.vn/PageHoChuaThuyDienEmbedEVN.aspx",
    "https://www.nsmo.vn/HeThongDien",
    "https://grid-india.in/en/",
    "https://meritindia.in/",
    "https://web.pln.co.id/",
    "https://www.bps.go.id/",
    "https://cebcare.ceb.lk/GenSum/GetEnergySummary",
    "https://www.evn.com.vn/",
]
H = {"Accept": "application/json", "User-Agent": "Mozilla/5.0"}


def out(*a):
    print(*a, flush=True)


def check(url):
    r = requests.get("https://check-host.net/check-http", params={"host": url, "max_nodes": 60}, headers=H,
                     timeout=60)
    j = r.json()
    rid, nodes = j.get("request_id"), j.get("nodes", {})
    time.sleep(25)
    res = {}
    for _ in range(6):
        res = requests.get(f"https://check-host.net/check-result/{rid}", headers=H, timeout=60).json()
        if res and all(v is not None for v in res.values()):
            break
        time.sleep(10)
    ok, bad = [], []
    for node, v in sorted(res.items()):
        info = nodes.get(node, [])
        where = f"{info[0] if info else ''} {info[1] if len(info) > 1 else ''} ({node.split('.')[0]})"
        if not v or v[0] is None:
            bad.append(f"{where}: no result")
            continue
        r0 = v[0]
        code = r0[3] if len(r0) > 3 else None
        (ok if r0[0] == 1 else bad).append(f"{where}: {code or ''} {r0[2] if len(r0) > 2 else ''}"[:90])
    out(f"\n==== {url}\n  OK from {len(ok)} nodes:")
    for x in ok:
        out("    " + x)
    out(f"  FAILED from {len(bad)} nodes:")
    for x in bad:
        out("    " + x)


def main():
    for u in TARGETS:
        try:
            check(u)
        except Exception as e:  # noqa: BLE001
            out(f"!! {u}: {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
