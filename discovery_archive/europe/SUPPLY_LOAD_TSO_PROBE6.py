"""Probe 6: Litgrid and EMS (WordPress sites): embedded data sources on the balance page, REST page search for consumption/balance/report documents. Prints only."""
import re
import signal
import sys

import requests

signal.signal(signal.SIGALRM, lambda *_: (_ for _ in ()).throw(TimeoutError("hard timeout")))
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def get(u):
    signal.alarm(60)
    try:
        r = requests.get(u, headers=H, timeout=(10, 40))
        signal.alarm(0)
        print("GET", u[:140], r.status_code, len(r.content), flush=True)
        return r
    except BaseException as e:  # noqa: BLE001
        signal.alarm(0)
        print("GET", u[:140], "ERR", str(e)[:80], flush=True)


r = get("https://www.litgrid.eu/sistema/elektros-energetikos-sistema/elektros-gamybos-ir-vartojimo-balanso-duomenys")
if r is not None and r.ok:
    t = r.text
    for pat in (r'<iframe[^>]+src=["\']?([^"\'\s>]+)', r'(https?://[^"\'\s<>]+(?:powerbi|tableau|datawrapper|flourish|grafana|json|csv|xlsx?|api)[^"\'\s<>]*)', r'data-(?:src|url)=["\']([^"\']+)'):
        for m in sorted(set(re.findall(pat, t, re.I)))[:15]:
            print("   found", m[:170])
    txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", re.sub(r"<script.*?</script>|<style.*?</style>", " ", t, flags=re.S)))
    i = txt.find("balans")
    print("   text:", txt[max(0, i - 100): i + 700])
for host, terms in (("https://www.litgrid.eu", ("balans", "vartojim", "nuostol", "ataskait")), ("https://ems.rs", ("consumption", "balance", "annual report", "losses", "potro", "bilans"))):
    for term in terms:
        r = get(f"{host}/wp-json/wp/v2/search?search={term}&per_page=8")
        if r is not None and r.ok:
            try:
                for it in r.json():
                    print("     ", term, "|", it.get("title", "")[:80], "|", it.get("url", "")[:110])
            except Exception as e:  # noqa: BLE001
                print("      parse", e)
sys.exit(0)
