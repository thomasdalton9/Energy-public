"""
Round 6 of CENTRAL_AMERICA_POWER_DISCOVERY.py: is Panama's natural-gas (LNG) use
published anywhere machine-readable?

Checked already: CND's monthly operations report (Informe Mensual de Operaciones,
PDF) has fuel *prices* and plant lists but no fuel volumes; CND's daily report has
gas-fired energy (MWh) only.

This round crawls the Secretaria de Energia (energia.gob.pa) statistics pages
(hydrocarbon imports / consumption, where LNG would appear) and ASEP, saves every
page and every xls/xlsx/csv found whose name or link text mentions gas / GNL /
importacion / hidrocarburos, and prints the links.
"""
import os
import re
import sys
import urllib.parse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import CENTRAL_AMERICA_POWER_DISCOVERY as d1  # noqa: E402
from CENTRAL_AMERICA_POWER_DISCOVERY2 import Tee  # noqa: E402

KEYS = ["estadist", "hidrocarb", "gas", "gnl", "importac", "combustib", "consumo", "balance", "datos", "mercado"]
FILES = re.compile(r"\.(xlsx?|csv|ods)(\?|$)", re.I)


def crawl(starts, host_part, depth=2, limit=120):
    seen, queue, n = set(), [(u, 0) for u in starts], 0
    files = []
    while queue and n < limit:
        u, dep = queue.pop(0)
        if u in seen:
            continue
        seen.add(u)
        n += 1
        r = d1.get(u, "PAGE")
        if r is None or "html" not in r.headers.get("content-type", ""):
            continue
        for v in d1.links(r):
            low = urllib.parse.unquote(v).lower()
            if FILES.search(low) and any(k in low for k in KEYS):
                files.append(v)
                print("    FILE " + v)
            elif host_part in urllib.parse.urlparse(v).netloc and dep < depth and any(k in low for k in KEYS) \
                    and not re.search(r"\.(pdf|png|jpe?g|gif|svg|css|js|docx?|zip)(\?|$)", low):
                queue.append((v, dep + 1))
    for f in sorted(set(files))[:40]:
        d1.get(f, "DOWNLOAD")


if __name__ == "__main__":
    os.makedirs(d1.RAW, exist_ok=True)
    log = open(os.path.join(d1.RAW, "log_round6.txt"), "w")
    sys.stdout = Tee(sys.__stdout__, log)
    try:
        crawl(["https://www.energia.gob.pa/", "https://www.energia.gob.pa/estadisticas/",
               "https://www.energia.gob.pa/hidrocarburos/", "https://www.energia.gob.pa/mercado-de-hidrocarburos/"],
              "energia.gob.pa")
        crawl(["https://www.asep.gob.pa/", "https://www.asep.gob.pa/electricidad/estadisticas/"], "asep.gob.pa",
              depth=1, limit=40)
    except Exception as e:  # noqa: BLE001
        print(f"FAILED: {type(e).__name__}: {e}", flush=True)
    sys.stdout.flush()
    sys.stdout = sys.__stdout__
    log.close()
