"""One-off probe: can GitHub reach the Energy Institute downloads page (plain requests and curl_cffi browser
impersonation)? Lists its data links and, if reachable, the panel CSV columns."""
import io
import re
import pandas as pd
PAGE = "https://www.energyinst.org/statistical-review/resources-and-data-downloads"
try:
    from curl_cffi import requests as cr
except ImportError:
    cr = None
for imp in ("chrome", "safari", "firefox"):
    if cr is None:
        break
    try:
        r = cr.get(PAGE, impersonate=imp, timeout=60)
        print(imp, r.status_code, len(r.text))
        if r.status_code == 200:
            links = sorted(set(re.findall(r'href="([^"]+\.(?:csv|xlsx|xls)[^"]*)"', r.text, flags=re.I)))
            for l in links:
                print("LINK", l)
            for l in links:
                if re.search(r"panel", l, re.I) and l.lower().endswith(".csv"):
                    x = cr.get(requests_url := ("https://www.energyinst.org" + l if l.startswith("/") else l),
                               impersonate=imp, timeout=300)
                    print("PANEL", x.status_code, len(x.content))
                    d = pd.read_csv(io.BytesIO(x.content), low_memory=False, encoding_errors="replace")
                    print(list(d.columns))
            break
    except Exception as e:
        print(imp, "error", e)
