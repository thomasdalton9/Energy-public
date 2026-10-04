"""One-off probe: list the Energy Institute downloads page's data links and the sheets of the all-data xlsx."""
import io
import re
import pandas as pd
import requests
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"}
PAGE = "https://www.energyinst.org/statistical-review/resources-and-data-downloads"
r = requests.get(PAGE, headers=UA, timeout=60)
print(r.status_code, len(r.text))
links = sorted(set(re.findall(r'href="([^"]+\.(?:csv|xlsx|xls)[^"]*)"', r.text, flags=re.I)))
for l in links:
    print("LINK", l)
for l in links:
    if re.search(r"all.data", l, re.I):
        x = requests.get(requests.compat.urljoin(PAGE, l), headers=UA, timeout=300)
        xl = pd.ExcelFile(io.BytesIO(x.content))
        print("SHEETS", xl.sheet_names)
        for s in xl.sheet_names:
            if re.search(r"lng|pipeline|trade", s, re.I):
                print("=====", s)
                print(pd.read_excel(xl, sheet_name=s, header=None).iloc[:12, :12].to_string())
        break
