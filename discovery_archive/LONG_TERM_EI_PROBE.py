"""One-off probe: the Energy Institute downloads page answers GitHub only to a browser TLS fingerprint
(curl_cffi). Lists every data asset linked and the sheets of the all-data workbook (LNG / pipeline trade)."""
import io
import re
import pandas as pd
from curl_cffi import requests as cr
PAGE = "https://www.energyinst.org/statistical-review/resources-and-data-downloads"
r = cr.get(PAGE, impersonate="chrome", timeout=60)
print(r.status_code)
for l in sorted(set(re.findall(r'href="([^"]*__data/assets[^"]*)"', r.text))):
    print("ASSET", l)
x = cr.get("https://www.energyinst.org/__data/assets/excel_doc/0008/1656215/EI-Stats-Review-ALL-data.xlsx",
           impersonate="chrome", timeout=300)
print("ALL", x.status_code, len(x.content), x.headers.get("Last-Modified"), x.headers.get("ETag"))
xl = pd.ExcelFile(io.BytesIO(x.content))
print("SHEETS", xl.sheet_names)
for s in xl.sheet_names:
    if re.search(r"lng|pipeline|trade|gas.*(prod|cons)", s, re.I):
        print("=====", s)
        print(pd.read_excel(xl, sheet_name=s, header=None).iloc[:8, :14].to_string(max_colwidth=30))
