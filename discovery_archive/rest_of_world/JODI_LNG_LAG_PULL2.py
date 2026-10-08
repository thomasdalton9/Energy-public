"""
One-off pull 2 for the JKM -> LNG-demand lag study: the current JODI Gas World database from the Beyond 20/20 table
viewer at jodidb.org (the static csv_beta zip on jodidata.org stops at 2018-08; the REST API needs an account).
Writes discovery_archive/rest_of_world/data/jodi_gas_world_jodidb.csv.gz (raw CSV as downloaded) if the download works,
and prints diagnostics of the fallbacks (wdsAPI GetData JSON, extract viewer) otherwise.
"""
import gzip
import html
import os
import re
import time

import requests

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(OUT, exist_ok=True)
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data/1.0"}
B = "http://www.jodidb.org"
REPORT = "38673"
s = requests.Session(); s.headers.update(H)


def get(url, **kw):
    for i in range(4):
        try:
            return s.get(url, timeout=kw.pop("timeout", 120), **kw)
        except requests.exceptions.ConnectionError as e:
            print("  retry", i + 1, url[:80], type(e).__name__, flush=True); time.sleep(5 * (i + 1))
    return s.get(url, timeout=120, **kw)


def post(url, data, timeout=600):
    for i in range(3):
        try:
            return s.post(url, data=data, timeout=timeout, headers={"Referer": f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}"})
        except requests.exceptions.ConnectionError as e:
            print("  retry", i + 1, url[:80], type(e).__name__, flush=True); time.sleep(5 * (i + 1))
    return s.post(url, data=data, timeout=timeout)


def describe(r, label):
    ct = r.headers.get("content-type", ""); cd = r.headers.get("content-disposition", "")
    body = r.content.decode("utf-8-sig", "replace")
    print(f"\n{label}: {r.status_code} {ct} {cd} {len(r.content)}B", flush=True)
    if "text/html" in ct or body.lstrip().lower().startswith("<!doctype") or body.lstrip().lower().startswith("<html"):
        print("   HTML:", re.findall(r"<title>(.*?)</title>", body, flags=re.S)[:1], re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", body))[:700], flush=True)
        return None
    lines = body.splitlines(); print("   lines:", len(lines)); print("\n".join(l[:400] for l in lines[:15]), flush=True)
    return body


r = get(f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}")
t = r.text
print("table page", r.status_code, len(t), "cookies", s.cookies.get_dict(), flush=True)
fields = {k: html.unescape(v) for k, v in re.findall(r'<input type="hidden" name="([^"]+)" id="[^"]*" value="([^"]*)"', t, flags=re.S)}
print("fields", len(fields), {k: fields[k] for k in ("sWD_Rows", "sWD_Cols", "sWD_Others", "sWD_SelectedItemsCount", "sWD_RowsItemsCount", "sWD_ColsItemsCount", "sCS_DownloadLimit") if k in fields}, flush=True)
js = get(f"{B}/TableViewer/tableView.js").text
i = js.find("function onDownload"); j = js.find("CreateHiddenFormField", i)
print("onDownload tail:", re.sub(r"\s+", " ", js[j: j + 1200]), flush=True)

saved = False
for fmt in ("CSV", "SSV"):
    data = dict(fields); data.update({"WD_Command": "Download", "WD_DownloadFormat": fmt})
    r = post(f"{B}/TableViewer/download.aspx", data)
    body = describe(r, f"download.aspx WD_Command=Download fmt={fmt}")
    if body and len(body) > 5000:
        with gzip.open(os.path.join(OUT, "jodi_gas_world_jodidb.csv.gz"), "wt") as f:
            f.write(body)
        print("   SAVED jodi_gas_world_jodidb.csv.gz", flush=True); saved = True
        break
    # the form may expect the action on tableView.aspx with CS_NextPage pointing at download.aspx
    data2 = dict(fields); data2.update({"WD_Command": "Download", "WD_DownloadFormat": fmt, "CS_NextPage": "/TableViewer/download.aspx"})
    r = post(f"{B}/TableViewer/tableView.aspx", data2)
    body = describe(r, f"tableView.aspx WD_Command=Download fmt={fmt}")
    if body and len(body) > 5000:
        with gzip.open(os.path.join(OUT, "jodi_gas_world_jodidb.csv.gz"), "wt") as f:
            f.write(body)
        print("   SAVED jodi_gas_world_jodidb.csv.gz", flush=True); saved = True
        break

if not saved:
    # fallback 1: wdsAPI GetData (JSON) - what the viewer uses to scroll the table
    data = {k: v for k, v in fields.items()}
    data.update({"WA": "1", "WA_Command": "GetData", "WA_FirstRow": "0", "WA_MaxRows": "40", "WA_FirstCol": "0", "WA_MaxCols": "40",
                 "sWD_DimActiveItemPos": fields.get("oWD_DimActiveItemId", "0,0,0,0,0")})
    r = post(f"{B}/wdsAPI.aspx", data)
    print("\nwdsAPI GetData:", r.status_code, r.headers.get("content-type"), len(r.content), flush=True)
    print(r.text[:3000], flush=True)
    r = post(f"{B}/wdsAPI.aspx?" + "&".join(f"{k}={requests.utils.quote(str(v))}" for k, v in data.items() if k.startswith("WA") or k.startswith("sWD_R") or k.startswith("sWD_C") or k.startswith("sWD_O") or k.startswith("sWD_Dim")), {})
    print("\nwdsAPI GetData (query string):", r.status_code, len(r.content), r.text[:1500], flush=True)
    # fallback 2: extract viewer
    r = get(f"{B}/ExtractViewer/extractView.aspx?ReportId={REPORT}")
    print("\nextractView:", r.status_code, len(r.text), re.findall(r"<title>(.*?)</title>", r.text, flags=re.S)[:1], flush=True)
    for m in re.findall(r'(?:href|action)="([^"]+)"', r.text):
        if any(k in m.lower() for k in ["extract", "download", "csv"]):
            print("   ", m[:160])
    print(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))[:1500])
print("DONE", sorted(os.listdir(OUT)), flush=True)
