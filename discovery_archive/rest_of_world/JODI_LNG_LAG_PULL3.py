"""
One-off pull 3 (JKM -> LNG demand lag study): JODI Gas World table from jodidb.org via the Beyond 20/20 download
prompt (tableView -> downloadPrompt.aspx -> download.aspx). Saves the raw CSV to
discovery_archive/rest_of_world/data/jodi_gas_world_jodidb.csv.gz when it comes back; prints every intermediate page.
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


def req(method, url, **kw):
    for i in range(4):
        try:
            return s.request(method, url, timeout=kw.pop("timeout", 300), **kw)
        except requests.exceptions.ConnectionError as e:
            print("  retry", i + 1, url[:80], type(e).__name__, flush=True); time.sleep(5 * (i + 1))
    return s.request(method, url, timeout=300, **kw)


def inputs(t):
    out = {}
    for m in re.finditer(r"<(input|select|textarea)\b([^>]*)>", t, flags=re.I | re.S):
        attrs = dict(re.findall(r'(\w+)\s*=\s*"([^"]*)"', m.group(2)))
        if "name" in attrs:
            out[attrs["name"]] = html.unescape(attrs.get("value", ""))
    return out


def strip(t, n=2500):
    return re.sub(r"\s+", " ", re.sub(r"<script.*?</script>", " ", re.sub(r"<style.*?</style>", " ", t, flags=re.S), flags=re.S | re.I))[:n]


def text(t, n=2500):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", re.sub(r"<script.*?</script>", " ", t, flags=re.S)))[:n]


r = req("GET", f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}")
t = r.text; fields = inputs(t)
print("table page", r.status_code, len(t), "fields", len(fields), flush=True)

# 1. download prompt
data = dict(fields); data.update({"WD_DownloadFormat": "CSV", "sCS_SpawnWindow": "True", "WD_PageBeforeSummary": "/TableViewer/tableView.aspx"})
r = req("POST", f"{B}/TableViewer/downloadPrompt.aspx", data=data, headers={"Referer": f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}"})
pt = r.text
print("\n== downloadPrompt:", r.status_code, len(pt), re.findall(r"<title>(.*?)</title>", pt, flags=re.S)[:1], flush=True)
print("TEXT:", text(pt, 2000), flush=True)
print("FORMS:", re.findall(r"<form[^>]*>", pt, flags=re.I), flush=True)
pf = inputs(pt)
vis = {k: v for k, v in pf.items() if not k.startswith(("sWD_", "WD_Exp", "WD_Item", "oWD_", "sCS_", "CS_", "LG_", "PR_", "IF_"))}
print("PROMPT INPUTS (non-state):", vis, flush=True)
for m in re.findall(r"<(?:input|button|a)[^>]*(?:onclick|href)=\"[^\"]*\"[^>]*>", pt, flags=re.I):
    print("   ctl:", m[:300], flush=True)
for m in re.findall(r"function\s+\w*(?:Download|OK|Submit|Ok)\w*\s*\([^)]*\)\s*\{.*?\n\}", pt, flags=re.S):
    print("   JS:", re.sub(r"\s+", " ", m)[:900], flush=True)
for src in re.findall(r'<script[^>]+src="([^"]+)"', pt):
    if "downloadPrompt" in src or "download" in src.lower():
        u = src if src.startswith("http") else B + (src if src.startswith("/") else "/TableViewer/" + src)
        js = req("GET", u).text
        print("   JS file", u, len(js), flush=True)
        for m in re.findall(r"function\s+\w+\s*\([^)]*\)\s*\{.*?\n\}", js, flags=re.S):
            if "Download" in m or "download" in m:
                print("     ", re.sub(r"\s+", " ", m)[:1200], flush=True)

# 2. final download: prompt's fields + WD_Command=Download
saved = False
for extra in ({"WD_Command": "Download"}, {"WD_Command": "Download", "WD_DownloadFormat": "CSV"}):
    data = dict(pf); data.update(extra)
    r = req("POST", f"{B}/TableViewer/download.aspx", data=data, headers={"Referer": f"{B}/TableViewer/downloadPrompt.aspx"})
    ct = r.headers.get("content-type", ""); body = r.content.decode("utf-8-sig", "replace")
    print(f"\n== download.aspx {extra}: {r.status_code} {ct} {r.headers.get('content-disposition', '')} {len(r.content)}B", flush=True)
    if "html" in ct.lower() or body.lstrip()[:1] == "<":
        print("   HTML:", re.findall(r"<title>(.*?)</title>", body, flags=re.S)[:1], text(body, 800), flush=True)
    else:
        lines = body.splitlines(); print("   lines:", len(lines)); print("\n".join(l[:300] for l in lines[:12]), flush=True)
        if len(body) > 5000:
            with gzip.open(os.path.join(OUT, "jodi_gas_world_jodidb.csv.gz"), "wt") as f:
                f.write(body)
            print("   SAVED", flush=True); saved = True
            break

# 3. wdsAPI GetData, minimal body as the viewer builds it
if not saved:
    body = ("WA=1&WA_Command=GetData&sWD_ReportId=38673&sWD_Rows=0,2&sWD_Cols=4,3&sWD_Others=1&sWD_DimActiveItemPos=0,0,0,0,0"
            "&WA_FirstRow=0&WA_MaxRows=20&WA_FirstCol=0&WA_MaxCols=20")
    for hdr in ({"Content-Type": "application/x-www-form-urlencoded"}, {"Content-Type": "text/plain"}):
        r = req("POST", f"{B}/wdsAPI.aspx", data=body, headers=hdr)
        print("\n== wdsAPI GetData", hdr, r.status_code, len(r.content), r.text[:1500], flush=True)
    r = req("GET", f"{B}/wdsAPI.aspx?" + body)
    print("\n== wdsAPI GetData GET", r.status_code, len(r.content), r.text[:800], flush=True)
print("DONE", sorted(os.listdir(OUT)), flush=True)
