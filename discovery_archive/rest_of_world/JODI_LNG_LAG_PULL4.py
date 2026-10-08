"""
One-off pull 4 (JKM -> LNG demand lag study): the current JODI Gas World database from the Beyond 20/20 table viewer
at jodidb.org (ReportId 38673; the jodidata.org static zip stops at 2018-08 and the REST API needs an account).

How: the viewer keeps its state in form fields; sWD_ReportView holds per-dimension selections as
  <Definition><Items type="handles"><String value="h1,h2,..."/></Items></Definition>
The download is limited to 15,000 cells, so the table (Country x BALANCE rows, TIME x Unit columns, Product as the
"other" dimension) is pulled in chunks: one product x one unit x one flow x all countries x half the months.
Each chunk is the viewer's CSV download, stored raw in discovery_archive/rest_of_world/data/jodi_gas_world_jodidb_raw.zip
(member p{product}_u{unit}_b{flow}_t{chunk}.csv) together with a discovery log (jodi_pull_log.txt). Parsing happens in
the private repo's analysis script.
"""
import html
import io
import os
import re
import time
import zipfile

import requests

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(OUT, exist_ok=True)
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data/1.0"}
B = "http://www.jodidb.org"
REPORT = "38673"
s = requests.Session(); s.headers.update(H)
LOG = open(os.path.join(OUT, "jodi_pull_log.txt"), "w")


def log(*a):
    msg = " ".join(str(x) for x in a)
    print(msg, flush=True); LOG.write(msg + "\n"); LOG.flush()


def req(method, url, **kw):
    for i in range(5):
        try:
            return s.request(method, url, timeout=kw.pop("timeout", 300), **kw)
        except requests.exceptions.ConnectionError as e:
            log("  retry", i + 1, url[:70], type(e).__name__); time.sleep(5 * (i + 1))
    return s.request(method, url, timeout=300, **kw)


def inputs(t):
    out = {}
    for m in re.finditer(r"<(input|select|textarea)\b((?:[^>\"]|\"[^\"]*\")*)>", t, flags=re.I | re.S):
        attrs = dict(re.findall(r'(\w+)\s*=\s*"([^"]*)"', m.group(2)))
        if "name" in attrs:
            out[attrs["name"]] = html.unescape(attrs.get("value", ""))
    return out


def text(t, n=600):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", re.sub(r"<script.*?</script>", " ", t, flags=re.S)))[:n]


REF_T = {"Referer": f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}"}
r = req("GET", f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}")
BASE = inputs(r.text)
RV0 = BASE["sWD_ReportView"]
log("table page", r.status_code, "fields", len(BASE), "counts", BASE.get("sWD_SelectedItemsCount"), "rows", BASE.get("sWD_Rows"), "cols", BASE.get("sWD_Cols"), "others", BASE.get("sWD_Others"))
DIMS = ["Country", "Product", "BALANCE", "Unit", "TIME"]


def build_rv(sel):
    """sel: dim name -> list of handles (None = all)."""
    rv = RV0
    for name, handles in sel.items():
        if handles is None:
            continue
        defn = '<Definition><Items type="handles"><String value="%s"/></Items></Definition>' % ",".join(str(h) for h in handles)
        rv, n = re.subn(r'(<Dim name="%s">.*?)<Definition>.*?</Definition>' % re.escape(name), lambda m: m.group(1) + defn, rv, count=1, flags=re.S)
        assert n == 1, name
    return rv


def download(sel, label=""):
    """Apply the selection, open the download prompt, fetch the CSV. Returns text or None."""
    data = dict(BASE); data["sWD_ReportView"] = build_rv(sel); data["WD_Command"] = ""
    r = req("POST", f"{B}/TableViewer/tableView.aspx", data=data, headers=REF_T)
    f = inputs(r.text)
    if "sWD_ReportView" not in f:
        log("  apply failed", label, text(r.text, 300)); return None
    data = dict(f); data.update({"WD_DownloadFormat": "CSV", "sCS_SpawnWindow": "True", "WD_PageBeforeSummary": "/TableViewer/tableView.aspx"})
    r = req("POST", f"{B}/TableViewer/downloadPrompt.aspx", data=data, headers=REF_T)
    pt = r.text
    urls = re.findall(r"OnDownload\(\s*['\"]([^'\"]+)['\"]\s*\)", pt)
    ivt = re.findall(r"OnDownloadIvt\(\s*['\"]([^'\"]*)['\"]\s*,\s*['\"]([^'\"]*)['\"]\s*,\s*['\"]([^'\"]*)['\"]\s*\)", pt)
    if not urls and not ivt:
        log("  prompt gave no link", label, "counts", f.get("sWD_SelectedItemsCount"), "|", text(pt, 300)); return None
    if urls:
        u = urls[0]; u = u if u.startswith("http") else B + (u if u.startswith("/") else "/TableViewer/" + u)
        r = req("GET", u, headers={"Referer": f"{B}/TableViewer/downloadPrompt.aspx"})
    else:
        mime, name, path = ivt[0]
        pf = inputs(pt); pf.update({"WD_Command": "Download", "WD_DownloadMimeType": mime, "WD_DownloadFilename": name, "WD_DownloadFilePath": path})
        r = req("POST", f"{B}/TableViewer/download.aspx", data=pf, headers={"Referer": f"{B}/TableViewer/downloadPrompt.aspx"})
    body = r.content.decode("utf-8-sig", "replace")
    if body.lstrip()[:1] == "<" or "text/html" in r.headers.get("content-type", ""):
        log("  download not CSV", label, r.status_code, r.headers.get("content-type"), text(body, 300)); return None
    return body


# ---------------------------------------------------------------- discovery
first = download({"Country": [0], "Product": [0], "BALANCE": None, "Unit": None, "TIME": [0, 1, 2]}, "discovery-1")
log("=== discovery 1 (country 0, product 0, all flows, all units, time 0-2):\n" + (first or "NONE"))
zbuf = io.BytesIO(); Z = zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED)
if first:
    Z.writestr("discovery_1.csv", first)
# single-handle tests to learn handle -> label and the handle range per dimension
labels = {}
for dim, tests, other in [("BALANCE", list(range(0, 18)), {"Country": [0], "Product": [0], "Unit": [0], "TIME": [0]}),
                          ("Unit", list(range(0, 5)), {"Country": [0], "Product": [0], "BALANCE": [0], "TIME": [0]}),
                          ("Product", list(range(0, 4)), {"Country": [0], "BALANCE": [0], "Unit": [0], "TIME": [0]}),
                          ("TIME", [0, 1, 2, 105, 106, 209, 210, 211, 212], {"Country": [0], "Product": [0], "BALANCE": [0], "Unit": [0]}),
                          ("Country", [0, 1, 2, 93, 94, 95, 96], {"Product": [0], "BALANCE": [0], "Unit": [0], "TIME": [0]})]:
    labels[dim] = {}
    for h in tests:
        sel = dict(other); sel[dim] = [h]
        c = download(sel, f"{dim}[{h}]")
        if c:
            lines = [l for l in c.splitlines() if l.strip()]
            labels[dim][h] = lines
            log(f"--- {dim} handle {h}: " + " || ".join(l[:160] for l in lines[:8]))
        else:
            log(f"--- {dim} handle {h}: no data")
        time.sleep(0.2)
# whole-dimension listings (one download each, other dims single)
for dim, other in [("TIME", {"Country": [0], "Product": [0], "BALANCE": [0], "Unit": [0]}), ("Country", {"Product": [0], "BALANCE": [0], "Unit": [0], "TIME": [0]})]:
    sel = dict(other); sel[dim] = None
    c = download(sel, f"{dim}-all")
    if c:
        Z.writestr(f"discovery_{dim}_all.csv", c)
        log(f"=== {dim} all (first 2500 chars):\n" + c[:2500])

# ---------------------------------------------------------------- main pull
nb = max(labels["BALANCE"].keys(), default=-1) + 1
nu = max(labels["Unit"].keys(), default=-1) + 1
npr = max(labels["Product"].keys(), default=-1) + 1
ntime = 211 if 210 in labels["TIME"] and 211 not in labels["TIME"] else (max(labels["TIME"].keys(), default=210) + 1)
log(f"=== handle counts: product {npr}, unit {nu}, balance {nb}, time {ntime}")
half = (ntime + 1) // 2
chunks = [list(range(0, half)), list(range(half, ntime))]
n_ok = n_fail = 0
for p in range(npr):
    for u in range(nu):
        for b in range(nb):
            for ci, tc in enumerate(chunks):
                sel = {"Country": None, "Product": [p], "BALANCE": [b], "Unit": [u], "TIME": tc}
                c = download(sel, f"p{p}u{u}b{b}t{ci}")
                if c:
                    Z.writestr(f"p{p}_u{u}_b{b}_t{ci}.csv", c); n_ok += 1
                else:
                    n_fail += 1
                    # retry in quarter chunks
                    q = len(tc) // 2
                    for cj, sub in enumerate((tc[:q], tc[q:])):
                        c2 = download({**sel, "TIME": sub}, f"p{p}u{u}b{b}t{ci}{cj}")
                        if c2:
                            Z.writestr(f"p{p}_u{u}_b{b}_t{ci}{cj}.csv", c2); n_ok += 1
                time.sleep(0.2)
            log(f"product {p} unit {u} flow {b}: ok {n_ok} fail {n_fail}")
Z.close()
with open(os.path.join(OUT, "jodi_gas_world_jodidb_raw.zip"), "wb") as fzip:
    fzip.write(zbuf.getvalue())
log("DONE", "members", n_ok, "failed", n_fail, "zip bytes", len(zbuf.getvalue()))
LOG.close()
