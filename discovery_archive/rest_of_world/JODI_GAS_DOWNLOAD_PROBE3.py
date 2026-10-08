"""One-off probe 3: the gas data-downloads page renders <file-list type="gas"> by JavaScript. Find the endpoint it
calls, try candidate file URLs, and dump the jodidb.org (Beyond 20/20) download mechanism."""
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data/1.0"}
import time
_get = requests.get
def get_retry(url, **kw):
    for i in range(4):
        try:
            return _get(url, **kw)
        except requests.exceptions.ConnectionError as e:
            print("  retry", i + 1, url[:80], type(e).__name__); time.sleep(5 * (i + 1))
    return _get(url, **kw)
requests.get = get_retry
B = "https://www.jodidata.org"
t = requests.get(B + "/gas/database/data-downloads.aspx", headers=H, timeout=60).text
scripts = re.findall(r"<script[^>]*>.*?</script>", t, flags=re.S | re.I)
print("scripts:", len(scripts))
for s in scripts:
    print("  ", s[:400].replace("\n", " "))
for src in re.findall(r'<script[^>]+src="([^"]+)"', t):
    u = src if src.startswith("http") else B + src
    try:
        js = requests.get(u, headers=H, timeout=60).text
        hits = [m for m in re.findall(r".{0,120}(?:file-list|fileList|FileList|/api/|files\?|downloads/gas|\.csv|\.zip).{0,160}", js)]
        print(f"\n== JS {u} {len(js)}B hits {len(hits)}")
        for h in hits[:40]:
            print("   ", h.replace("\n", " ")[:300])
    except Exception as e:  # noqa: BLE001
        print("JS failed", u, type(e).__name__)
print()
cands = ["/api/files?type=gas", "/api/filelist?type=gas", "/api/file-list?type=gas", "/api/downloads?type=gas", "/api/downloads/gas",
         "/api/gas/files", "/api/files/gas", "/_resources/files/downloads/gas-data/annual-csv/2024.csv",
         "/_resources/files/downloads/gas-data/annual-csv/primary/2024.csv", "/_resources/files/downloads/gas-data/2024.csv",
         "/_resources/files/downloads/gas-data/csv/2024.csv", "/_resources/files/downloads/gas-data/gasyear2026.csv",
         "/_resources/files/downloads/gas-data/annual-csv/gasyear2026.csv", "/jodi-publisher/gas/27/", "/jodi-publisher/gas/",
         "/jodi-publisher/gas/27/jodi-data-availability-by-country.csv", "/_resources/files/downloads/gas-data/", "/api/", "/api/gasdata?dateFrom=2026-01&dateTo=2026-08&flows=IMPLNG&pageSize=100"]
for c in cands:
    try:
        r = requests.get(B + c, headers=H, timeout=60)
        print("GET", c, r.status_code, r.headers.get("content-type", "")[:40], len(r.content), r.text[:200].replace("\n", " "))
    except Exception as e:  # noqa: BLE001
        print("GET", c, "failed", type(e).__name__)
# jodidb
try:
    r = requests.get("http://www.jodidb.org/TableViewer/tableView.aspx?ReportId=38673", headers=H, timeout=60)
    t2 = r.text
    print("\n== jodidb forms/hidden inputs")
    for m in re.findall(r"<form[^>]*>", t2): print("  ", m[:300])
    for m in re.findall(r'<input[^>]*type="hidden"[^>]*>', t2): print("  ", m[:200])
    i = t2.find("onDownload"); print("  onDownload ctx:", re.sub(r"\s+", " ", t2[max(0, i-200): i+1500]))
    for src in re.findall(r'<script[^>]+src="([^"]+)"', t2):
        u = src if src.startswith("http") else "http://www.jodidb.org" + (src if src.startswith("/") else "/TableViewer/" + src)
        try:
            js = requests.get(u, headers=H, timeout=60).text
            j = js.find("onDownload"); print(f"  JS {u} {len(js)}B onDownload at {j}")
            if j >= 0: print("   ", re.sub(r"\s+", " ", js[j: j+1500]))
        except Exception as e:  # noqa: BLE001
            print("  JS failed", u, type(e).__name__)
    for u in ["http://www.jodidb.org/TableViewer/download.aspx?ReportId=38673", "http://www.jodidb.org/TableViewer/downloadCsv.aspx?ReportId=38673",
              "http://www.jodidb.org/TableViewer/tableView.aspx?ReportId=38673&IF_Format=csv"]:
        try:
            r = requests.get(u, headers=H, timeout=60); print("  GET", u, r.status_code, r.headers.get("content-type"), len(r.content), r.text[:150].replace("\n", " "))
        except Exception as e:  # noqa: BLE001
            print("  GET", u, "failed", type(e).__name__)
except Exception as e:  # noqa: BLE001
    print("jodidb failed", type(e).__name__, str(e)[:200])
