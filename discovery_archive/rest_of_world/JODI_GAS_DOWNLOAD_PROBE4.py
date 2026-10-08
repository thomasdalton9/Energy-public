"""One-off probe 4: replicate the Beyond 20/20 (jodidb.org) CSV download of the JODI Gas World table."""
import html, re, time
import requests

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data/1.0"}
B = "http://www.jodidb.org"
s = requests.Session(); s.headers.update(H)
for i in range(4):
    try:
        r = s.get(B + "/TableViewer/tableView.aspx?ReportId=38673", timeout=90); break
    except requests.exceptions.ConnectionError as e:
        print("retry", i, type(e).__name__); time.sleep(5)
t = r.text
print("page", r.status_code, len(t), "cookies", s.cookies.get_dict())
js = s.get(B + "/TableViewer/tableView.js", timeout=60).text
i = js.find("onDownload"); print("=====onDownload=====\n", js[i:i + 3500], "\n=====")
u = s.get(B + "/Common/Util/util.js", timeout=60).text
for fn in ["function CreateHiddenFormField", "function IsSpawnWindow", "function SubmitForm", "function GoToPage", "function SetNextPage"]:
    j = u.find(fn); print(fn, j, "\n", u[j:j + 700] if j >= 0 else "")
api = s.get(B + "/wdsAPI.js", timeout=60).text
print("=====wdsAPI.js=====\n", api[:5000], "\n=====")
# form fields
fields = {k: html.unescape(v) for k, v in re.findall(r'<input type="hidden" name="([^"]+)" id="[^"]*" value="([^"]*)"', t, flags=re.S)}
print("fields:", len(fields), list(fields)[:12])
print("ReportView:", fields.get("sWD_ReportView", "")[:1500])
for target, extra in [("/TableViewer/tableView.aspx", {"WD_DownloadFormat": "CSV", "CS_NextPage": "/TableViewer/download.aspx"}),
                      ("/TableViewer/download.aspx", {"WD_DownloadFormat": "CSV"}),
                      ("/TableViewer/tableView.aspx", {"WD_DownloadFormat": "CSV"})]:
    data = dict(fields); data.update(extra)
    try:
        r = s.post(B + target, data=data, timeout=300, headers={"Referer": B + "/TableViewer/tableView.aspx?ReportId=38673"})
        ct = r.headers.get("content-type", ""); cd = r.headers.get("content-disposition", "")
        print(f"\nPOST {target} {extra}: {r.status_code} {ct} {cd} {len(r.content)}B")
        body = r.content.decode("utf-8", "replace")
        if "text/html" in ct:
            print("   title:", re.findall(r"<title>(.*?)</title>", body, flags=re.S)[:1], " | ", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", body))[:600])
        else:
            lines = body.splitlines(); print("   lines:", len(lines)); print("\n".join(l[:300] for l in lines[:12]))
    except Exception as e:  # noqa: BLE001
        print("POST failed", target, type(e).__name__, str(e)[:200])
