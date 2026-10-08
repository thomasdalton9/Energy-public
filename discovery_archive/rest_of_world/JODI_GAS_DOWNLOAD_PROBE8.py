"""One-off probe 8: (a) how the ItemSelection control calls getItems.aspx (item handles + labels per dimension) and
what it returns; (b) whether a Definition written straight into sWD_ReportView is honoured by tableView.aspx."""
import html, re, time
import requests

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data/1.0"}
B = "http://www.jodidb.org"; REPORT = "38673"
s = requests.Session(); s.headers.update(H)
def req(method, url, **kw):
    for i in range(4):
        try:
            return s.request(method, url, timeout=kw.pop("timeout", 120), **kw)
        except requests.exceptions.ConnectionError as e:
            print("  retry", i + 1, url[:80], type(e).__name__, flush=True); time.sleep(5 * (i + 1))
    return s.request(method, url, timeout=120, **kw)
def inputs(t):
    out = {}
    for m in re.finditer(r"<(input|select|textarea)\b((?:[^>\"]|\"[^\"]*\")*)>", t, flags=re.I | re.S):
        attrs = dict(re.findall(r'(\w+)\s*=\s*"([^"]*)"', m.group(2)))
        if "name" in attrs:
            out[attrs["name"]] = html.unescape(attrs.get("value", ""))
    return out
def func_at(js, idx):
    i = js.rfind("function", 0, idx); j = js.find("{", i); depth = 0; k = j
    while k < len(js):
        if js[k] == "{": depth += 1
        elif js[k] == "}":
            depth -= 1
            if depth == 0: break
        k += 1
    return re.sub(r"\s+", " ", js[i:k + 1])
js = req("GET", f"{B}/Common/ItemSelection/ItemSelection.js").text
seen = set()
for m in re.finditer(r"GetItemsPage|getItems\.aspx|M_strPostBackUrl \+|XMLHttp|\.open\(|ajax|send\(", js):
    f_ = func_at(js, m.start())
    if f_[:80] not in seen:
        seen.add(f_[:80]); print("\n== fn using", m.group(0), ":\n", f_[:2500])
# (a) call getItems.aspx with the dimView form state for dim 2 (BALANCE) and dim 4 (TIME)
REF = {"Referer": f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}"}
r = req("GET", f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}"); f = inputs(r.text)
for dim in ("2", "1", "3", "4", "0"):
    data = dict(f); data.update({"WD_Command": "ShowDim", "WD_DimIndex": dim, "sWD_ActiveMember": dim, "sWD_ActiveDim": dim})
    dv = req("POST", f"{B}/TableViewer/dimView.aspx", data=data, headers=REF).text; fd = inputs(dv)
    for extra in ({"WD_FirstItem": "0", "WD_ItemsCount": "250"}, {"WD_ItemFirstRow_Dim" + dim: "0", "WD_DisplayedItemsCount": "250"}, {}):
        d = dict(fd); d.update(extra)
        rr = req("POST", f"{B}/TableViewer/getItems.aspx", data=d, headers={"Referer": f"{B}/TableViewer/dimView.aspx"})
        body = rr.text
        print(f"\n== getItems dim {dim} {extra}: {rr.status_code} {rr.headers.get('content-type')} {len(body)}B")
        print("   ", body[:1800].replace("\n", " "))
        if len(body) > 200 and "Error" not in body[:300]:
            break
# (b) direct Definition in ReportView: BALANCE handles "0,1"; TIME except nothing
rv = f["sWD_ReportView"]
rv2 = re.sub(r'(<Dim name="BALANCE">.*?<Definition>)<All/>(</Definition>)', r'\1<Items type="handles"><String value="0,1"/></Items>\2', rv, flags=re.S)
data = dict(f); data["sWD_ReportView"] = rv2; data["WD_Command"] = ""
rr = req("POST", f"{B}/TableViewer/tableView.aspx", data=data, headers=REF); fr = inputs(rr.text)
print("\n== direct Definition (Items handles 0,1):", rr.status_code, re.findall(r"<title>(.*?)</title>", rr.text, flags=re.S)[:1], "SelectedItemsCount=", fr.get("sWD_SelectedItemsCount"), "Rows=", fr.get("sWD_RowsItemsCount"))
m = re.search(r'<Dim name="BALANCE">.*?</Dim>', fr.get("sWD_ReportView", ""), flags=re.S); print("   ", re.sub(r".*?<Definition>", "<Definition>", m.group(0))[:300] if m else fr.get("sWD_ReportView", "")[:200])
if "Error" in rr.text[:4000]: print("   ", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", rr.text))[:300])
# the table page itself: do row labels show the BALANCE items? print the first data-row labels
labels = re.findall(r"<t[dh][^>]*class=\"[^\"]*(?:RowLabel|Label|ItemLabel)[^\"]*\"[^>]*>(.*?)</t[dh]>", rr.text, flags=re.S | re.I)
print("   row/col labels on table page:", [re.sub(r"<[^>]+>", "", x).strip() for x in labels][:40])
