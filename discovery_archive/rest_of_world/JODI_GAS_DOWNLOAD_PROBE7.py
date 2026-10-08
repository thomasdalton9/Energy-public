"""One-off probe 7: format of WD_AddSelItems_DimN / WD_RemSelItems_DimN in the Beyond 20/20 ItemSelection control
(SaveSelection source), where the item list comes from, and trial selections on the JODI gas table."""
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
def func(js, name):
    i = js.find(name)
    if i < 0: return ""
    j = js.find("{", i); depth = 0; k = j
    while k < len(js):
        if js[k] == "{": depth += 1
        elif js[k] == "}":
            depth -= 1
            if depth == 0: break
        k += 1
    return re.sub(r"\s+", " ", js[i:k + 1])
js = req("GET", f"{B}/Common/ItemSelection/ItemSelection.js").text
for name in ("SaveSelection = function", "function SaveSelection", ".SaveSelection", "strSelected =", "GetSelectedItemsString", "SelectAll = function", "function SelectAll", "LoadItems", "GetItems", "wdsRequest", "ExecGetItems"):
    f_ = func(js, name)
    if f_:
        print(f"\n== {name}:\n{f_[:2500]}")
print("\n== aspx/urls in ItemSelection.js:", sorted(set(re.findall(r"[\w/]+\.aspx[^\"']{0,80}", js)))[:20])
for h in re.findall(r".{0,250}strSelected.{0,250}", js)[:12]:
    print("   sel:", re.sub(r"\s+", " ", h))
for h in re.findall(r".{0,200}WA_Command.{0,200}", js)[:10]:
    print("   WA:", re.sub(r"\s+", " ", h))
# wdsAPI: how does the control get items? look at wdsAPI.js for GetItems
api = req("GET", f"{B}/wdsAPI.js").text
for name in ("ExecGetItems", "ExecGetDimItems", "GetItems", "ExecGetLabels", "p_LoadData"):
    f_ = func(api, name)
    if f_: print(f"\n== wdsAPI {name}:\n{f_[:1500]}")
print("\n== wdsAPI Exec functions:", re.findall(r"this\.(Exec\w+)\s*=", api))
# trials on BALANCE (dim 2): keep only item 0
REF = {"Referer": f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}"}
r = req("GET", f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}"); f = inputs(r.text)
data = dict(f); data.update({"WD_Command": "ShowDim", "WD_DimIndex": "2", "sWD_ActiveMember": "2", "sWD_ActiveDim": "2"})
d2 = req("POST", f"{B}/TableViewer/dimView.aspx", data=data, headers=REF).text; f2 = inputs(d2)
rem_csv = ",".join(str(i) for i in range(1, 14))
for lab, add, rem in [("csv", "0", rem_csv), ("semicolon", "0", ";".join(str(i) for i in range(1, 14))), ("ranges", "0", "1-13"),
                      ("xml", '<Items><Item pos="0"/></Items>', '<Items>' + "".join(f'<Item pos="{i}"/>' for i in range(1, 14)) + '</Items>'),
                      ("clear+add", "0", "<All/>")]:
    for target in ("tableView", "dimView"):
        data = dict(f2); data.update({"WD_AddSelItems_Dim2": add, "WD_RemSelItems_Dim2": rem, "CS_NextPage": f"/TableViewer/{target}.aspx", "WD_Command": "", "sWD_ActiveMember": "2", "sWD_ActiveDim": "2"})
        r = req("POST", f"{B}/TableViewer/{target}.aspx", data=data, headers={"Referer": f"{B}/TableViewer/dimView.aspx"})
        fr = inputs(r.text); rv = fr.get("sWD_ReportView", "")
        m = re.search(r'<Dim name="BALANCE">.*?</Dim>', rv, flags=re.S)
        title = re.findall(r"<title>(.*?)</title>", r.text, flags=re.S)[:1]
        print(f"\n== trial {lab} -> {target}: {r.status_code} {title} SelectedItemsCount={fr.get('sWD_SelectedItemsCount')} CurrentDimSelectedCount={fr.get('sWD_CurrentDimSelectedCount')} Rows={fr.get('sWD_RowsItemsCount')}")
        if m: print("   BALANCE:", re.sub(r".*?<Groups", "<Groups", m.group(0))[:500])
        if "Error" in r.text[:4000]: print("   ", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))[:300])
