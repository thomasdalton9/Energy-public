"""One-off probe 6: Beyond 20/20 item-selection round trip on the JODI gas table: the ItemSelection control's JS
(what it writes into WD_SelItems_DimN), the item lists per dimension, and a trial selection posted back to
tableView.aspx to see how the server rewrites sWD_ReportView / sWD_SelectedItemsCount."""
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
REF = {"Referer": f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}"}
r = req("GET", f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}"); t = r.text; f = inputs(t)

def show_dim(dim):
    data = dict(f); data.update({"WD_Command": "ShowDim", "WD_DimIndex": str(dim), "sWD_ActiveMember": str(dim), "sWD_ActiveDim": str(dim)})
    return req("POST", f"{B}/TableViewer/dimView.aspx", data=data, headers=REF).text

d2 = show_dim(2)
# 1. ItemSelection control JS
for src in re.findall(r'<script[^>]+src="([^"]+)"', d2):
    if "ItemSel" in src or "itemSel" in src or "Selection" in src:
        u = src if src.startswith("http") else B + (src if src.startswith("/") else "/TableViewer/" + src)
        js = req("GET", u).text
        print(f"\n== JS {u} {len(js)}B")
        for m in re.finditer(r"(?:this\.|function\s+)(\w*(?:SaveSelection|Selected|SelItems|Initialize|GetState|SaveState|Serialize|Expand)\w*)\s*=?\s*function\s*\([^)]*\)\s*\{", js):
            start = m.start(); depth = 0; i = js.find("{", start)
            j = i
            while j < len(js):
                if js[j] == "{": depth += 1
                elif js[j] == "}":
                    depth -= 1
                    if depth == 0: break
                j += 1
            body = js[start: j + 1]
            if len(body) < 3500 and any(k in body for k in ("WD_", "SelItems", "ExpItems", "AddSel", "RemSel", "value")):
                print("  ", re.sub(r"\s+", " ", body)[:1800])
        for h in re.findall(r".{0,160}(?:WD_SelItems|WD_AddSelItems|WD_RemSelItems|WD_ExpItems|SelItemsVarName|AddSelItemsVarName|ExpItemsVarName).{0,200}", js)[:30]:
            print("   hit:", re.sub(r"\s+", " ", h)[:380])
print("\nscripts on dimView:", re.findall(r'<script[^>]+src="([^"]+)"', d2))
# 2. item lists: print the HTML around the first items
for dim, label in ((2, "BALANCE"), (1, "Product"), (3, "Unit"), (4, "TIME"), (0, "Country")):
    dd = d2 if dim == 2 else show_dim(dim)
    fi = inputs(dd)
    print(f"\n== dim {dim} {label}: DisplayedItemsCount={fi.get('WD_DisplayedItemsCount')} ItemStateId={fi.get('WD_ItemStateId')} SelItems={fi.get('WD_SelItems_Dim%d' % dim, '?')[:200]} ExpItems={fi.get('WD_ExpItems_Dim%d' % dim, '?')[:100]}")
    body = re.sub(r"<script.*?</script>", " ", dd, flags=re.S)
    i = body.find("ItemSelectionCtl")
    seg = body[i: i + 200000] if i >= 0 else body
    rows = re.findall(r"<(?:tr|div|li|option|span)[^>]*(?:Item|item)[^>]*>.*?</(?:tr|div|li|option|span)>", seg, flags=re.S)
    print("   item-ish elements:", len(rows))
    for rrow in rows[:8]:
        print("   ", re.sub(r"\s+", " ", rrow)[:400])
    labs = re.findall(r"(?:title|value|data-\w+)=\"([^\"]{2,80})\"", seg)
    print("   attribute values sample:", labs[:40])
    txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "|", seg))
    print("   text:", txt[:1500])
# 3. trial selection round trips for BALANCE (dim 2)
f2 = inputs(d2)
for fmt, val in [("positions csv", "0,1"), ("positions semicolon", "0;1"), ("xml items", "<Items><Item pos=\"0\"/><Item pos=\"1\"/></Items>"), ("range", "<Range first=\"0\" last=\"1\"/>")]:
    data = dict(f2); data.update({"WD_SelItems_Dim2": val, "CS_NextPage": "/TableViewer/tableView.aspx", "WD_Command": "", "sWD_ActiveMember": "2", "sWD_ActiveDim": "2"})
    r = req("POST", f"{B}/TableViewer/tableView.aspx", data=data, headers={"Referer": f"{B}/TableViewer/dimView.aspx"})
    fr = inputs(r.text)
    rv = fr.get("sWD_ReportView", "")
    m = re.search(r'<Dim name="BALANCE">.*?</Dim>', rv, flags=re.S)
    print(f"\n== trial {fmt}: {r.status_code} title={re.findall(r'<title>(.*?)</title>', r.text, flags=re.S)[:1]} SelectedItemsCount={fr.get('sWD_SelectedItemsCount')} Rows={fr.get('sWD_RowsItemsCount')} BALANCE={m.group(0)[:600] if m else rv[:300]}")
    if "Error" in r.text[:3000]:
        print("   ", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))[:400])
