"""One-off probe 5: how the Beyond 20/20 viewer expresses item selections in sWD_ReportView (to get the JODI gas table
under the 15,000-cell download limit): search the viewer's JS for the Definition XML schema, dump the full ReportView
and the dimension-selection page."""
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
r = req("GET", f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}"); t = r.text; f = inputs(t)
print("=====ReportView=====\n", f.get("sWD_ReportView"), "\n=====")
print({k: v for k, v in f.items() if k.startswith(("sWD_", "oWD_", "WD_", "IF_"))})
# JS search
seen = set()
for page in (t,):
    for src in re.findall(r'<script[^>]+src="([^"]+)"', page):
        if "google" in src: continue
        u = src if src.startswith("http") else B + (src if src.startswith("/") else "/TableViewer/" + src)
        if u in seen: continue
        seen.add(u)
        js = req("GET", u).text
        hits = re.findall(r".{0,200}(?:<Definition|</Definition|<All/>|<Items|<Item |<Range|<Codes|<Code |Definition>|ItemSelection|WD_Command\", \"\w+\").{0,200}", js)
        if hits:
            print(f"\n== {u} {len(js)}B")
            for h in hits[:60]:
                print("   ", re.sub(r"\s+", " ", h)[:420])
# dimension view page (dim 2 = BALANCE)
for dim in ("2", "4"):
    data = dict(f); data.update({"WD_Command": "ShowDim", "WD_DimIndex": dim, "sWD_ActiveMember": dim})
    r = req("POST", f"{B}/TableViewer/dimView.aspx", data=data, headers={"Referer": f"{B}/TableViewer/tableView.aspx?ReportId={REPORT}"})
    print(f"\n== dimView dim {dim}: {r.status_code} {len(r.text)} {re.findall(r'<title>(.*?)</title>', r.text, flags=re.S)[:1]}")
    print("   text:", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", re.sub(r"<script.*?</script>", " ", r.text, flags=re.S)))[:1500])
    fi = inputs(r.text)
    print("   new/changed inputs:", {k: v[:200] for k, v in fi.items() if f.get(k) != v})
    for m in re.findall(r"<(?:a|input|button)[^>]*(?:onclick|href)=\"javascript:[^\"]*\"[^>]*>", r.text, flags=re.I)[:40]:
        print("   ctl:", m[:220])
    for src in re.findall(r'<script[^>]+src="([^"]+)"', r.text):
        if "dimView" in src:
            u = src if src.startswith("http") else B + (src if src.startswith("/") else "/TableViewer/" + src)
            js = req("GET", u).text
            print(f"   JS {u} {len(js)}B")
            for m in re.findall(r"function\s+\w+\s*\([^)]*\)\s*\{.*?\n\}", js, flags=re.S):
                if any(k in m for k in ("WD_Command", "Select", "Definition", "Items")):
                    print("     ", re.sub(r"\s+", " ", m)[:700])
