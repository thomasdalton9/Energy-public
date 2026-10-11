"""Probe 4: ARE Benin - SBPE (3) and CEB hourly-load documents: list, download, parse (manual)."""
import io, json, requests, pdfplumber
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36", "Accept": "application/json, */*"}
API = "https://backoffice.are.bj/api/documents"
items = []; p = 1
while True:
    j = requests.get(API, params={"page": p}, headers=H, timeout=40).json()
    items += j["data"]
    if p >= j["meta"]["last_page"]: break
    p += 1
sel = [i for i in items if "sbpe" in i["category_slug"] or "ceb" in i["category_slug"] or "statistique" in (i["title"] + i["category"]).lower() or "rapport" in i["title"].lower()]
nofile = [i for i in items if not i["file"]]
print("RESULT items", len(items), "without file", len(nofile))
for i in sel:
    f = i["file"] or {}
    print("RESULT SEL", i["created_at_raw"], i["category_slug"][:30], "|", i["title"][:60], "|", i["slug"], "|", f.get("mime"), f.get("size"), (f.get("url") or "").replace("https://backoffice.are.bj/uploads/documents/", ""))
dates = sorted(i["created_at_raw"] for i in items)
print("RESULT date range all docs", dates[0], dates[-1])
for i in sel:
    f = i["file"]
    if not f or not ("sbpe" in i["category_slug"] or "ceb" in i["category_slug"]): continue
    try:
        r = requests.get(f["url"], headers=H, timeout=120); b = r.content
        print("RESULT DL", i["title"][:50], r.status_code, len(b), b[:6])
        if b[:4] == b"%PDF":
            with pdfplumber.open(io.BytesIO(b)) as pdf:
                print("RESULT PAGES", len(pdf.pages))
                t = "\n".join((pg.extract_text() or "") for pg in pdf.pages)
                print("RESULT TEXTLEN", len(t)); print("RESULT TEXT", t[:1500].replace("\n", " // "))
        elif b[:2] == b"PK":
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(b), data_only=True)
            for ws in wb:
                print("RESULT SHEET", ws.title, ws.dimensions)
                for row in list(ws.iter_rows(values_only=True))[:25]: print("RESULT ROW", row[:14])
    except Exception as e:
        print("RESULT DLERR", repr(e)[:200])
