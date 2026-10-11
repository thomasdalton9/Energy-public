"""Probe 3: ARE Benin - paginate all documents, list categories, SBPE files, download+parse samples (manual)."""
import io, json, re, collections, requests
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36", "Accept": "application/json, */*"}
API = "https://backoffice.are.bj/api/documents"
items = []
p = 1
while True:
    try:
        j = requests.get(API, params={"page": p}, headers=H, timeout=40).json()
    except Exception as e:
        print("ERR page", p, e); break
    items += j["data"]
    if p >= j["meta"]["last_page"]: break
    p += 1
print("pages", p, "items", len(items))
cats = collections.Counter((i["category_slug"], i["category"]) for i in items)
for (s, n), c in cats.most_common(): print("CAT", c, s, "|", n)
for i in items:
    if "sbee" in (i["title"] + i["description"]).lower() or "ceb" in i["title"].lower() or "statist" in (i["title"]+i["category"]).lower() or "ceb" in i["category_slug"]:
        print("OTHER", i["created_at_raw"], i["category_slug"], "|", i["title"][:70], "|", i["file"]["mime"], i["file"]["size_formatted"])
sb = [i for i in items if "sbpe" in i["category_slug"]]
print("SBPE n", len(sb))
for i in sb:
    print("SB", i["created_at_raw"], i["title"][:60], "|", i["file"]["mime"], i["file"]["size"], i["file"]["url"].replace("https://backoffice.are.bj/uploads/documents/", ""))
json.dump(sb, open("sbpe_list.json", "w"))

import pdfplumber
n = len(sb)
idx = sorted({round(k * (n - 1) / 11) for k in range(12)}) if n else []
for k in idx:
    i = sb[k]; u = i["file"]["url"].replace("\\/", "/")
    try:
        r = requests.get(u, headers=H, timeout=90)
        b = r.content
        print("\nDL", k, i["created_at_raw"], i["title"][:50], r.status_code, len(b), b[:8])
        if b[:4] == b"%PDF":
            with pdfplumber.open(io.BytesIO(b)) as pdf:
                print("PAGES", len(pdf.pages))
                t = "\n".join((pg.extract_text() or "") for pg in pdf.pages)
                print("TEXTLEN", len(t)); print(t[:1800] if k in (idx[0], idx[len(idx)//2], idx[-1]) else t[:300])
                tb = pdf.pages[0].extract_tables(); print("TABLES p1", len(tb), [len(x) for x in tb])
        else:
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(b), data_only=True)
            for ws in wb: print("SHEET", ws.title, ws.dimensions)
    except Exception as e:
        print("DLERR", k, u, repr(e)[:200])
