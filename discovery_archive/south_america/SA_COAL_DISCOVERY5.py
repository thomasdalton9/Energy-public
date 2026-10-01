"""
South America coal - discovery round 5:
  - Chile: Cochilco yearbook database xlsx (base-de-datos-anuario-de-estadisticas-cochilco-2000-2025.xlsx):
    coal (carbon) production rows and units
  - Peru: MINEM Anuario Minero 2021-2025 publications on gob.pe: attachments and coal (carbon) lines
"""
import io
import re
import sys

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "es,en;q=0.8"}
S = requests.Session()
S.headers.update(H)
pd.set_option("display.width", 250)


def get(url, **kw):
    kw.setdefault("timeout", (15, 180))
    try:
        r = S.get(url, **kw)
        print(f"  GET {r.url[:220]} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)} bytes")
        return r
    except Exception as e:
        print(f"  GET {url[:220]} -> {type(e).__name__}: {str(e)[:300]}")
        return None


def section(t):
    print("\n" + "=" * 100 + f"\n{t}\n" + "=" * 100, flush=True)


def chile():
    section("Chile Cochilco")
    r = get("https://www.cochilco.cl/web/download/1036/2026/16850/base-de-datos-anuario-de-estadisticas-cochilco-2000-2025.xlsx")
    if r is None or r.status_code != 200:
        return
    xl = pd.ExcelFile(io.BytesIO(r.content))
    print(f"     sheets: {xl.sheet_names}")
    for s in xl.sheet_names:
        df = xl.parse(s, header=None)
        hits = [i for i, row in df.iterrows() if re.search(r"carb[oó]n|coal", " ".join(map(str, row.values)), re.I)]
        if hits:
            print(f"\n     --- {s} {df.shape}")
            print(df.head(8).to_string(max_colwidth=25)[:2500])
            for i in hits[:20]:
                print(f"     r{i}: {[v for v in df.iloc[i].values if str(v) != 'nan'][:40]}")


def pdf_grep(content, pat, maxhits=60, maxpages=400):
    import pdfplumber
    n = 0
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        print(f"     pages: {len(pdf.pages)}")
        for i, p in enumerate(pdf.pages[:maxpages]):
            t = p.extract_text() or ""
            for ln in t.splitlines():
                if re.search(pat, ln, re.I):
                    print(f"     p{i + 1}: {ln[:250]}")
                    n += 1
                    if n >= maxhits:
                        return


def peru():
    section("Peru MINEM Anuario Minero")
    for path in ["/institucion/minem/informes-publicaciones/8434958-anuario-minero-2025",
                 "/institucion/minem/informes-publicaciones/6827926-anuario-minero-2024",
                 "/institucion/minem/informes-publicaciones/5804716-anuario-minero-2023"]:
        r = get("https://www.gob.pe" + path)
        if r is None or r.status_code != 200:
            continue
        files = sorted(set(re.findall(r'https://cdn\.www\.gob\.pe/uploads/document/file/\d+/[^"?#\s]+', r.text)))
        print(f"     {path}: {files[:20]}")
        for f in files:
            if re.search(r"\.xlsx?$", f, re.I):
                x = get(f)
                if x is None or x.status_code != 200:
                    continue
                xl = pd.ExcelFile(io.BytesIO(x.content))
                print(f"     sheets {xl.sheet_names[:40]}")
                for s in xl.sheet_names:
                    df = xl.parse(s, header=None)
                    for i, row in df.iterrows():
                        t = " | ".join(str(v) for v in row.values if str(v) != "nan")
                        if re.search(r"carb[oó]n|antracita|bitumin", t, re.I):
                            print(f"     [{s} r{i}] {t[:300]}")
            elif f.lower().endswith(".pdf"):
                x = get(f)
                if x is not None and x.status_code == 200 and x.content[:4] == b"%PDF":
                    pdf_grep(x.content, r"carb[oó]n|antracita|bitumin")
        break


if __name__ == "__main__":
    for w in sys.argv[1:] or ["chile", "peru"]:
        try:
            globals()[w]()
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"{w} failed: {type(e).__name__}: {e}")
