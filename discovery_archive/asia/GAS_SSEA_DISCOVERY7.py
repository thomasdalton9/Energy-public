"""
South & Southeast Asia gas, round 7 (GitHub Actions): layout of Ditjen Migas Table 1.6 (monthly gas production by KKKS)
in the 2020-2023 books, which the first Indonesia pull test failed to parse; and an old NSO .xls that xlrd could not
decode (encoding_override test).
"""
import io
import re

import pandas as pd
import pdfplumber
import requests
import urllib3
import xlrd

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
B = "https://migas.esdm.go.id/cms/uploads/informasi-publik/Stat_tahunan/"
BOOKS = [B + f for f in ("Statistik-Migas-Semester-I-2020-N.pdf", "Statistik-Migas-Semester-I-2021.pdf",
                         "Statistik-Migas-2021.pdf", "Statistik-Migas-Semester-I-2022.pdf", "Statistik-Migas-2022.pdf",
                         "Statistik-Migas-Semester-I-2023.pdf", "Statistik-Migas-2023.pdf")] + \
        ["https://migas.esdm.go.id/cms/uploads/Statistik%20Migas/78a193bc85f6b42a4810001d4f815b96.pdf"]


def out(*a):
    print(*a, flush=True)


for u in BOOKS:
    out(f"\n######## {u.rsplit('/', 1)[-1]}")
    try:
        r = requests.get(u, headers=H, timeout=(20, 240), verify=False)
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            for k, p in enumerate(pdf.pages):
                tx = p.extract_text() or ""
                if re.search(r"Produksi Gas Bumi", tx, re.I) and re.search(r"MMSCFD", tx) and k > 10:
                    out(f"---- page {k + 1} ----\n{tx[:2200]}\n ... \n{tx[-700:]}")
    except Exception as e:  # noqa: BLE001
        out(f"  {type(e).__name__}: {e}")

for u in ("https://www.nso.gov.vn/wp-content/uploads/2019/04/Bieu-01-2016.xls",
          "https://www.nso.gov.vn/wp-content/uploads/2020/11/Bieu1206.xls"):
    out(f"\n######## {u}")
    c = requests.get(u, headers=H, timeout=(20, 120), verify=False).content
    for enc in ("cp1252", "cp1258", "utf-8", "latin-1"):
        try:
            book = xlrd.open_workbook(file_contents=c, encoding_override=enc)
            x = pd.read_excel(book, sheet_name=None, header=None, engine="xlrd")
            out(f"  {enc}: sheets {list(x)[:12]}")
            for name, df in x.items():
                m = df.apply(lambda col: col.astype(str).str.contains("thi[eê]n nhi[eê]n|Kh. ..t", case=False)).any(axis=1)
                if m.any():
                    out(f"   {name}:\n{df.iloc[max(0, m.idxmax() - 8):m.idxmax() + 1].to_string()[:1500]}")
                    break
            break
        except Exception as e:  # noqa: BLE001
            out(f"  {enc}: {type(e).__name__}: {str(e)[:150]}")
