"""
South & Southeast Asia gas discovery, round 5 (GitHub Actions):
  ID  ESDM Highlight 'Time series' form (kategori=Produksi Gas, start, end) - daily gas production history?
      Buku Statistik Migas: gas production table page in older books (2019, Semester I 2021, 2022) + Table 1.7
  PK  PBS monthly import workbook (Import_August-2026.xlsx): LNG rows
"""
import io
import re

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "en,id;q=0.8"}
T = (15, 150)
S = requests.Session()
S.headers.update(H)


def out(s=""):
    print(s, flush=True)


def get(u, **kw):
    try:
        r = S.get(u, timeout=T, verify=False, **kw)
        out(f"\n### {u} {kw.get('params') or ''}\n  -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}b "
            f"final={r.url}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"\n### {u}\n  ERROR {type(e).__name__}: {str(e)[:200]}")
        return None


def indonesia():
    for params in ({"kategori": "Produksi Gas", "start": "01/01/2026", "end": "30/09/2026"},
                   {"kategori": "Produksi Gas", "start": "2026-01-01", "end": "2026-09-30"},
                   {"kategori": "Produksi Gas", "start": "01-01-2024", "end": "31-12-2024"}):
        r = get("https://www.esdm.go.id/id/highlight", params=params)
        if r is None:
            continue
        t = r.text
        i = t.find("Time series")
        seg = t[i:i + 12000]
        out("  SEG: " + re.sub(r"\s+", " ", re.sub(r"<option[^>]*>[^<]*</option>", "", seg))[:6000])
        for s in re.findall(r"<script[^>]*>(.*?)</script>", t, re.S | re.I):
            if re.search(r"Highcharts|Chart\(|series|labels|data\s*:", s):
                out("  SCRIPT: " + re.sub(r"\s+", " ", s)[:5000])
    for u in ("https://migas.esdm.go.id/cms/uploads/informasi-publik/Stat_tahunan/Statistik-Migas-2019.pdf",
              "https://migas.esdm.go.id/cms/uploads/informasi-publik/Stat_tahunan/Statistik-Migas-Semester-I-2021.pdf",
              "https://migas.esdm.go.id/cms/uploads/informasi-publik/Stat_tahunan/Statistik-Migas-2022.pdf",
              "https://migas.esdm.go.id/cms/uploads/Statistik%20Migas/78a193bc85f6b42a4810001d4f815b96.pdf",
              "https://migas.esdm.go.id/cms/uploads/Statistik%20Migas/Buku%20Statistik%202024.pdf"):
        r = get(u)
        if r is None or r.content[:4] != b"%PDF":
            continue
        import pdfplumber
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            out(f"  pages {len(pdf.pages)}; LM={r.headers.get('last-modified')}")
            for k, p in enumerate(pdf.pages):
                tx = p.extract_text() or ""
                if re.search(r"Produksi Gas Bumi|Natural Gas Production|Pemanfaatan Gas|Utilization of Domestic|"
                             r"Ekspor LNG|LNG Export|Neraca Gas", tx, re.I) and (re.search(r"\bJan", tx) or "TOTAL" in tx
                                                                                  or "Pemanfaatan" in tx):
                    out(f"  ---- page {k + 1} ----\n{tx[:1800]}\n  ...\n{tx[-900:]}")


def pakistan():
    r = get("https://www.pbs.gov.pk/wp-content/uploads/2020/07/Import_August-2026.xlsx")
    if r is not None and r.status_code == 200:
        x = pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None)
        out(f"  sheets: {list(x)}")
        for name, df in x.items():
            out(f"  --- {name} {df.shape}\n{df.head(12).to_string()[:2500]}")
            m = df.apply(lambda c: c.astype(str).str.contains("LNG|Liquefied|Natural gas|PETROLEUM", case=False)).any(axis=1)
            out(df[m].to_string()[:3000])


if __name__ == "__main__":
    for f in (indonesia, pakistan):
        out(f"\n\n{'#' * 30} {f.__name__} {'#' * 30}")
        try:
            f()
        except Exception as e:  # noqa: BLE001
            out(f"!! {f.__name__}: {type(e).__name__}: {e}")
