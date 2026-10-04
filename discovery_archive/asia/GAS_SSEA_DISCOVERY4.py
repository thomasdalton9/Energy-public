"""
South & Southeast Asia gas discovery, round 4 (GitHub Actions):
  ID  ESDM 'Highlight' page form (daily / monthly / yearly gas production by date?) and Buku Statistik Migas 2024
      monthly gas tables
  PK  PBS monthly trade releases (LNG import quantity), media search
  VN  every NSO monthly 'Bieu' workbook in the media library (history depth) and the natural-gas row of a few
"""
import io
import re
from urllib.parse import urljoin

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "en,id;q=0.8,vi;q=0.6"}
T = (15, 120)
S = requests.Session()
S.headers.update(H)


def out(s=""):
    print(s, flush=True)


def get(u, **kw):
    try:
        r = S.get(u, timeout=T, verify=False, **kw)
        out(f"\n### {u} {kw.get('params') or ''}\n  -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}b "
            f"final={r.url} LM={r.headers.get('last-modified')}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"\n### {u}\n  ERROR {type(e).__name__}: {str(e)[:200]}")
        return None


def plain(html):
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", html, flags=re.S | re.I)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))


def indonesia():
    r = get("https://www.esdm.go.id/id/highlight")
    if r is not None:
        t = r.text
        for m in re.finditer(r"<form.*?</form>", t, re.S | re.I):
            out("  FORM: " + re.sub(r"\s+", " ", m.group(0))[:3000])
        i = t.find("Tanggal Akses")
        out("  RAW: " + re.sub(r"\s+", " ", t[max(0, i - 1500):i + 6000]))
        for s in re.findall(r"<script[^>]*>(.*?)</script>", t, re.S | re.I):
            if re.search(r"highlight|ajax|filter|tanggal", s, re.I):
                out("  SCRIPT: " + re.sub(r"\s+", " ", s)[:2500])
    for params in ({"tanggal": "2025-06-30"}, {"date": "2025-06-30"}, {"tanggal_akses": "30-06-2025"}):
        r = get("https://www.esdm.go.id/id/highlight", params=params)
        if r is not None:
            ctx = plain(r.text)
            i = ctx.find("Produksi Gas")
            out("  -> " + ctx[max(0, i - 300):i + 400])
    r = get("https://migas.esdm.go.id/cms/uploads/Statistik%20Migas/Buku%20Statistik%202024.pdf")
    if r is not None and r.content[:4] == b"%PDF":
        import pdfplumber
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            out(f"  pdf pages {len(pdf.pages)}")
            hits = 0
            for k, p in enumerate(pdf.pages):
                tx = p.extract_text() or ""
                if k < 8:
                    out(f"  ---- page {k + 1} (toc) ----\n{tx[:2500]}")
                elif re.search(r"gas bumi|natural gas", tx, re.I) and re.search(r"Januari|January|Jan\b", tx):
                    out(f"  ---- page {k + 1} ----\n{tx[:3000]}")
                    hits += 1
                    if hits >= 8:
                        break


def pakistan():
    for q in ("import", "LNG", "trade", "External_Trade", "Monthly_Review"):
        r = get(f"https://www.pbs.gov.pk/wp-json/wp/v2/media?search={q}&per_page=30&orderby=date&order=desc&_fields=date,source_url")
        if r is not None and r.status_code == 200:
            out(f"  total={r.headers.get('X-WP-Total')}")
            for m in r.json():
                out(f"    {m['date']} {m['source_url']}")
    for u in ("https://www.pbs.gov.pk/trade-statistics/", "https://www.pbs.gov.pk/trade-statistics-2/",
              "https://www.pbs.gov.pk/trade-detail/", "https://www.pbs.gov.pk/trade-summary/"):
        r = get(u)
        if r is not None and r.status_code == 200:
            ls = sorted(set(re.findall(r'href="([^"]+\.(?:pdf|xlsx?))"', r.text)))
            out(f"  files ({len(ls)}): " + "\n    ".join(ls[:60]))


def vietnam():
    files = []
    for p in range(1, 60):
        r = S.get("https://www.nso.gov.vn/wp-json/wp/v2/media", params={"search": "Bieu", "per_page": 100, "page": p,
                  "_fields": "date,source_url"}, timeout=T, verify=False)
        if r.status_code != 200:
            break
        js = r.json()
        if not js:
            break
        for m in js:
            u = m["source_url"]
            if re.search(r"\.xlsx?$", u, re.I) and not re.search(r"CPI|ksms|Lao-dong|Du-bao|gia", u, re.I):
                files.append((m["date"], u))
    out(f"  {len(files)} Bieu workbooks:")
    for d, u in files:
        out(f"    {d} {u}")
    for d, u in files[:: max(1, len(files) // 8)][:9]:
        r = get(u)
        if r is None or r.status_code != 200:
            continue
        try:
            x = pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None)
        except Exception as e:  # noqa: BLE001
            out(f"  read fail {e}")
            continue
        out(f"  sheets: {list(x)[:40]}")
        for name, df in x.items():
            m = df.apply(lambda c: c.astype(str).str.contains("Khí đốt thiên nhiên|Khí thiên nhiên|Natural gas", case=False,
                                                              regex=True)).any(axis=1)
            if m.any():
                out(f"  --- {name}: header\n{df.iloc[2:7].to_string()[:1500]}\n  rows:\n{df[m].to_string()[:800]}")


if __name__ == "__main__":
    for f in (indonesia, pakistan, vietnam):
        out(f"\n\n{'#' * 30} {f.__name__} {'#' * 30}")
        try:
            f()
        except Exception as e:  # noqa: BLE001
            out(f"!! {f.__name__}: {type(e).__name__}: {e}")
