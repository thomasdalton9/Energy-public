"""
TRINIDAD_GAS.py reads only the newest MEEI Consolidated Monthly Bulletin,
which covers the current year to date (Jan-May 2026). To backfill from
2021, find the older bulletins - ideally the full-year (Jan-Dec) edition
for each of 2021-2025. Lists every bulletin-looking link (xlsx/xls/pdf)
on the listing page and on any linked archive/"previous bulletins" pages,
then opens the newest xlsx found for each year and prints which months
its gas sheet ("3A,3B") covers.
"""
import io
import re
from urllib.parse import urljoin

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 120)
START = "https://www.energy.gov.tt/?p=11949"


def out(*a):
    print(*a, flush=True)


def page_links(url):
    try:
        r = requests.get(url, headers=H, timeout=T)
    except Exception as e:
        out(f"  ERR {url}: {e}")
        return [], []
    files = sorted(set(urljoin(url, h) for h in re.findall(r'href="([^"]+\.(?:xlsx?|pdf))"', r.text, re.I)))
    pages = sorted(set(urljoin(url, h) for h, t in re.findall(r'href="([^"]+)"[^>]*>([^<]{0,120})<', r.text)
                       if re.search(r"bulletin|archive|previous|historic|statistic|20(1|2)\d", t + h, re.I)
                       and "energy.gov.tt" in urljoin(url, h) and not re.search(r"\.(xlsx?|pdf|jpg|png)$", h, re.I)))
    out(f"GET {url} -> {r.status_code}: {len(files)} files, {len(pages)} candidate pages")
    return files, pages


files, pages = page_links(START)
seen = {START}
for p in pages[:40]:
    if p in seen:
        continue
    seen.add(p)
    f2, _ = page_links(p)
    files += f2
files = sorted(set(files))
bull = [f for f in files if re.search(r"bulletin", f, re.I)]
out(f"\n{len(bull)} bulletin files:")
for f in bull:
    out("  ", f)


def year_key(u):
    m = re.search(r"(20\d\d)", u.rsplit("/", 1)[-1]) or re.search(r"/(20\d\d)/", u)
    return int(m.group(1)) if m else 0


xls = [f for f in bull if re.search(r"\.xlsx?$", f, re.I)]
for f in sorted(xls, key=lambda u: (year_key(u), u)):
    try:
        c = requests.get(f, headers=H, timeout=T).content
        xl = pd.ExcelFile(io.BytesIO(c))
        sheet = next((s for s in xl.sheet_names if s.strip().replace(" ", "") in ("3A,3B", "3A3B")), None)
        if sheet is None:
            out(f"\n{f.rsplit('/', 1)[-1]}: no 3A,3B sheet; sheets {xl.sheet_names[:12]}")
            continue
        df = xl.parse(sheet, header=None)
        months = sorted({pd.Timestamp(v).strftime("%Y-%m") for row in df.head(8).itertuples(index=False) for v in row
                         if hasattr(v, "year")})
        out(f"\n{f.rsplit('/', 1)[-1]}: sheet {sheet!r} months {months}")
    except Exception as e:
        out(f"\n{f.rsplit('/', 1)[-1]}: ERR {type(e).__name__}: {str(e)[:150]}")
