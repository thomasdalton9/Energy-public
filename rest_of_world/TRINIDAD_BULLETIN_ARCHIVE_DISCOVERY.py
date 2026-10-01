"""
TRINIDAD_GAS.py reads only the newest MEEI Consolidated Monthly Bulletin,
which covers the current year to date (Jan-May 2026). The ministry page
links only that one. To backfill from 2021, look for older bulletins:
1. WordPress media library API (lists every uploaded file, linked or not):
   /wp-json/wp/v2/media?search=...
2. The bulletin category listing, paginated, and each post's files.
3. The "historical oil and gas production data" page's files.
For every bulletin-looking xlsx found, print which months its gas sheet
("3A,3B") covers; for the historical files, print their sheets and the
first rows.
"""
import io
import re
from urllib.parse import urljoin

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 120)
SITE = "https://www.energy.gov.tt"
CATEGORY = SITE + "/category/publications/energy-industry-bulletins/"
HIST = SITE + "/data/historical-oil-and-gas-production-data/"


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    try:
        return requests.get(url, headers=H, timeout=T, **kw)
    except Exception as e:
        out(f"  ERR {url}: {e}")
        return None


def files_on(url):
    r = get(url)
    if r is None:
        return []
    return sorted(set(urljoin(url, h) for h in re.findall(r'href="([^"]+\.(?:xlsx?|pdf|csv))"', r.text, re.I)))


files = set()

out("=== 1. WordPress media API")
for term in ["bulletin", "Bulletin", "consolidated", "MEEI", "monthly"]:
    for page in range(1, 6):
        r = get(f"{SITE}/wp-json/wp/v2/media", params={"search": term, "per_page": 100, "page": page})
        if r is None or r.status_code != 200:
            out(f"  media search {term!r} p{page}: HTTP {getattr(r, 'status_code', None)}")
            break
        items = r.json()
        for it in items:
            u = it.get("source_url", "")
            if re.search(r"\.(xlsx?|pdf)$", u, re.I):
                files.add(u)
        out(f"  media search {term!r} p{page}: {len(items)} items")
        if len(items) < 100:
            break

out("\n=== 2. Category listing")
posts = set()
for page in range(1, 15):
    url = CATEGORY if page == 1 else f"{CATEGORY}page/{page}/"
    r = get(url)
    if r is None or r.status_code != 200:
        out(f"  {url}: HTTP {getattr(r, 'status_code', None)}")
        break
    found = set(re.findall(r'href="(https://www\.energy\.gov\.tt/[a-z0-9\-]+/)"', r.text))
    found = {p for p in found if re.search(r"bulletin|20\d\d", p)}
    out(f"  {url}: {len(found)} post links")
    posts |= found
for p in sorted(posts):
    f = files_on(p)
    out(f"  post {p}: {len(f)} files")
    files |= set(f)

out("\n=== 3. Historical production data page")
hist_files = files_on(HIST)
for f in hist_files:
    out("  ", f)

bull = sorted(f for f in files if re.search(r"bulletin", f, re.I))
out(f"\n{len(bull)} bulletin files in total:")
for f in bull:
    out("  ", f)


def gas_sheet_months(content):
    xl = pd.ExcelFile(io.BytesIO(content))
    sheet = next((s for s in xl.sheet_names if s.strip().replace(" ", "") in ("3A,3B", "3A3B")), None)
    if sheet is None:
        return None, xl.sheet_names[:15]
    df = xl.parse(sheet, header=None)
    # months with any numeric data in the column below
    months = []
    for r in range(min(10, len(df))):
        for c, v in enumerate(df.iloc[r]):
            if hasattr(v, "year"):
                col = pd.to_numeric(df.iloc[r + 1:, c], errors="coerce")
                if col.notna().sum() >= 3:
                    months.append(pd.Timestamp(v).strftime("%Y-%m"))
    return sheet, sorted(set(months))


out("\n=== bulletin xlsx coverage")
for f in [f for f in bull if re.search(r"\.xlsx?$", f, re.I)]:
    r = get(f)
    if r is None or r.status_code != 200:
        out(f"  {f}: HTTP {getattr(r, 'status_code', None)}")
        continue
    try:
        sheet, months = gas_sheet_months(r.content)
        out(f"  {f.rsplit('/', 1)[-1]}: sheet {sheet!r} months-with-data {months}")
    except Exception as e:
        out(f"  {f.rsplit('/', 1)[-1]}: ERR {type(e).__name__}: {str(e)[:150]}")

out("\n=== historical files: sheets and first rows")
for f in [f for f in hist_files if re.search(r"\.(xlsx?|csv)$", f, re.I)]:
    r = get(f)
    if r is None or r.status_code != 200:
        out(f"  {f}: HTTP {getattr(r, 'status_code', None)}")
        continue
    try:
        if f.lower().endswith(".csv"):
            sheets = {"csv": pd.read_csv(io.BytesIO(r.content), header=None)}
        else:
            sheets = pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None)
        out(f"\n  ##### {f.rsplit('/', 1)[-1]}: sheets {list(sheets)[:20]}")
        for name, df in list(sheets.items())[:8]:
            out(f"   --- sheet {name!r} shape {df.shape}")
            for row in df.head(12).itertuples(index=False):
                out("      ", [str(x)[:16] for x in row if str(x) != "nan"][:14])
            out("      ...last row:", [str(x)[:16] for x in df.iloc[-1] if str(x) != "nan"][:14])
    except Exception as e:
        out(f"  {f.rsplit('/', 1)[-1]}: ERR {type(e).__name__}: {str(e)[:150]}")
