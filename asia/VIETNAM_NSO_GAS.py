"""
Vietnam natural gas output, monthly, from the National Statistics Office (NSO, formerly GSO) monthly
socio-economic report tables: https://www.nso.gov.vn/bao-cao-tinh-hinh-kinh-te-xa-hoi-hang-thang/
Found via discovery_archive/asia/GAS_SSEA_DISCOVERY3.py / 4.py.

Each month NSO publishes a workbook of statistical tables ("Biểu", e.g. 02.-Bieu-thang-9.2026.xlsx) with a sheet of
main industrial products ("Một số sản phẩm chủ yếu của ngành công nghiệp", sheet 10.SPCNthang / SP / 03SPCN) holding
  Khí đốt thiên nhiên dạng khí   (natural gas, gaseous)   Triệu m3 (million m3 in the month)
  Khí hoá lỏng (LPG)                                       Nghìn tấn (thousand tonnes)
for the previous month (Sơ bộ = preliminary, or Thực hiện = actual) and the current month (Ước tính = estimate).
The workbooks are found through the site's WordPress media library (wp-json/wp/v2/media?search=Bieu).

Writes output/Data and Chart Outputs/vietnam_gas.xlsx:
  Production  monthly: Natural_gas_mcm (million m3 in the month), Natural_gas_mcm_per_day (= month / days),
              LPG_kt (thousand tonnes), Status (Actual / Preliminary / Estimate), File (workbook it came from)
  Files       every workbook already read (the incremental record)

Incremental: the Production and Files sheets are the history store. Each run lists the media library and downloads only
workbooks not read yet (the newest is re-read if its URL is new; a month's value is replaced when a later workbook
gives a better status - Actual > Preliminary > Estimate - or the same status from a later upload). Runs on the 1st and 15th.

    python3 asia/VIETNAM_NSO_GAS.py [--out path] [--max-files N]
"""
import argparse
import io
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
import requests
import urllib3

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

urllib3.disable_warnings()
MEDIA = "https://www.nso.gov.vn/wp-json/wp/v2/media"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 150)
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "vietnam_gas.xlsx")
START = "2014-01-01"   # media uploaded before this are not monthly report tables
SKIP = re.compile(r"CPI|SCOLI|ksms|Lao-?dong|Du-bao|gia-vang|baocao_Q|Bieu_baocao|chuyen-de|XHKD|Khung|mau-bao-cao|"
                  r"thong-cao|Bieu-so-0|HTQT|Chuong|Bieu-do|VDS|Bieu-1_|Bieu-2_|_EN|-EN|ENG|\.En_|E\.xls", re.I)
RANK = {"Actual": 3, "Preliminary": 2, "Estimate": 1}
GAS = re.compile(r"kh[íi]\s*đốt\s*thi[êe]n\s*nhi[êe]n", re.I)
LPG = re.compile(r"kh[íi]\s*ho[áa]\s*l[ỏo]ng|LPG", re.I)


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    for i in range(4):
        try:
            r = requests.get(url, headers=H, timeout=T, verify=False, **kw)
            if r.status_code in (400, 404):
                return r
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            if i == 3:
                raise
            out(f"  retry {url[:100]}: {e}")
            time.sleep(5 * (i + 1))


def list_workbooks():
    """[(upload date, url)] of the monthly statistical-table workbooks, oldest first."""
    files, page = [], 1
    while True:
        r = get(MEDIA, params={"search": "Bieu", "per_page": 100, "page": page, "orderby": "date", "order": "desc",
                               "_fields": "date,source_url"})
        if r.status_code != 200:
            break
        js = r.json()
        if not js:
            break
        for m in js:
            u = m["source_url"]
            if re.search(r"\.xlsx?$", u, re.I) and not SKIP.search(u.rsplit("/", 1)[-1]) and m["date"] >= START:
                files.append((m["date"], u))
        if page >= int(r.headers.get("X-WP-TotalPages", page)):
            break
        page += 1
    return sorted(set(files))


def _status(t):
    t = t.lower()
    if "thực hiện" in t or "chính thức" in t:
        return "Actual"
    if "sơ bộ" in t:
        return "Preliminary"
    if "ước" in t:
        return "Estimate"
    return None


def parse_workbook(content, url):
    """{(month Timestamp, item): (value, status)} from the industrial-products sheet of one workbook."""
    book = None
    for how in ("auto", "openpyxl", "xlrd-cp1258", "xlrd-cp1252"):
        try:
            if how == "auto":
                book = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None)
            elif how == "openpyxl":
                book = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None, engine="openpyxl")
            else:   # old .xls whose code page xlrd cannot guess
                import xlrd
                wb = xlrd.open_workbook(file_contents=content, encoding_override=how.split("-")[1])
                book = pd.read_excel(wb, sheet_name=None, header=None, engine="xlrd")
            break
        except Exception:  # noqa: BLE001
            continue
    if book is None:
        out(f"  unreadable {url.rsplit('/', 1)[-1]}")
        return {}
    m = re.search(r"/uploads/(\d{4})/", url)
    year_hint = int(m.group(1)) if m else None
    res = {}
    for name, df in book.items():
        if df.shape[1] < 3:
            continue
        col0 = df.iloc[:, 0].fillna("").map(str)
        gas_rows = [i for i in range(len(df)) if GAS.search(col0.iat[i])]
        if not gas_rows or "quy" in str(name).lower() or "quý" in str(name).lower():
            continue
        first = gas_rows[0]
        # header = the rows above the first product row (up to 12)
        top = max(0, first - 12)
        hdr = {}
        for j in range(2, df.shape[1]):
            t = " ".join(str(df.iat[i, j]) for i in range(top, first) if pd.notna(df.iat[i, j]))
            hdr[j] = re.sub(r"\s+", " ", t)
        # the status words can sit in a merged cell to the left: carry them forward across the header
        last_status = None
        years = [int(y) for t in hdr.values() for y in re.findall(r"(20\d\d)", t)]
        file_year = max(set(years), key=years.count) if years else year_hint
        cols = {}
        for j, t in hdr.items():
            st = _status(t) or last_status
            if _status(t):
                last_status = _status(t)
            if not t or "so với" in t.lower() or "%" in t or "cộng dồn" in t.lower():
                continue
            t2 = re.sub(r"\d+\s*tháng", " ", t, flags=re.I)        # drop cumulative 'N tháng'
            mm = re.search(r"tháng\s*0?(\d{1,2})(?!\d)", t2, re.I)
            if not mm or not st:
                continue
            mon = int(mm.group(1))
            yy = re.search(r"(?:năm|/)\s*(20\d\d)", t2)
            yr = int(yy.group(1)) if yy else file_year
            if not (1 <= mon <= 12) or not yr:
                continue
            cols[j] = (yr, mon, st, bool(yy))
        if not cols:
            continue
        # a December column without its own year in a January release belongs to the previous year
        jan = any(m_ == 1 for _, m_, _, _ in cols.values())
        cols = {j: (pd.Timestamp(yr - 1 if (m_ == 12 and jan and not explicit) else yr, m_, 1), st)
                for j, (yr, m_, st, explicit) in cols.items()}
        for i in range(first, min(first + 4, len(df))):
            label = col0.iat[i]
            unit = str(df.iat[i, 1]).lower()
            item = ("gas" if GAS.search(label) and "m3" in unit.replace("³", "3") else
                    "lpg" if LPG.search(label) and "tấn" in unit else None)
            if not item:
                continue
            for j, (d, st) in cols.items():
                v = pd.to_numeric(df.iat[i, j], errors="coerce")
                if pd.notna(v) and v > 0:
                    res[(d, item)] = (float(v), st)
        if res:
            break
    return res


def load(path):
    try:
        prod = pd.read_excel(path, sheet_name="Production", index_col=0)
        prod.index = pd.to_datetime(prod.index)
        files = pd.read_excel(path, sheet_name="Files", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame(), pd.DataFrame()
    return prod, files


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--max-files", type=int, default=0, help="test: read at most N new workbooks (newest first)")
    args = ap.parse_args()
    prod, files = load(args.out)
    done = set(files.index) if not files.empty else set()
    listing = list_workbooks()
    new = [(d, u) for d, u in listing if u not in done]
    out(f"{len(listing)} statistical-table workbooks listed, {len(new)} not read yet")
    if args.max_files:
        new = new[-args.max_files:]

    def fetch(item):
        d, u = item
        try:
            r = get(u)
            return d, u, (parse_workbook(r.content, u) if r.status_code == 200 else {})
        except Exception as e:  # noqa: BLE001
            out(f"  {u}: {type(e).__name__}: {e}")
            return d, u, None

    with ThreadPoolExecutor(4) as ex:
        results = list(ex.map(fetch, new))
    # start from the saved table: (month, item) -> (value, status, upload date, file)
    best = {}
    if not prod.empty:
        for d, row in prod.iterrows():
            for item, col in (("gas", "Natural_gas_mcm"), ("lpg", "LPG_kt")):
                if col in row and pd.notna(row[col]):
                    best[(d, item)] = (row[col], row.get("Status", "Estimate"), str(row.get("Uploaded", "")),
                                       row.get("File", ""))
    rows_files = []
    for d, u, res in sorted(results, key=lambda x: x[0]):
        if res is None:
            continue          # download failed: retry next run
        rows_files.append({"url": u, "uploaded": d, "months": len({k[0] for k in res})})
        for (mon, item), (v, st) in res.items():
            cur = best.get((mon, item))
            if cur is None or (RANK[st], d) >= (RANK.get(cur[1], 0), cur[2]):
                best[(mon, item)] = (v, st, d, u.rsplit("/", 1)[-1])
        if res:
            g = [f"{k[0]:%Y-%m}={v[0]:.0f}" for k, v in sorted(res.items()) if k[1] == "gas"]
            out(f"  {u.rsplit('/', 1)[-1]}: gas {', '.join(g)}")
    if not best:
        raise SystemExit("No Vietnam gas data")
    months = sorted({k[0] for k in best})
    rec = []
    for mon in months:
        g, l_ = best.get((mon, "gas")), best.get((mon, "lpg"))
        rec.append({"date": mon, "Natural_gas_mcm": g[0] if g else None, "LPG_kt": l_[0] if l_ else None,
                    "Status": g[1] if g else (l_[1] if l_ else None), "Uploaded": g[2] if g else None,
                    "File": g[3] if g else None})
    prod = pd.DataFrame(rec).set_index("date").sort_index()
    prod.insert(1, "Natural_gas_mcm_per_day", (prod["Natural_gas_mcm"] / prod.index.days_in_month).round(3))
    prod["Natural_gas_mcm"] = prod["Natural_gas_mcm"].round(2)
    prod["LPG_kt"] = pd.to_numeric(prod["LPG_kt"], errors="coerce").round(2)
    prod.index.name = "date"
    nf = pd.DataFrame(rows_files).set_index("url") if rows_files else pd.DataFrame()
    files = pd.concat([files, nf]) if not files.empty else nf
    files = files[~files.index.duplicated(keep="last")]
    files.index.name = "url"
    tail = prod.dropna(subset=["Natural_gas_mcm"]).tail(14)
    out(tail.to_string())
    last12 = prod["Natural_gas_mcm"].dropna().tail(12)
    if len(last12):
        out(f"last 12 months: {last12.sum() / 1000:.2f} bcm/yr = {last12.sum() / 365 / 28.3168:.2f} bcf/d")
    notes = [
        "UNITS",
        "Natural_gas_mcm: natural gas (gaseous) produced in the month, million m3 (NSO 'Triệu m3'). "
        "Natural_gas_mcm_per_day = month total / days in month. 1 million m3 = 35.31 million cubic feet; "
        "1 mcm/d = 35.3 MMSCFD.",
        "LPG_kt: liquefied petroleum gas produced (gas processing plants and refineries), thousand tonnes in the month.",
        "Status: Actual (Thực hiện / official) > Preliminary (Sơ bộ, published a month later) > Estimate (Ước tính, the "
        "current month in each release). Each month keeps the best status seen; the newest month is an estimate.",
        "",
        "COVERAGE",
        f"Monthly {prod.index.min():%Y-%m}..{prod.index.max():%Y-%m} (gaps where a month's workbook is missing or "
        f"unreadable). {len(files)} workbooks read so far (Files sheet). NSO publishes around the 6th of the "
        "following month (estimate of the current month at month-end since 2023).",
        "Domestic production only: Vietnam's LNG imports (Thi Vai terminal, from 2023) are not in these tables.",
        "",
        "SOURCE",
        "National Statistics Office of Vietnam (NSO / GSO), monthly socio-economic report, statistical tables: "
        "'Một số sản phẩm chủ yếu của ngành công nghiệp' (main industrial products). "
        "https://www.nso.gov.vn/bao-cao-tinh-hinh-kinh-te-xa-hoi-hang-thang/",
    ]
    xlsx_notes.write_workbook(args.out, {"Production": prod, "Files": files}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
