"""
Round 2-3 of the Peru marginal-cost gap probe (round 1: PERU_COES_NODE_DISCOVERY.py found that COES's
costos-marginales ExportarMasivo file leaves out whole half-hours for EVERY node - 936 of 1,440 in Mar-2026 -
so another node can't fill Santa Rosa's gaps).

Does COES's daily operation report (IEOD, file browser 'Post Operacion/Reportes/IEOD/<year>/<month>/<day>/')
carry a complete half-hourly marginal cost for SANTA ROSA 220 on days the export is short? For a few such days
it lists the day folder (and sub-folders), opens every spreadsheet and prints each sheet that mentions marginal
costs, with how many half-hourly Santa Rosa values it holds and, for comparison, the export's values that day.

Usage: python3 PERU_COES_IEOD_CMG_DISCOVERY.py [YYYY-MM-DD ...]
"""
import html
import io
import re
import sys
import time
import zipfile
from urllib.parse import quote

import pandas as pd
import requests

PORTAL = "https://www.coes.org.pe/Portal/"
EXPORT = PORTAL + "mercadomayorista/costosmarginales/ExportarMasivo"
S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                                "Chrome/124.0 Safari/537.36"})
DAYS = sys.argv[1:] or ["2026-03-05", "2023-09-16", "2024-04-19", "2026-07-25", "2025-07-10"]   # last = complete day
MONTHS_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "setiembre", "octubre",
             "noviembre", "diciembre"]


def req(method, url, **kw):
    for i in range(4):
        try:
            r = S.request(method, url, timeout=180, **kw)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            print(f"    retry {i + 1} {url[:80]}: {type(e).__name__}", flush=True)
            time.sleep(5 * (i + 1))
    return None


def browse(path):
    r = req("POST", PORTAL + "browser/vistadatos", data={"baseDirectory": path, "url": path, "indicador": "",
                                                          "initialLink": "", "orderFolder": ""})
    if r is None:
        return []
    out = []
    for m in re.finditer(r"openBlob\('([^']+)',\s*'(\w)'", r.text):
        it = (html.unescape(m.group(1)), m.group(2))
        if it not in out:
            out.append(it)
    return out


def pick(items, *patterns):
    for pat in patterns:
        for p, k in items:
            if k == "D" and re.search(pat, p.rstrip("/").split("/")[-1], re.I):
                return p
    return None


def export_day(d):
    a = pd.Timestamp(d)
    r = req("GET", f"{EXPORT}?fechaInicio={a:%d/%m/%Y}&fechaFin={a + pd.Timedelta(days=1):%d/%m/%Y}")
    if r is None or r.content[:2] != b"PK":
        return None
    raw = pd.read_excel(io.BytesIO(r.content), header=None)
    hdr = next(i for i in range(15) if "NOMBRE BARRA" in [str(v).strip() for v in raw.iloc[i].values])
    raw.columns = [str(c).strip() for c in raw.iloc[hdr].values]
    raw = raw.iloc[hdr + 1:]
    sr = raw[raw["NODO EMD"].astype(str).str.strip() == "STAROSA220"]
    t = pd.to_datetime(sr["FECHA HORA"].astype(str), format="%d/%m/%Y %H:%M", errors="coerce")
    return pd.Series(pd.to_numeric(sr["TOTAL"], errors="coerce").values, index=t).dropna()


def scan(content, name):
    try:
        xl = pd.ExcelFile(io.BytesIO(content))
    except Exception as e:  # noqa: BLE001
        print(f"      {name}: not a workbook ({type(e).__name__})", flush=True)
        return
    print(f"      {name}: sheets {xl.sheet_names}", flush=True)
    for sh in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sh, header=None)
        txt = df.astype(object).where(df.notna(), "").astype(str)
        flat = " ".join(txt.values.ravel()[:20000]).upper()
        if not re.search(r"COSTO[S]? MARGINAL|CMG|C\.M\.|BARRA", flat):
            continue
        mask = txt.apply(lambda c: c.str.contains(r"SANTA ROSA|STAROSA|S\.ROSA|SROSA", case=False, regex=True).fillna(False))
        hits = [(int(i), int(j)) for i, j in zip(*mask.values.nonzero())]
        print(f"        sheet '{sh}' {df.shape}: mentions marginal cost; Santa Rosa cells: {hits[:6]}", flush=True)
        for i, j in hits[:2]:
            col = pd.to_numeric(df.iloc[i + 1:i + 60, j], errors="coerce").dropna()
            row = pd.to_numeric(df.iloc[i, j + 1:j + 60], errors="coerce").dropna()
            print(f"          at ({i},{j}) '{df.iat[i, j]}': numbers below {len(col)} {list(col.round(2)[:6])} | "
                  f"to the right {len(row)} {list(row.round(2)[:6])}", flush=True)
        # header rows (first 12) to see the layout
        print("          head:", " || ".join(" | ".join(str(v)[:18] for v in df.iloc[r, :12]) for r in range(min(6, len(df)))),
              flush=True)


for d in DAYS:
    day = pd.Timestamp(d)
    print(f"\n===== {d}", flush=True)
    ex = export_day(d)
    if ex is not None:
        print(f"  export: {len(ex)} half-hours at STAROSA220, mean {ex.mean():.2f}; times missing: "
              f"{sorted(set(pd.date_range(day + pd.Timedelta(minutes=30), periods=48, freq='30min').strftime('%H:%M')) - set(ex.index.strftime('%H:%M')))[:20]}",
              flush=True)
    year_items = browse(f"Post Operación/Reportes/IEOD/{day.year}/")
    mpath = pick(year_items, rf"^0?{day.month}\b", rf"^{day.month:02d}", MONTHS_ES[day.month - 1])
    print(f"  year folder: {len(year_items)} items; month folder: {mpath}", flush=True)
    if not mpath:
        print("   ", [p for p, k in year_items][:20], flush=True)
        continue
    month_items = browse(mpath)
    dpath = pick(month_items, rf"^0?{day.day}$", rf"^{day.day:02d}\b", rf"\b{day.day:02d}[./-]?{day.month:02d}")
    print(f"  day folder: {dpath}  (month has {len(month_items)} items, e.g. {[p.split('/')[-2] if p.endswith('/') else p.split('/')[-1] for p, k in month_items[:5]]})", flush=True)
    if not dpath:
        continue
    queue, files = [dpath], []
    for _ in range(3):   # day folder + two levels of sub-folders
        nxt = []
        for path in queue:
            for p, k in browse(path):
                (nxt if k == "D" else files).append(p)
        queue = nxt
    print(f"  files ({len(files)}): {[f.split('/')[-1] for f in files]}", flush=True)
    for f in files:
        if not re.search(r"cmg.*\.zip$", f, re.I):   # round 2: the day folder's marginal-cost zip
            continue
        r = req("GET", PORTAL + "browser/download?url=" + quote(f))
        if r is None:
            continue
        z = zipfile.ZipFile(io.BytesIO(r.content))
        print(f"    {f.split('/')[-1]}: {len(r.content) / 1e6:.1f} MB, members "
              f"{[(i.filename, i.file_size) for i in z.infolist()][:30]}", flush=True)
        for i in z.infolist()[:12]:
            data = z.read(i)
            low = i.filename.lower()
            if low.endswith((".xlsx", ".xls", ".xlsm")):
                scan(data, i.filename)
            elif low.endswith((".csv", ".txt", ".prn", ".dat")):
                text = data.decode("latin-1", "replace")
                lines = text.splitlines()
                sr = [ln for ln in lines if re.search(r"SANTA ROSA|STAROSA|SROSA", ln, re.I)]
                print(f"      {i.filename}: {len(lines)} lines; head {lines[:3]}; Santa Rosa lines {len(sr)}: {sr[:3]}",
                      flush=True)
            else:
                print(f"      {i.filename}: {data[:120]!r}", flush=True)
