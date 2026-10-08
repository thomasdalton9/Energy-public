"""
Pull China's monthly installed power capacity by type from the National Energy Administration (NEA, nea.gov.cn) releases
"国家能源局发布1-N月份全国电力(工业)统计数据" (national power industry statistics).

Each release has a table: cumulative installed generating capacity at the end of the statistics month (total, hydro,
thermal, nuclear, wind, solar; unit 万千瓦 = 10,000 kW; / 100 = GW) and, as published, the capacity ADDED in the year to
date by type, plus operating hours, investment etc. (not read). The table is HTML text in most releases and a PICTURE
(png/jpg) in some, mostly 2025-26 releases: those are read by OCR (tesseract on each table cell, grid found with OpenCV),
and every release - OCR or HTML - must pass these checks or it is left out and listed as failed on the 'Releases' sheet:
  - the five types add up to the total within 15 万千瓦 (NEA's own total differs from the sum of its rounded rows by 4-6),
  - the capacity added in the year to date by type adds up to its total (15 万千瓦),
  - total / solar / wind agree with the 亿千瓦 figures in the release's text summary (0.06 亿千瓦).
A value the OCR cannot read is left blank, never guessed.

  Data      end-of-month capacity, GW, and the year-to-date additions (GW) as published; one row per month for which NEA
            published a release (no January rows: NEA publishes January-February together, giving the end-February stock;
            months with no release in the list are simply absent - nothing is interpolated or derived)
  Releases  one row per release: date, title, url, period, method, status

Release finding (the list page is script-rendered; its ds_*.json holds the whole list): asia/nea_common.py.
Incremental: releases already read with status 'ok' are not fetched again, except the two latest.

    python3 asia/CHINA_NEA_CAPACITY.py --out "output/Data and Chart Outputs/china_nea_capacity_monthly.xlsx"
Needs for image releases: tesseract-ocr + tesseract-ocr-chi-sim (apt), opencv-python-headless, numpy (pip).
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from urllib.parse import urljoin

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import nea_common as nc  # noqa: E402
import xlsx_notes  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "china_nea_capacity_monthly.xlsx")
TYPES = ["Hydro", "Thermal", "Nuclear", "Wind", "Solar"]
STOCK_COLS = ["Total_GW"] + [f"{t}_GW" for t in TYPES] + ["Other_GW"]
ADD_COLS = ["Added_YTD_Total_GW"] + [f"Added_YTD_{t}_GW" for t in TYPES]
COLS = STOCK_COLS + ADD_COLS + ["Method", "Release_Date", "Release_URL"]
TYPE_CHARS = {"Hydro": "水", "Thermal": "火", "Nuclear": "核", "Wind": "风", "Solar": "太阳"}
SUM_TOL = 15        # 万千瓦


# ------------------------------------------------------------------ period of a release
def period_of(title, release_date):
    """Month-end the release describes (Timestamp of the 1st of that month), from its title and publication date."""
    pub = pd.Timestamp(release_date)
    ym = re.search(r"(\d{4})年", title)
    mm = re.search(r"1\s*[-–—~～－﹣]\s*(\d{1,2})月", title)
    if mm:
        n = int(mm.group(1))
        y = int(ym.group(1)) if ym else (pub.year if n <= pub.month else pub.year - 1)
        return pd.Timestamp(y, n, 1)
    if ym:                                  # '2021年全国电力工业统计数据' = the whole year
        return pd.Timestamp(int(ym.group(1)), 12, 1)
    return None


# ------------------------------------------------------------------ HTML table (as text, cells separated by ' | ')
def rows_from_text(text):
    """[(label, [cells])] of the first '一览表' table of a page's text (row starts are cells that begin on a new line)."""
    m = re.search(r"指\s*标\s*名\s*称", text) or re.search("一览表", text)
    if not m:
        return []
    toks = text[m.start():].split("|")
    rows, cur = [], None
    for tok in toks:
        s = tok.strip()
        new_row = bool(re.match(r"^[ \t]*\n[ \t]*\S", tok)) or cur is None
        if new_row and s:
            if cur:
                rows.append(cur)
            cur = [s]
        elif cur is not None:
            cur.append(s)
    if cur:
        rows.append(cur)
    return rows


def num(s):
    s = re.sub(r"[*▲\s,]", "", s or "")
    return float(s) if re.fullmatch(r"-?\d+(\.\d+)?", s) else None


def html_rows(text):
    """-> [(label, value_str)] with the value being the last-but-one cell (the cumulative figure; the last is its y/y)."""
    out = []
    for cells in rows_from_text(text):
        label = re.sub(r"\s+", "", cells[0])
        vals = list(cells[2:])
        while vals and vals[-1] == "":
            vals.pop()
        out.append((label, vals[-2] if len(vals) >= 2 else ""))
    return out


# ------------------------------------------------------------------ OCR of an image table
def ensure_ocr():
    if shutil.which("tesseract") is None:
        pre = [] if os.geteuid() == 0 else ["sudo"]
        subprocess.run(pre + ["apt-get", "install", "-y", "-q", "tesseract-ocr", "tesseract-ocr-chi-sim"], check=False,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        import cv2  # noqa: F401
    except ImportError:
        subprocess.run([sys.executable, "-m", "pip", "install", "-q", "opencv-python-headless"], check=False)
    return shutil.which("tesseract") is not None


def _cluster(idx, gap=4):
    out = []
    for i in idx:
        if out and i - out[-1][-1] <= gap:
            out[-1].append(i)
        else:
            out.append([i])
    return [int(sum(c) / len(c)) for c in out]


def _ocr(img, lang, psm, whitelist=None):
    import cv2
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
        p = f.name
    cv2.imwrite(p, img)
    cmd = ["tesseract", p, "-", "-l", lang, "--psm", str(psm)]
    if whitelist:
        cmd += ["-c", f"tessedit_char_whitelist={whitelist}"]
    out = subprocess.run(cmd, capture_output=True, text=True).stdout.strip()
    os.unlink(p)
    return out


def _cell(g, x0, x1, y0, y1, scale=3):
    import cv2
    c = g[y0 + 3:y1 - 2, x0 + 4:x1 - 3]
    if c.size == 0:
        return None
    c = cv2.resize(c, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    return cv2.copyMakeBorder(c, 20, 20, 20, 20, cv2.BORDER_CONSTANT, value=255)


def _vote(g, x0, x1, y0, y1):
    """Digits of a value cell read three ways (scale / page-segmentation); a value must come out the same at least twice,
    otherwise '' (a blank is retried by the other variants, never guessed)."""
    got = []
    for scale, psm in ((3, 7), (4, 7), (5, 8)):
        c = _cell(g, x0, x1, y0, y1, scale)
        got.append(re.sub(r"[^0-9.\-]", "", _ocr(c, "eng", psm, "0123456789.-*")))
    for v in got:
        if v and got.count(v) >= 2:
            return v
    return ""


def ocr_rows(path):
    """[(label, value_str)] of a table image: grid found with OpenCV, label cells by chi_sim, and (only for the installed
    and added capacity blocks) the cumulative value cell - the second-to-last column - voted over three readings."""
    import cv2
    import numpy as np
    g = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
    if g is None:
        return []
    h, w = g.shape
    b = cv2.threshold(g, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    hk = cv2.morphologyEx(b, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (int(w * 0.35), 1)))
    ys = _cluster(np.where(hk.sum(axis=1) > 0.25 * w * 255)[0])
    vk = cv2.morphologyEx(b, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, int(h * 0.15))))
    xs = _cluster(np.where(vk.sum(axis=0) > 0.15 * h * 255)[0])
    if len(xs) < 5 or len(ys) < 3:
        return []
    vi = len(xs) - 3          # columns: name | unit | [month | y/y |] cumulative | y/y -> the cumulative one
    bands = [(y0, y1) for y0, y1 in zip(ys[:-1], ys[1:]) if y1 - y0 >= 20]
    labels = [re.sub(r"\s+", "", _ocr(_cell(g, xs[0], xs[1], y0, y1), "chi_sim", 7)) for y0, y1 in bands]
    need = set()
    for pred in (is_cap_label, is_add_label):
        i = next((k for k, l in enumerate(labels) if pred(l)), None)
        if i is not None:
            need.update(range(i, min(i + 6, len(labels))))
    ic = next((k for k, l in enumerate(labels) if is_cap_label(l)), None)
    if ic is not None:       # an added-capacity header damaged by OCR: the next capacity-like row after the type rows
        ia = next((k for k, l in enumerate(labels) if k > ic + 5 and ("机容" in l or "装机" in l) and "千伏" not in l), None)
        if ia is not None:
            need.update(range(ia, min(ia + 6, len(labels))))
    return [(labels[k], _vote(g, xs[vi], xs[vi + 1], *bands[k]) if k in need else "") for k in range(len(bands))]


def docx_rows(data):
    """[(label, value_str)] of the first table of a .docx attachment (python-docx)."""
    import io
    from docx import Document
    doc = Document(io.BytesIO(data))
    out = []
    for tbl in doc.tables[:1]:
        for row in tbl.rows:
            cells = [re.sub(r"\s+", "", c.text) for c in row.cells]
            vals = cells[2:]
            while vals and vals[-1] == "":
                vals.pop()
            out.append((cells[0], vals[-2] if len(vals) >= 2 else ""))
    return out


# ------------------------------------------------------------------ rows -> capacity record
def is_cap_label(l):
    """Installed-capacity header row (tolerates OCR damage: '机容' / '装机' / '设备容量' is enough)."""
    return ("机容" in l or "装机" in l or "设备容量" in l) and "新增" not in l and "千伏" not in l


def is_add_label(l):
    return ("机容" in l or "装机" in l) and "新增" in l


def record_from_rows(rows, strict=True, tol_frac=0.0):
    """-> ({'cap': {...}, 'add': {...}, 'note': ...}, None) in 万千瓦, or (None, reason). strict: the types must add up to
    the total (OCR); for a parsed HTML table the published total is kept even where NEA's own rows leave a residual
    (up to 1% in 2023), which then shows as 'Other' in the workbook."""
    labels = [r[0] for r in rows]
    i_cap = next((i for i, l in enumerate(labels) if is_cap_label(l)), None)
    if i_cap is None:
        return None, "no installed-capacity row"
    i_add = next((i for i, l in enumerate(labels) if is_add_label(l)), None)
    if i_add is None:      # header damaged by OCR: the next capacity-like row after the five type rows
        i_add = next((i for i, l in enumerate(labels) if i > i_cap + 5 and ("机容" in l or "装机" in l) and "千伏" not in l), None)
    rec = {"total": num(rows[i_cap][1])}
    for k, t in enumerate(TYPES, 1):
        rec[t] = num(rows[i_cap + k][1]) if i_cap + k < len(rows) else None
        lab = labels[i_cap + k] if i_cap + k < len(rows) else ""
        if TYPE_CHARS[t] not in lab and len(lab) >= 2 and not re.search(r"[A-Za-z0-9]", lab) and lab.count(lab[0]) != len(lab):
            pass      # a garbled OCR label is tolerated: the order of the rows and the sum check decide
    if any(v is None for v in rec.values()):
        return None, "capacity row(s) unreadable"
    if strict and abs(sum(rec[t] for t in TYPES) - rec["total"]) > max(SUM_TOL, tol_frac * rec["total"]):
        return None, f"types sum to {sum(rec[t] for t in TYPES):.0f} vs total {rec['total']:.0f} 万千瓦"
    add = {"total": None, **{t: None for t in TYPES}}
    note = None
    if i_add is not None and i_add + 5 < len(rows):
        add = {"total": num(rows[i_add][1])}
        for k, t in enumerate(TYPES, 1):
            add[t] = num(rows[i_add + k][1])
        vals = [add[t] for t in TYPES]
        if add["total"] is None:
            add, note = {"total": None, **{t: None for t in TYPES}}, "added-capacity total unreadable"
        elif strict and all(v is not None for v in vals) and abs(sum(vals) - add["total"]) > max(SUM_TOL, tol_frac * add["total"]):
            add, note = {"total": None, **{t: None for t in TYPES}}, "added capacity left out: types do not add up to its total"
        elif any(v is None for v in vals):
            note = "an added-capacity cell unreadable (left blank)"
    return {"cap": rec, "add": add, "note": note}, None


def text_check(rec, text):
    """Total / solar / wind against the 亿千瓦 figures of the release's summary sentence."""
    t = re.sub(r"\s+", "", nc.body_of(text))
    probs = []
    for key, pat in (("total", r"累计发电装机容量(?:约)?([\d.]+)亿千瓦"), ("Solar", r"太阳能发电装机容量(?:约)?([\d.]+)亿千瓦"),
                     ("Wind", r"风电装机容量(?:约)?([\d.]+)亿千瓦")):
        m = re.search(pat, t)
        if m and abs(rec["cap"][key] / 10000.0 - float(m.group(1))) > 0.06:
            probs.append(f"{key} {rec['cap'][key] / 10000.0:.2f} vs text {m.group(1)} 亿千瓦")
    return "; ".join(probs) or None


SAVE_IMAGES = [None]


def read_release(r, ocr_ok):
    """-> (period, record, method, status). status 'ok' or a reason."""
    period = period_of(r["title"], r["date"])
    if period is None:
        return None, None, "", "period not found in title"
    html = nc.fetch(r["url"])
    if not html:
        return period, None, "", "page not found"
    text = nc.page_text(html)
    rows = html_rows(text)
    method = "HTML table"
    rec, why = (record_from_rows(rows, strict=False) if rows else (None, "no HTML table"))
    if rec is None:
        imgs = [x for x in re.findall(r'<img[^>]+src="([^"]+)"', html) if not re.search(r"logo|icon|ewm|qr|1\.gif", x, re.I)]
        if not imgs:
            att = re.findall(r'href="([^"]+\.docx)"', html, re.I)
            if att:
                data = nc.fetch(urljoin(r["url"], att[0]), binary=True)
                try:
                    rows = docx_rows(data) if data else []
                except Exception as e:      # noqa: BLE001
                    return period, None, "", f"attachment unreadable ({type(e).__name__})"
                rec, why = record_from_rows(rows, strict=False) if rows else (None, "empty attachment")
                if rec is not None:
                    bad = text_check(rec, text)
                    return (period, None, "Word attachment", f"check failed: {bad}") if bad else (period, rec, "Word attachment", "ok")
            return period, None, "", f"no table ({why})"
        if not ocr_ok:
            return period, None, "OCR", "image table, OCR tools missing"
        data = nc.fetch(urljoin(r["url"], imgs[0]), binary=True)
        if not data:
            return period, None, "OCR", "table image not found"
        if SAVE_IMAGES[0]:
            os.makedirs(SAVE_IMAGES[0], exist_ok=True)
            with open(os.path.join(SAVE_IMAGES[0], f"{r['date']}{os.path.splitext(imgs[0])[1] or '.png'}"), "wb") as f:
                f.write(data)
        with tempfile.NamedTemporaryFile(suffix=os.path.splitext(imgs[0])[1] or ".png", delete=False) as f:
            f.write(data)
            path = f.name
        try:
            rows = ocr_rows(path)
        finally:
            os.unlink(path)
        method = "OCR of image table"
        rec, why = (record_from_rows(rows, tol_frac=0.0 if period >= pd.Timestamp(2024, 1, 1) else 0.012)
                    if rows else (None, "no grid found in image"))
    if rec is None:
        return period, None, method, f"check failed: {why}"
    bad = text_check(rec, text)
    if bad:
        return period, None, method, f"check failed: {bad}"
    return period, rec, method, "ok" + (f" ({rec['note']})" if rec["note"] else "")


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--save-images", default=None, help="development: keep every table image fetched in this folder")
    args = ap.parse_args()
    SAVE_IMAGES[0] = args.save_images
    try:
        held = pd.read_excel(args.out, sheet_name="Data", index_col=0)
        held.index = pd.to_datetime(held.index, errors="coerce")
        held = held[held.index.notna()].reindex(columns=COLS)
    except (FileNotFoundError, ValueError, KeyError):
        held = pd.DataFrame(columns=COLS)
    try:
        rel_old = pd.read_excel(args.out, sheet_name="Releases")
    except (FileNotFoundError, ValueError):
        rel_old = pd.DataFrame(columns=["Release_Date", "Title", "URL", "Period", "Method", "Status"])
    status = dict(zip(rel_old["URL"], rel_old["Status"].astype(str)))
    rels = [r for r in nc.releases() if r["kind"] == "cap" and r["date"] >= "2021-01-01"]
    if not rels:
        raise SystemExit("NEA release list is empty - nothing to do.")
    latest = {r["url"] for r in rels[-2:]}
    ocr_ok = None
    new, log_rows = {}, []
    for r in rels:
        old = rel_old[rel_old["URL"] == r["url"]]
        if str(status.get(r["url"], "")).startswith("ok") and r["url"] not in latest:
            log_rows.append((r["date"], r["title"], r["url"], old["Period"].iloc[0], old["Method"].iloc[0], status[r["url"]]))
            continue
        if ocr_ok is None:
            ocr_ok = ensure_ocr()
        period, rec, method, st = read_release(r, ocr_ok)
        nc.log(f"  {r['date']} {r['title'][:34]} [{period:%Y-%m}] {method}: {st}" if period is not None
               else f"  {r['date']} {r['title'][:34]}: {st}")
        if rec is not None and st.startswith("ok"):
            row = {f"{t}_GW": rec["cap"][t] / 100.0 for t in TYPES}
            row["Total_GW"] = rec["cap"]["total"] / 100.0
            row["Other_GW"] = round(row["Total_GW"] - sum(row[f"{t}_GW"] for t in TYPES), 2)   # total less the five types, as published
            row.update({f"Added_YTD_{t}_GW": (None if rec["add"][t] is None else rec["add"][t] / 100.0) for t in TYPES})
            row["Added_YTD_Total_GW"] = None if rec["add"]["total"] is None else rec["add"]["total"] / 100.0
            row.update({"Method": method, "Release_Date": r["date"], "Release_URL": r["url"]})
            new.setdefault(period, row)          # the earliest release of a period wins
        log_rows.append((r["date"], r["title"], r["url"], f"{period:%Y-%m}" if period is not None else "", method, st))
    d = held
    if new:
        add = pd.DataFrame.from_dict(new, orient="index", columns=COLS)
        d = add.combine_first(held) if len(held) else add
    d = d.reindex(columns=COLS).sort_index()
    d.index.name = "month"
    if d.empty:
        raise SystemExit("No data at all - nothing to save.")
    out = {"Data": d, "Releases": pd.DataFrame(log_rows, columns=["Release_Date", "Title", "URL", "Period", "Method", "Status"]).set_index("Release_Date")}
    ser = [(f"{t}_GW", t, "GW installed", "installed capacity by type", "stacked_bar") for t in TYPES + ["Other"]]
    out["Series"] = pd.DataFrame(ser, columns=["column", "label", "unit", "chart", "kind"]).set_index("column")
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, out, NOTES_LINES, NOTES_SECTION_TITLES)
    failed = out["Releases"][~out["Releases"]["Status"].astype(str).str.startswith("ok")]
    miss = [f"{p:%Y-%m}" for p in pd.date_range(d.index.min(), d.index.max(), freq="MS") if p not in d.index]
    print(f"Saved {len(d)} month(s) ({d.index.min():%Y-%m}..{d.index.max():%Y-%m}); months without a release/reading: {miss}; "
          f"{len(failed)} release(s) failed to read to {args.out}")


NOTES_LINES = [
    "UNITS",
    "*_GW: cumulative installed generating capacity at the end of the month, GW (source unit 万千瓦, 10,000 kW, / 100): "
    "Total, Hydro (incl. pumped storage), Thermal (coal, gas, biomass etc.), Nuclear, Wind, Solar (NEA's categories; the "
    "five types add up to the total within a few 万千瓦 in 2020-22 and from Dec 2023; in 2023 NEA's own monthly rows left up "
    "to 26 GW unallocated, shown as Other_GW = total less the five types, as published; OCR releases must add up). Added_YTD_*_GW: capacity added in the "
    "year to date (January to the statistics month) by type as published; blank where the release has no such rows "
    "(some 2026 releases) or a cell could not be read.",
    "One row per month for which NEA published a release. January has no row of its own: NEA publishes January-February "
    "together (end-February stock). A month with no release in NEA's list (see the Releases sheet) is absent: nothing is "
    "interpolated or derived. Method: 'HTML table' (parsed from the page) or 'OCR of image table' (tesseract on each cell of "
    "the picture NEA posts instead of a table in some releases; the release is kept only if the sums and the text summary agree).",
    "",
    "SOURCE",
    "National Energy Administration (国家能源局), '国家能源局发布N月份全国电力工业统计数据' (renamed '全国电力统计数据' in 2026) on "
    "https://www.nea.gov.cn/ (press list https://www.nea.gov.cn/xwfb/ and the 综合司 column https://www.nea.gov.cn/sjzz/ghs/).",
    "",
    "UPDATES",
    "Incremental: releases already read are not fetched again, except the latest two. The list page is script-rendered; "
    "the pull reads the ds_*.json files behind it (asia/nea_common.py).",
]
NOTES_SECTION_TITLES = {"UNITS", "SOURCE", "UPDATES"}

if __name__ == "__main__":
    main()
