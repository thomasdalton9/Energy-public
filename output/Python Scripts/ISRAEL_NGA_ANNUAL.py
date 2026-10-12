"""
Israel natural gas ANNUAL balance from the Natural Gas Authority's yearly 'Natural Gas Sector Review' (סקירת משק הגז
הטבעי, Ministry of Energy and Infrastructure, Hebrew PDF, published in the spring/summer after the year):

  https://www.gov.il/BlobFolder/reports/ng_<YYYY>/he/ng_<YYYY>.pdf   (reviews 2018-2021, underscore)
  https://www.gov.il/BlobFolder/reports/ng-<YYYY>/he/ng-<YYYY>.pdf   (reviews 2022 on, hyphen)

gov.il PAGES are Cloudflare-blocked from GitHub Actions; BlobFolder files download with curl_cffi (impersonate chrome).
Reviews found (Oct 2026): 2018-2025 (no file under either pattern for 2008-2017). Both patterns are tried every run for
every year from 2008 to this year; a known review is downloaded again only when its Last-Modified changes.

WHAT IS READ (BCM per year, as published; nothing is computed except MCM -> BCM / 1000):
  * Year-row tables in the PDF's TEXT layer: 2025 review 'נתוני צריכה - תמונה כללית' (2018-2025: total, export,
    domestic, power, cogeneration, industry total / connected to transmission / distribution / CNG) and 2022 review
    (2016-2022: total, domestic, power, industry incl. distribution, distribution, export total / Jordan / Egypt).
  * Year-column tables (2018, 2019, 2020 reviews 'נתונים ומגמות שנתיות'; 2021 key-figures table): supply total, Tamar,
    LNG buoy, Leviathan, consumption total / domestic / power / industry / export, distribution (MCM).
  * Vector bar charts (2022-2025 reviews): each chart's own DATA LABELS. A chart is found by its white plot box, its
    legend (colour swatch + label), its bars (filled rectangles in a legend colour); each numeric label is matched to
    the bar it sits in (or on top of: the column total when the legend has a total line). Years come from the axis
    year labels, else are counted back from the review year at the right-most column (every chart ends at the review
    year; uneven column spacing = not mapped). Every label must agree with its bar's height on the chart's own scale
    (value / height, median of the chart; 6% or ~1 pt tolerance) or it is dropped; an off-bar label is matched to an
    unlabelled bar of the same column only if exactly one fits the scale. Segments without a printed label stay
    BLANK (bar heights are used only to check, never as data). 2023 review charts and 2023/2024 summary tables are
    images (no text layer): nothing is read from them.
  Precedence: for each year and item the NEWEST review that publishes it wins (table before chart within a review),
  except that an older review's figure with more decimals that rounds to the newest one is kept (7.24 for a 7.2
  label); every value read from every review is kept on the 'By review' sheet, the one used on 'Sources used'.

BASIS: from the 2024 review (for 2023-24) and the 2025 review (for every year) cogeneration is a separate sector.
Older reviews (and the 2024 review for 2020-22) split cogeneration between power and industry: those figures are
in the '..._prev_basis_bcm' columns and are never mixed into the current-basis columns. The 2019 review's 'industry'
for 2019 (2.44) includes the 0.22 exported to Jordan (later reviews: 2.22).

CHECKS ('Checks' sheet, per year): fields + buoy + EMG imports = supply total; domestic + export = total;
power + cogeneration + industry = domestic (and the previous basis); Egypt + Jordan = export; IEC + private = power;
industry parts = industry. Tolerance 0.06 BCM per summed item (labels are rounded to 0.1). A check that fails blanks
the PARTS of that group for that year (note on 'Checks'); a gap that unlabelled chart segments explain (their measured
bar height covers the gap) passes as 'explained by unlabelled segment(s)'.

MONTHLY: the 2022-2025 reviews carry a monthly stacked-area chart of domestic use by sector and export; a series is
kept only when all 12 month labels are matched to its band and they add up to the review's own annual figure
(+-0.15 BCM), else left out (sheet 'Monthly', notes on 'Releases'). The monthly supply-by-field LINE chart is not read
(labels sit above or below crossing lines and cannot be matched reliably).

SCHEDULE (owner's choice): every 5 months (1 Jan, 1 Jun, 1 Nov): the review appears once a year (2025 review released
3 Jun 2026). A run only sends HEAD requests for known reviews and tries the missing years; it exits without rewriting
anything when no new or changed review exists. Probe: discovery_archive/middle_east/ISRAEL_NGA_REVIEW_PROBE.py.
Workflow: .github/workflows/israel_gas_annual.yml.

Usage: python3 middle_east/ISRAEL_NGA_ANNUAL.py [--out ...] [--pdf-dir DIR (read local PDFs ng-YYYY.pdf, no download)]
                                                [--reparse (parse every review again)]
"""

import argparse
import hashlib
import io
import os
import re
import sys
from datetime import datetime, timezone

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "israel_gas_annual_nga.xlsx")
URL_PATTERNS = ["https://www.gov.il/BlobFolder/reports/ng_{y}/he/ng_{y}.pdf",
                "https://www.gov.il/BlobFolder/reports/ng-{y}/he/ng-{y}.pdf"]
FIRST_YEAR = 2008
TOL = 0.06          # BCM per summed item

HEB = re.compile(r"[֐-׿]")
NUM = re.compile(r"^-?\d+(?:\.\d+)?%?$")
YEAR = re.compile(r"^(19|20)\d\d$")
MONTHS = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"]

# Annual columns, in sheet order (all BCM per year)
COLUMNS = [
    "Supply_Tamar_bcm", "Supply_Leviathan_bcm", "Supply_Karish_bcm", "Supply_Yam_Tethys_bcm",
    "Supply_EMG_import_bcm", "Supply_LNG_buoy_bcm", "Supply_total_bcm",
    "Domestic_Power_bcm", "Domestic_Cogeneration_bcm", "Domestic_Industry_bcm",
    "Industry_transmission_bcm", "Industry_distribution_bcm", "Industry_CNG_bcm",
    "Power_IEC_bcm", "Power_private_bcm",
    "Domestic_Power_prev_basis_bcm", "Domestic_Industry_prev_basis_bcm", "Power_private_incl_cogen_prev_basis_bcm",
    "Cogeneration_for_industry_prev_basis_bcm",
    "Domestic_total_bcm", "Export_Egypt_bcm", "Export_Jordan_bcm", "Export_total_bcm", "Total_bcm",
]


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def norm(t):
    """Hebrew label for matching: no spaces, one kind of quote, final letters kept."""
    return re.sub(r"\s+", "", t).replace("״", '"').replace("”", '"').replace("'", "")


# ------------------------------------------------------------------ PDF geometry helpers

def logical(t):
    return t[::-1] if HEB.search(t) else t


def colour(c):
    if c is None:
        return None
    if isinstance(c, (int, float)):
        c = (c, c, c)
    if not all(isinstance(x, (int, float)) for x in c):
        return None
    if len(c) == 1:
        c = (c[0],) * 3
    if len(c) == 4:  # CMYK
        k = c[3]
        c = tuple((1 - x) * (1 - k) for x in c[:3])
    return tuple(round(float(x), 2) for x in c)


def words(page):
    """pdfplumber words with numeric fragments on one baseline merged ('1','.','4' -> '1.4'); Hebrew words come out
    in logical order from pdfplumber here (checked on the 2018-2025 reviews)."""
    if "_words" in page.__dict__:
        return page.__dict__["_words"]
    ws = [dict(w) for w in page.extract_words(x_tolerance=1.5, y_tolerance=1.5, keep_blank_chars=False)]
    ws.sort(key=lambda w: (round(w["top"]), w["x0"]))
    out = []
    frag = re.compile(r"^[\d.%\-]+$")
    for w in ws:
        if out and frag.match(w["text"]) and frag.match(out[-1]["text"]) and abs(out[-1]["top"] - w["top"]) < 1.2 \
                and 0 <= w["x0"] - out[-1]["x1"] < 1.0:
            out[-1]["text"] += w["text"]
            out[-1]["x1"] = w["x1"]
            continue
        out.append(w)
    for w in out:
        w["xc"] = (w["x0"] + w["x1"]) / 2
        w["yc"] = (w["top"] + w["bottom"]) / 2
    page.__dict__["_words"] = out
    return out


def inside(o, b, pad=0.5):
    return o["x0"] >= b[0] - pad and o["x1"] <= b[2] + pad and o["top"] >= b[1] - pad and o["bottom"] <= b[3] + pad


def chart_boxes(page):
    """White filled plot-area rectangles (the charts are drawn on one each), innermost only."""
    W, H = page.width, page.height
    boxes = []
    for r in page.rects:
        c = colour(r.get("non_stroking_color")) if r.get("fill") else None
        if c and all(x > 0.97 for x in c) and 150 < r["width"] < W * 0.95 and 70 < r["height"] < H * 0.9:
            boxes.append((r["x0"], r["top"], r["x1"], r["bottom"]))
    boxes = [b for b in boxes if not any(o != b and o[0] >= b[0] - 1 and o[2] <= b[2] + 1 and o[1] >= b[1] - 1
                                          and o[3] <= b[3] + 1 for o in boxes)]
    return sorted(boxes, key=lambda b: (b[1], b[0]))


def read_chart(page, box):
    ws = [dict(w) for w in words(page) if inside(w, box)]
    rects = [r for r in page.rects if inside(r, box) and r.get("fill")]
    legend, swatch_ids = [], set()
    cands = [(r, colour(r["non_stroking_color"]), "bar") for r in rects if 2.5 < r["height"] <= 6.5 and 2.5 < r["width"] <= 26]
    cands += [(ln, colour(ln.get("stroking_color")), "line") for ln in page.lines if inside(ln, box)
              and abs(ln["top"] - ln["bottom"]) < 1 and 8 < ln["x1"] - ln["x0"] <= 26]
    for r, c, kind in cands:
        yc = (r["top"] + r["bottom"]) / 2
        if not [w for w in ws if abs(w["yc"] - yc) < 4 and r["x1"] - 0.5 <= w["x0"] < r["x1"] + 40 and not NUM.match(w["text"])]:
            continue
        nxt = min([s[0]["x0"] for s in cands if abs((s[0]["top"] + s[0]["bottom"]) / 2 - yc) < 3 and s[0]["x0"] > r["x1"]]
                  or [1e9])
        lab = sorted([w for w in ws if abs(w["yc"] - yc) < 4 and r["x1"] - 0.5 <= w["x0"] < nxt
                      and not NUM.match(w["text"])], key=lambda w: -w["x0"])
        legend.append({"colour": c, "label": " ".join(logical(w["text"]) for w in lab), "kind": kind, "y": yc,
                       "obj": r})
        swatch_ids.add(id(r))
        for w in lab:
            w["legend"] = True
    leg_cols = {lg["colour"]: lg["label"] for lg in legend if lg["kind"] == "bar"}
    bars = []
    for r in rects:
        c = colour(r["non_stroking_color"])
        if id(r) not in swatch_ids and c in leg_cols and r["width"] < 60 and r["height"] > 0.05:
            b = dict(r)
            b["series"] = leg_cols[c]
            b["xc"] = (b["x0"] + b["x1"]) / 2
            bars.append(b)
    return {"box": box, "words": ws, "legend": legend, "bars": bars, "page": page}


def chart_values(ch, review_year):
    """{year: {series: value}} from the chart's data labels, plus {year: {series: measured}} for unlabelled bars
    and notes. See the module docstring for the rules."""
    ws, bars, legend = ch["words"], ch["bars"], ch["legend"]
    notes = []
    if not bars:
        return {}, {}, ["no bars"]
    total_name = next((lg["label"] for lg in legend if lg["kind"] == "line"), None)
    nums = [w for w in ws if not w.get("legend") and NUM.match(w["text"]) and not YEAR.match(w["text"])
            and not w["text"].endswith("%")]
    xs = sorted(b["xc"] for b in bars)
    cols = []
    for x in xs:
        if cols and x - cols[-1][-1] < 2.5:
            cols[-1].append(x)
        else:
            cols.append([x])
    colx = [sum(c) / len(c) for c in cols]
    width = sorted(b["x1"] - b["x0"] for b in bars)[len(bars) // 2]
    for b in bars:
        b["col"] = min(range(len(colx)), key=lambda i: abs(colx[i] - b["xc"]))
    # clustered bars (one series per column, columns in tight groups): merge each group into one category
    gaps = [b - a for a, b in zip(colx, colx[1:])]
    group_of = list(range(len(colx)))
    if gaps and all(len({b["series"] for b in bars if b["col"] == i}) == 1 for i in range(len(colx))):
        small = [g for g in gaps if g < width * 1.3]
        if small and len(small) < len(gaps):
            gi = 0
            for i, g in enumerate(gaps):
                group_of[i + 1] = gi if g < width * 1.3 else gi + 1
                gi = group_of[i + 1]
    ncat = max(group_of) + 1
    catx = [sum(colx[i] for i in range(len(colx)) if group_of[i] == g) / group_of.count(g) for g in range(ncat)]
    yr_axis = [w for w in ws if YEAR.match(w["text"]) and w["top"] > max(b["bottom"] for b in bars) - 2]
    cat_year = {}
    if len(yr_axis) >= 3:
        yx = sorted((w["xc"], int(w["text"])) for w in yr_axis)
        for g, x in enumerate(catx):
            j = min(range(len(yx)), key=lambda j: abs(yx[j][0] - x))
            if abs(yx[j][0] - x) < max(width * 1.6, 12):
                cat_year[g] = yx[j][1]
        method = "axis year labels"
    else:
        cg = [b - a for a, b in zip(catx, catx[1:])]
        step = sorted(cg)[len(cg) // 2] if cg else 1
        if cg and (max(cg) > step * 1.3 or min(cg) < step * 0.7):
            return {}, {}, [f"uneven category spacing (step {step:.1f}); years not mapped"]
        for g in range(ncat):
            cat_year[g] = review_year - (ncat - 1 - g)
        method = "counted back from the review year"
    notes.append(f"{ncat} categories, years from {method}")
    stack = {}
    for b in bars:
        stack.setdefault(b["col"], []).append(b)
    pairs = []
    for w in nums:
        v = float(w["text"])
        inx = [b for b in bars if b["x0"] - 1 <= w["xc"] <= b["x1"] + 1]
        cont = [b for b in inx if b["top"] - 1.5 <= w["yc"] <= b["bottom"] + 1.5]
        if cont:
            b = min(cont, key=lambda b: b["height"])
            pairs.append([v, b["height"], b["series"], b["col"], w, b])
            continue
        if not inx:
            continue
        c = inx[0]["col"]
        top = min(b["top"] for b in stack[c])
        if 0 <= top - w["bottom"] < 18:
            multi = len({b["series"] for b in stack[c]}) > 1 or len([lg for lg in legend if lg["kind"] == "bar"]) > 1
            if total_name and multi:
                h = max(b["bottom"] for b in stack[c]) - top
                pairs.append([v, h, total_name, c, w, None])
            else:
                b = min(stack[c], key=lambda b: b["top"])
                pairs.append([v, b["height"], b["series"], c, w, b])
    ks = sorted(p[0] / p[1] for p in pairs if p[1] > 6 and p[0] > 0)
    if not ks:
        return {}, {}, notes + ["no labelled bars"]
    k = ks[len(ks) // 2]
    labelled = {id(p[5]) for p in pairs if p[5] is not None and abs(p[0] - k * p[1]) <= max(0.06 * p[0], 1.2 * k + 0.06)}
    out = {}
    for v, h, s, c, w, b in pairs:
        tol = max(0.06 * v, 1.2 * k + 0.06)
        if abs(v - k * h) > tol:
            fit = [o for o in stack[c] if id(o) not in labelled and abs(v - k * o["height"]) <= tol]
            if len(fit) == 1:
                s, b = fit[0]["series"], fit[0]
                labelled.add(id(b))
            else:
                notes.append(f"label {w['text']} ({cat_year.get(group_of[c])}, {s}) off the chart scale "
                             f"({k * h:.2f}); not used")
                continue
        yr = cat_year.get(group_of[c])
        if yr is None:
            continue
        d = out.setdefault(yr, {})
        if s in d and d[s] is not None and d[s] != v:
            notes.append(f"two labels for {yr} {s}: {d[s]} / {v}; not used")
            d[s] = None
        elif s not in d:
            d[s] = v
    out = {y: {s: v for s, v in d.items() if v is not None} for y, d in out.items()}
    measured = {}
    for b in bars:
        if id(b) not in labelled:
            yr = cat_year.get(group_of[b["col"]])
            if yr is not None and b["series"] not in out.get(yr, {}):
                measured.setdefault(yr, {})
                measured[yr][b["series"]] = measured[yr].get(b["series"], 0) + k * b["height"]
    return out, measured, notes


# ------------------------------------------------------------------ chart specs (by legend labels)

def _m(label, *keys):
    n = norm(label)
    return any(norm(k) in n for k in keys)


def supply_series(label):
    for keys, col in ((("תטיס",), "Supply_Yam_Tethys_bcm"), (("EMG",), "Supply_EMG_import_bcm"),
                      (("BUOY",), "Supply_LNG_buoy_bcm"), (("תמר",), "Supply_Tamar_bcm"),
                      (("לווייתן", "לוויתן"), "Supply_Leviathan_bcm"), (("כריש",), "Supply_Karish_bcm"),
                      (('סה"כ',), "Supply_total_bcm")):
        if _m(label, *keys):
            return col
    return None


def chart_spec(legend):
    """(name, function label -> column, basis-aware) for the charts used; None for the rest."""
    labs = [lg["label"] for lg in legend]
    has = lambda *k: any(_m(lab, *k) for lab in labs)  # noqa: E731
    if has("תמר") and has("BUOY") and has("לווייתן", "לוויתן"):
        return "supply by source", supply_series
    if has("מצרים") and has("ירדן") and len(labs) <= 3:
        return "exports by country", lambda s: ("Export_Egypt_bcm" if _m(s, "מצרים") else
                                                "Export_Jordan_bcm" if _m(s, "ירדן") else None)
    if has("משק המקומי", "משק מקומי") and has("יצוא", "ייצוא") and len(labs) <= 3:
        return "domestic and export", lambda s: ("Domestic_total_bcm" if _m(s, "מקומי") else
                                                 "Total_bcm" if _m(s, "כלל המשק") else
                                                 "Export_total_bcm" if _m(s, "יצוא", "ייצוא") else None)
    if has("חשמל") and has("תעשייה") and has("יצוא", "ייצוא") and any(lg["kind"] == "line" for lg in legend):
        return "use by sector", lambda s: ("Total_bcm" if _m(s, 'סה"כצריכה', "כלל") else
                                           "Domestic_Power_bcm" if _m(s, "חשמל") else
                                           "Domestic_Cogeneration_bcm" if _m(s, "קוגנרציה") else
                                           "Domestic_Industry_bcm" if _m(s, "תעשייה") else
                                           "Export_total_bcm" if _m(s, "יצוא", "ייצוא") else None)
    if has('חח"י') and has('יח"פים') and not has("תעשייה") and any(lg["kind"] == "line" for lg in legend):
        return "power by producer", lambda s: ("Power_IEC_bcm" if _m(s, 'חח"י') else
                                               "Power_private_incl_cogen_prev_basis_bcm" if _m(s, "קוגנרציה") else
                                               "Power_private_bcm" if _m(s, 'יח"פים') else
                                               "Domestic_Power_bcm" if _m(s, 'סה"כ') else None)
    if has("CNG") and has("ההולכה") and has("חלוקה"):
        return "industry and cogeneration", lambda s: ("Industry_transmission_bcm" if _m(s, "ההולכה") else
                                                       "Industry_distribution_bcm" if _m(s, "חלוקה") else
                                                       "Industry_CNG_bcm" if _m(s, "CNG") else
                                                       "Cogeneration_for_industry_prev_basis_bcm" if _m(s, "קוגנרציהלתעשייה") else
                                                       "Domestic_Cogeneration_bcm" if _m(s, "קוגנרציה") else None)
    return None


def current_basis(review_year, data_year):
    """True when a review's power / industry figures for data_year have cogeneration as a separate sector. The 2024
    review says so for 2023 and 2024 only (its 2020-22 bars are the old split); the 2025 review restates every year;
    earlier reviews put cogeneration partly in power and partly in industry. Later reviews are assumed to keep the
    2025 basis (their sector table is checked against cogeneration + power + industry = domestic)."""
    return review_year >= 2025 or (review_year == 2024 and data_year >= 2023)


PREV = {"Domestic_Power_bcm": "Domestic_Power_prev_basis_bcm",
        "Domestic_Industry_bcm": "Domestic_Industry_prev_basis_bcm"}


def basis_fix(name, review_year, data_year, values, col_of):
    """Map one chart category to columns; previous-basis power/industry go to the '_prev_basis' columns."""
    out = {}
    for s, v in values.items():
        col = col_of(s)
        if col:
            out[col] = v
    if name in ("use by sector", "power by producer") and not current_basis(review_year, data_year):
        for a, b in PREV.items():
            if a in out:
                out[b] = out.pop(a)
        out.pop("Domestic_Cogeneration_bcm", None)
    return out


# ------------------------------------------------------------------ tables in the text layer

def rows_of(page):
    rows = {}
    for w in words(page):
        rows.setdefault(round(w["top"] / 2.0), []).append(w)
    out = []
    for key in sorted(rows):
        r = sorted(rows[key], key=lambda w: w["x0"])
        if out and abs(out[-1][0]["top"] - r[0]["top"]) < 2.2:
            out[-1] = sorted(out[-1] + r, key=lambda w: w["x0"])
        else:
            out.append(r)
    return out


def num(t):
    t = t.replace(",", "")
    return float(t) if NUM.match(t) and not t.endswith("%") else None


def year_row_tables(page):
    """Tables with one row per year (year at the right end of the row): 2025 and 2022 review layouts, identified
    by their header words and checked by their own sums. Returns [(year, {col: value})], notes."""
    text = norm(" ".join(logical(w["text"]) for w in words(page)))
    out, notes = [], []
    for r in rows_of(page):
        toks = [w["text"] for w in r]
        if not toks or not YEAR.match(toks[-1]) or any(t.endswith("%") for t in toks) or all(YEAR.match(t) for t in toks):
            continue
        vals = [num(t) for t in toks[:-1]]
        if any(v is None for v in vals):
            continue
        y = int(toks[-1])
        if len(vals) == 9 and "קוגנרציה" in text and "CNG" in text:
            cng, dist, ind_t, ind, cog, pw, dom, exp, tot = vals
            ok = (abs(ind - (cng + dist + ind_t)) <= 3 * TOL and abs(dom - (pw + cog + ind)) <= 3 * TOL
                  and abs(tot - (dom + exp)) <= 2 * TOL)
            rec = {"Industry_CNG_bcm": cng, "Industry_distribution_bcm": dist, "Industry_transmission_bcm": ind_t,
                   "Domestic_Industry_bcm": ind, "Domestic_Cogeneration_bcm": cog, "Domestic_Power_bcm": pw,
                   "Domestic_total_bcm": dom, "Export_total_bcm": exp, "Total_bcm": tot}
        elif len(vals) == 8 and "למצרים" in text and "לירדן" in text:
            eg, jo, exp, dist, ind, pw, dom, tot = vals
            ok = (abs(exp - (eg + jo)) <= 2 * TOL and abs(tot - (dom + exp)) <= 2 * TOL
                  and abs(dom - (pw + ind)) <= 2 * TOL)
            rec = {"Export_Egypt_bcm": eg, "Export_Jordan_bcm": jo, "Export_total_bcm": exp,
                   "Industry_distribution_bcm": dist, "Domestic_Industry_prev_basis_bcm": ind,
                   "Domestic_Power_prev_basis_bcm": pw, "Domestic_total_bcm": dom, "Total_bcm": tot}
        else:
            continue
        if ok:
            out.append((y, rec))
        else:
            notes.append(f"table row {y} does not add up; not used: {' '.join(toks)}")
    return out, notes


def label_item(label):
    """Row label of a year-column table -> column (or None)."""
    n = norm(label)
    if "%" in label or 'ק"מ' in label or "פרישת" in n:
        return None
    if "אספקת" in n and 'סה"כ' in n:
        return "Supply_total_bcm"
    if "מקשרימי" in n:
        return "Supply_LNG_buoy_bcm"
    if "לוויתן" in n or "לווייתן" in n:
        return "Supply_Leviathan_bcm"
    if "כריש" in n:
        return "Supply_Karish_bcm"
    if "תמר" in n:
        return "Supply_Tamar_bcm"
    if "צריכת" in n and ("במשק" in n or "המקומי" in n):
        return "Domestic_total_bcm"
    if "לייצור" in n and "חשמל" in n:
        return "Domestic_Power_prev_basis_bcm"
    if "בתעשייה" in n:
        return "Domestic_Industry_prev_basis_bcm"
    if "החלוקה" in n:
        return "Industry_distribution_bcm"
    if "לייצוא" in n or "לייצוא" in n:
        return "Export_total_bcm"
    if 'סה"כצריכת' in n:
        return "Total_bcm"
    return None


def year_col_tables(page):
    """Tables with a header row of years and one row per item (2018-2021 reviews). Label words (right of the year
    columns) are given to the nearest value row. Returns [(year, {col: value})], notes."""
    rws = rows_of(page)
    out, notes = {}, []
    for hi, hr in enumerate(rws):
        yrs = [w for w in hr if YEAR.match(w["text"])]
        if len(yrs) < 4:
            continue
        ycols = sorted((w["xc"], int(w["text"])) for w in yrs)
        step = min(b[0] - a[0] for a, b in zip(ycols, ycols[1:]))
        xmax = max(w["x1"] for w in yrs) + 8
        vrows, lab_words = [], []
        for r in rws[hi + 1:]:
            if len([w for w in r if YEAR.match(w["text"])]) >= 4:
                break
            vals = {}
            for w in r:
                if w["xc"] < xmax:
                    j = min(range(len(ycols)), key=lambda j: abs(ycols[j][0] - w["xc"]))
                    if abs(ycols[j][0] - w["xc"]) < step * 0.45:
                        vals[ycols[j][1]] = w["text"]
            lab_words += [w for w in r if w["x0"] >= xmax - 8 and w["xc"] > xmax]
            if len(vals) >= len(ycols) - 2 and all(num(t) is not None or t in ("-", "0") or t.endswith("%")
                                                   for t in vals.values()):
                vrows.append((r, vals))
        if not vrows:
            continue
        labels = {i: [] for i in range(len(vrows))}
        for w in lab_words:
            i = min(range(len(vrows)), key=lambda i: abs(vrows[i][0][0]["yc"] - w["yc"]))
            labels[i].append(w)
        for i, (r, vals) in enumerate(vrows):
            lw = sorted(labels[i], key=lambda w: (round(w["top"]), -w["x0"]))
            label = " ".join(logical(w["text"]) for w in lw)
            col = label_item(label)
            if not col or any(t.endswith("%") for t in vals.values()):
                continue
            scale = 0.001 if "MCM" in label else 1.0
            for y, t in vals.items():
                v = num(t)
                if v is not None:
                    out.setdefault(y, {})[col] = round(v * scale, 6)
    return sorted(out.items()), notes


# ------------------------------------------------------------------ monthly (area / line charts)

def _ys_at(obj, x):
    """y values where a path crosses the vertical line at x."""
    pts = obj.get("pts") or []
    if pts:
        lo, hi = min(p[0] for p in pts), max(p[0] for p in pts)
        if lo - 12 <= x < lo:
            x = lo + 0.01
        elif hi < x <= hi + 12:
            x = hi - 0.01
    ys = []
    for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
        if (x1 - x) * (x2 - x) <= 0 and x1 != x2:
            ys.append(y1 + (y2 - y1) * (x - x1) / (x2 - x1))
        elif x1 == x2 == x:
            ys += [y1, y2]
    return ys


def monthly_chart(page, box, review_year):
    """Month x series values from a monthly area/line chart: legend swatches -> filled areas (label inside the band)
    or lines (label nearest the line, within 9 pt). Returns ({series: [12 values or None]}, legend labels)."""
    ws = [w for w in words(page) if inside(w, box)]
    month_w = {}
    for w in ws:
        lw = logical(w["text"])
        for i, m in enumerate(MONTHS):
            if lw == m or lw.startswith(m):
                month_w[i] = w["xc"]
    if len(month_w) < 12:
        return None, []
    rects = [r for r in page.rects if inside(r, box) and r.get("fill")]
    sw = [(r, colour(r["non_stroking_color"])) for r in rects if 2.5 < r["height"] <= 6.5 and 2.5 < r["width"] <= 26]
    sw += [(ln, colour(ln.get("stroking_color"))) for ln in page.lines if inside(ln, box)
           and abs(ln["top"] - ln["bottom"]) < 1 and 8 < ln["x1"] - ln["x0"] <= 26]
    curves = [c for c in page.curves if inside(c, box, 2)]
    legend = {}
    for r, c in sw:
        yc = (r["top"] + r["bottom"]) / 2
        nxt = min([s[0]["x0"] for s in sw if abs((s[0]["top"] + s[0]["bottom"]) / 2 - yc) < 3 and s[0]["x0"] > r["x1"]]
                  or [1e9])
        lab = sorted([w for w in ws if abs(w["yc"] - yc) < 4 and r["x1"] - 0.5 <= w["x0"] < nxt
                      and not NUM.match(w["text"])], key=lambda w: -w["x0"])
        if lab:
            legend[c] = " ".join(logical(w["text"]) for w in lab)
            for w in lab:
                w["legend"] = True
    if not legend:
        return None, []
    areas = {legend[colour(c["non_stroking_color"])]: c for c in curves if c.get("fill")
             and colour(c.get("non_stroking_color")) in legend and c["x1"] - c["x0"] > 100}
    lines_ = {legend[colour(c["stroking_color"])]: c for c in curves if not c.get("fill")
              and colour(c.get("stroking_color")) in legend and c["x1"] - c["x0"] > 100}
    # line charts (monthly supply by field) are not read: their labels sit above or below the points and the
    # lines cross, so labels cannot be matched to a line reliably (checked on the 2025 review: 12-month sums fail)
    series = areas
    if not series:
        return None, list(legend.values())
    vals = {s: [None] * 12 for s in series}
    nums = [w for w in ws if not w.get("legend") and NUM.match(w["text"]) and not w["text"].endswith("%")]
    for w in nums:
        i = min(month_w, key=lambda i: abs(month_w[i] - w["xc"]))
        if abs(month_w[i] - w["xc"]) > 14:
            continue
        hit = []
        for s, c in series.items():
            ys = _ys_at(c, w["xc"])
            if not ys:
                continue
            if areas:
                if min(ys) - 1 <= w["yc"] <= max(ys) + 1:
                    hit.append((max(ys) - min(ys), s))
            else:
                # line labels sit above their point: offset (point y - label centre) of about 4-16 pt
                offs = [y - w["yc"] for y in ys if 3 <= y - w["yc"] <= 17]
                if offs:
                    hit.append((min(offs), s))
        if not hit:
            continue
        hit.sort()
        if len(hit) > 1 and not areas and hit[1][0] - hit[0][0] < 3:
            continue  # two lines equally close: ambiguous
        s = hit[0][1]
        if vals[s][i] is None:
            vals[s][i] = float(w["text"])
        elif vals[s][i] != float(w["text"]):
            vals[s][i] = "x"
    return vals, list(legend.values())


# ------------------------------------------------------------------ one review

def parse_review(content, review_year):
    """All values a review publishes that we can read: [(data_year, column, value, source)], monthly rows, notes."""
    import pdfplumber
    recs, monthly, notes = [], [], []
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for pi, page in enumerate(pdf.pages, 1):
            try:
                t_rows, t_notes = year_row_tables(page)
                notes += [f"p{pi}: {n}" for n in t_notes]
                for y, rec in t_rows:
                    recs += [(y, c, v, f"p{pi} table") for c, v in rec.items()]
                c_rows, _ = year_col_tables(page)
                for y, rec in c_rows:
                    recs += [(y, c, v, f"p{pi} table") for c, v in rec.items()]
            except Exception as e:  # noqa: BLE001
                notes.append(f"p{pi}: table read failed: {e}")
            for box in chart_boxes(page):
                try:
                    ch = read_chart(page, box)
                except Exception as e:  # noqa: BLE001
                    notes.append(f"p{pi}: chart read failed: {e}")
                    continue
                if not ch["bars"]:
                    mv, labs = monthly_chart(page, box, review_year)
                    if mv:
                        monthly.append((pi, mv, labs))
                    continue
                spec = chart_spec(ch["legend"])
                if not spec:
                    continue
                name, col_of = spec
                vals, measured, cnotes = chart_values(ch, review_year)
                if len({round(b["xc"]) for b in ch["bars"]}) > 60:
                    continue   # daily/monthly bar charts
                notes += [f"p{pi} {name}: {n}" for n in cnotes]
                for y, d in vals.items():
                    for c, v in basis_fix(name, review_year, y, d, col_of).items():
                        recs.append((y, c, v, f"p{pi} chart '{name}'"))
                for y, d in measured.items():
                    for c, v in basis_fix(name, review_year, y, d, col_of).items():
                        if c:
                            recs.append((y, "measured:" + c, round(v, 3), f"p{pi} chart '{name}' (bar height)"))
    return recs, monthly, notes


# ------------------------------------------------------------------ annual table and checks

CHECKS = [  # (name, total column, parts, parts that may be missing because they did not exist that year)
    ("supply by source", "Supply_total_bcm", ["Supply_Yam_Tethys_bcm", "Supply_EMG_import_bcm", "Supply_LNG_buoy_bcm",
                                              "Supply_Tamar_bcm", "Supply_Leviathan_bcm", "Supply_Karish_bcm"]),
    ("domestic + export", "Total_bcm", ["Domestic_total_bcm", "Export_total_bcm"]),
    ("sectors", "Domestic_total_bcm", ["Domestic_Power_bcm", "Domestic_Cogeneration_bcm", "Domestic_Industry_bcm"]),
    ("sectors (previous basis)", "Domestic_total_bcm", ["Domestic_Power_prev_basis_bcm",
                                                        "Domestic_Industry_prev_basis_bcm"]),
    ("industry parts", "Domestic_Industry_bcm", ["Industry_transmission_bcm", "Industry_distribution_bcm",
                                                 "Industry_CNG_bcm"]),
    ("export by country", "Export_total_bcm", ["Export_Egypt_bcm", "Export_Jordan_bcm"]),
    ("power by producer", "Domestic_Power_bcm", ["Power_IEC_bcm", "Power_private_bcm"]),
    ("power by producer (previous basis)", "Domestic_Power_prev_basis_bcm", ["Power_IEC_bcm",
                                                                             "Power_private_incl_cogen_prev_basis_bcm"]),
]


def pick(byrev):
    """Annual values: newest review first, table before chart; also the review/source of each value."""
    vals, src, measured = {}, {}, {}
    order = byrev.assign(_t=byrev["Source"].str.contains("table").astype(int)) \
                 .sort_values(["Review", "_t"], ascending=[False, False])
    for r in order.itertuples(index=False):
        if r.Column.startswith("measured:"):
            key = (r.Year, r.Column.split(":", 1)[1])
            measured.setdefault(key, (r.Review, r.Value))
            continue
        key = (r.Year, r.Column)
        if key not in vals:
            vals[key] = r.Value
            src[key] = f"{r.Review} review {r.Source}"
        elif decimals(r.Value) > decimals(vals[key]) and \
                round(r.Value, decimals(vals[key])) == round(vals[key], decimals(vals[key])):
            # an older review gives the SAME figure to more decimals (e.g. 7.24 vs a chart label 7.2): keep it
            src[key] = f"{r.Review} review {r.Source} (same figure as {src[key].split(' review')[0]} review's " \
                       f"{vals[key]}, more decimals)"
            vals[key] = r.Value
    return vals, src, measured


def decimals(v):
    """Decimals of a published figure; at least 1 (labels like '22.0' are read as the float 22.0)."""
    t = f"{v:.6f}".rstrip("0")
    return max(1, len(t.split(".")[1]) if "." in t else 0)


def build_annual(byrev):
    vals, src, measured = pick(byrev)
    years = sorted({y for y, _ in vals})
    ann = pd.DataFrame(index=years, columns=COLUMNS, dtype=float)
    for (y, c), v in vals.items():
        if c in ann.columns:
            ann.loc[y, c] = v
    checks, blank = [], []
    for y in years:
        for name, tot, parts in CHECKS:
            t = ann.loc[y, tot]
            have = [p for p in parts if pd.notna(ann.loc[y, p])]
            if pd.isna(t) or len(have) < 2 and not (len(have) == 1 and name == "supply by source"):
                continue
            s = sum(ann.loc[y, p] for p in have)
            gap = t - s
            tol = TOL * (len(have) + 1)
            # unlabelled chart segments of the same review as the parts: measured bar heights
            rev = {src.get((y, p), "").split(" ")[0] for p in have}
            meas = sum(v for (yy, c), (rv, v) in measured.items() if yy == y and c in parts and c not in have
                       and str(rv) in rev)
            if abs(gap) <= tol:
                status = "ok"
            elif meas and abs(gap - meas) <= tol + 0.05:
                status = f"ok - gap {gap:.2f} explained by unlabelled segment(s) measured at {meas:.2f} (not recorded)"
            else:
                status = "FAILED - parts blanked"
                # an older review whose own parts add up to the total: use its parts for this year
                newest = max(int(r) for r in rev if r.isdigit()) if any(r.isdigit() for r in rev) else 9999
                for rv in sorted({int(r) for r in byrev.loc[byrev["Year"] == y, "Review"]}, reverse=True):
                    if rv >= newest:
                        continue
                    sub = byrev[(byrev["Review"] == rv) & (byrev["Year"] == y) & byrev["Column"].isin(parts)]
                    sub = sub.assign(_t=sub["Source"].str.contains("table").astype(int)).sort_values("_t", ascending=False)
                    older = sub.drop_duplicates("Column").set_index("Column")
                    if len(older) >= 2 and abs(t - older["Value"].sum()) <= TOL * (len(older) + 1):
                        for p in have:
                            ann.loc[y, p] = float("nan")
                            src.pop((y, p), None)
                        for p, r in older.iterrows():
                            ann.loc[y, p] = r["Value"]
                            src[(y, p)] = f"{rv} review {r['Source']}"
                        status = (f"ok - parts from the {rv} review (newest review's parts add to {s:.2f}, "
                                  f"{rv} review's to {older['Value'].sum():.2f})")
                        have = list(older.index)
                        s = older["Value"].sum()
                        gap = t - s
                        break
                if status.startswith("FAILED"):
                    blank.append((y, have))
            checks.append({"Year": y, "Check": name, "Total": t, "Sum_of_parts": round(s, 3), "Gap": round(gap, 3),
                           "Parts_present": ", ".join(p.replace("_bcm", "") for p in have), "Status": status})
    for y, have in blank:
        for p in have:
            ann.loc[y, p] = float("nan")
    ann = ann.dropna(axis=1, how="all")
    ann.index.name = "Year"
    ann.insert(0, "date", [pd.Timestamp(year=int(y), month=1, day=1) for y in ann.index])
    srcs = pd.DataFrame([{"Year": y, "Column": c, "From": s} for (y, c), s in sorted(src.items())])
    return ann, pd.DataFrame(checks), srcs


def monthly_rows(monthly, review_year, annual_vals):
    """Validated monthly rows [(month date, column, value)] from one review's monthly charts, and notes."""
    rows, notes = [], []
    for pi, mv, labs in monthly:
        for s, months in mv.items():
            col = supply_series(s) if any(_m(lab, "תמר") for lab in labs) and any(_m(lab, "כריש") for lab in labs) \
                else ("Domestic_Power_bcm" if _m(s, "חשמל") else "Domestic_Cogeneration_bcm" if _m(s, "קוגנרציה")
                      else "Domestic_Industry_bcm" if _m(s, "תעשייה") else "Export_total_bcm" if _m(s, "יצוא")
                      else None)
            if not col:
                continue
            if col in PREV and not current_basis(review_year, review_year):
                col = PREV[col]
            elif col == "Domestic_Cogeneration_bcm" and not current_basis(review_year, review_year):
                continue
            if any(v is None or v == "x" for v in months):
                notes.append(f"p{pi} monthly {col}: {sum(v is None or v == 'x' for v in months)} month(s) without "
                             f"a clear label; not used")
                continue
            ann = annual_vals.get((review_year, col))
            tot = sum(months)
            if ann is None or abs(tot - ann) > 0.15:
                notes.append(f"p{pi} monthly {col}: 12 months add to {tot:.2f} vs annual {ann}; not used")
                continue
            for i, v in enumerate(months):
                rows.append((pd.Timestamp(year=review_year, month=i + 1, day=1), col, v))
            notes.append(f"p{pi} monthly {col}: 12 months = {tot:.2f} vs annual {ann} - used")
    return rows, notes


# ------------------------------------------------------------------ download and workbook

def fetch(url, head=False):
    from curl_cffi import requests as creq
    for attempt in range(3):
        try:
            r = (creq.head if head else creq.get)(url, impersonate="chrome", timeout=120)
            return r
        except Exception as e:  # noqa: BLE001
            log(f"  {url}: {e} (attempt {attempt + 1})")
    return None


def read_sheet(path, name):
    try:
        return pd.read_excel(path, sheet_name=name)
    except Exception:  # noqa: BLE001
        return None


def notes_lines(releases, n_years):
    return [
        "Israel natural gas annual balance - Natural Gas Authority 'Natural Gas Sector Review' (Ministry of Energy "
        "and Infrastructure), Hebrew PDFs on gov.il BlobFolder",
        "",
        "UNITS",
        "All volumes BCM per year (billion cubic metres) as published; Monthly sheet BCM per month. MCM figures "
        "(distribution, older reviews) are divided by 1,000.",
        "",
        "SOURCE AND METHOD",
        "Reviews: https://www.gov.il/BlobFolder/reports/ng_<YYYY>/he/ng_<YYYY>.pdf (2018-2021) and "
        ".../ng-<YYYY>/he/ng-<YYYY>.pdf (2022 on); list on 'Releases'. gov.il pages are Cloudflare-blocked, the "
        "BlobFolder files are not (curl_cffi).",
        "Values come from the PDF text layer only: the tables of the 2018-2022 and 2025 reviews and the DATA LABELS of "
        "the vector bar charts of the 2022-2025 reviews, matched to their bars (each label must agree with its bar's "
        "height on the chart's own scale). Segments without a printed label are blank; bar heights are never used as "
        "data. 2023 review charts and the 2023/2024 summary tables are images: nothing read from them.",
        "Each year/item takes the NEWEST review that publishes it (table before chart); every value of every review "
        "is on 'By review', the review/page used for each Annual cell on 'Sources used'.",
        "",
        "COLUMNS",
        "Supply_*: supply by source (Tamar, Leviathan, Karish, Yam Tethys, EMG pipeline imports from Egypt to 2012, "
        "the offshore LNG buoy); Supply_total = review's total supply (differs from Total = use by up to 0.1).",
        "Domestic_Power / Cogeneration / Industry: CURRENT basis (cogeneration a separate sector; 2025 review for every "
        "year, 2024 review for 2023-24). Industry = connected to transmission + distribution + CNG.",
        "*_prev_basis: reviews before the split (and the 2024 review for 2020-22) put cogeneration partly in power and "
        "partly in industry; never mixed with the current-basis columns. The 2019 review's 2019 industry (2.44) "
        "includes the 0.22 exported to Jordan; later reviews give 2.22.",
        "Power_IEC / Power_private: gas for power by Israel Electric Corp. and private producers (IPPs); "
        "Power_private_incl_cogen_prev_basis = 'IPPs and cogeneration' of older reviews.",
        "Export_Egypt / Export_Jordan: by destination metered in that country (gas to Egypt via the Jordan North line "
        "counts as Egypt). Total = domestic + export.",
        "",
        "CHECKS",
        "Per year (sheet 'Checks'): fields = supply total, domestic + export = total, sectors = domestic, industry parts "
        "= industry, Egypt + Jordan = export, IEC + private = power; tolerance 0.06 BCM per summed item. A failed check "
        "blanks the parts for that year. A gap that the measured height of unlabelled chart segments explains passes "
        "(the unlabelled segment itself stays blank).",
        "",
        "SCHEDULE",
        "Owner's choice: checked every 5 months (1 Jan, 1 Jun, 1 Nov); the review appears once a year (2025 review "
        "released 3 Jun 2026). A run exits without "
        f"rewriting when no new or changed review exists. {len(releases)} reviews read, {n_years} years.",
        "Script: middle_east/ISRAEL_NGA_ANNUAL.py; probe: discovery_archive/middle_east/ISRAEL_NGA_REVIEW_PROBE.py.",
    ]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--pdf-dir", help="read local review PDFs (ng-YYYY.pdf / ng_YYYY.pdf) instead of downloading")
    ap.add_argument("--reparse", action="store_true", help="parse every review again")
    args = ap.parse_args()

    rel = read_sheet(args.out, "Releases")
    byrev = read_sheet(args.out, "By review")
    mon_old = read_sheet(args.out, "Monthly long")
    known = {} if rel is None else {int(r.Review): r._asdict() for r in rel.itertuples(index=False)}
    this_year = datetime.now(timezone.utc).year
    todo = {}   # review year -> (url, content, last_modified)

    if args.pdf_dir:
        for fn in sorted(os.listdir(args.pdf_dir)):
            m = re.search(r"ng[-_](20\d\d)\.pdf$", fn)
            if m:
                todo[int(m.group(1))] = ("file://" + fn, open(os.path.join(args.pdf_dir, fn), "rb").read(), "")
    else:
        for y in range(FIRST_YEAR, this_year + 1):
            k = known.get(y)
            if k and not args.reparse:
                r = fetch(k["URL"], head=True)
                lm = r.headers.get("Last-Modified", "") if r is not None and r.status_code == 200 else None
                if lm is None or lm == (k.get("Last_Modified") or ""):
                    continue
                log(f"{y}: Last-Modified changed ({k.get('Last_Modified')} -> {lm})")
            for pat in URL_PATTERNS:
                url = pat.format(y=y)
                r = fetch(url)
                if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
                    todo[y] = (url, r.content, r.headers.get("Last-Modified", ""))
                    log(f"{y}: {url} ({len(r.content):,} bytes)")
                    break
    if not todo and byrev is not None:
        log("No new or changed review; nothing to do.")
        return 0

    recs = [] if byrev is None else [r for r in byrev.to_dict("records") if int(r["Review"]) not in todo]
    mon = [] if mon_old is None else [r for r in mon_old.to_dict("records") if int(r["Review"]) not in todo]
    rel_rows = {y: k for y, k in known.items() if y not in todo}
    monthly_by_rev = {}
    for y, (url, content, lm) in sorted(todo.items()):
        r_recs, monthly, notes = parse_review(content, y)
        log(f"review {y}: {len(r_recs)} values, {len(monthly)} monthly chart(s)")
        for n in notes:
            log(f"   {n}")
        for d, c, v, s in r_recs:
            recs.append({"Review": y, "Year": int(d), "Column": c, "Value": v, "Source": s})
        monthly_by_rev[y] = (monthly, notes)
        rel_rows[y] = {"Review": y, "URL": url, "Last_Modified": lm, "Bytes": len(content),
                       "SHA256": hashlib.sha256(content).hexdigest()[:16],
                       "Read_utc": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M"),
                       "Values": len(r_recs), "Notes": ""}
    byrev = pd.DataFrame(recs, columns=["Review", "Year", "Column", "Value", "Source"])
    byrev = byrev.drop_duplicates(["Review", "Year", "Column", "Source"])
    vals, _, _ = pick(byrev)
    for y, (monthly, notes) in monthly_by_rev.items():
        rows, mnotes = monthly_rows(monthly, y, vals)
        mon += [{"Review": y, "date": d, "Column": c, "Value": v} for d, c, v in rows]
        chk = [n for n in notes if "FAILED" in n or "does not add up" in n]
        rel_rows[y]["Notes"] = "; ".join(mnotes + chk)[:2000]
    ann, checks, srcs = build_annual(byrev)
    releases = pd.DataFrame([rel_rows[y] for y in sorted(rel_rows)])
    sheets = {"Annual": ann}
    if mon:
        ml = pd.DataFrame(mon)
        wide = ml.pivot_table(index="date", columns="Column", values="Value", aggfunc="first")
        wide = wide[[c for c in COLUMNS if c in wide.columns]]
        sheets["Monthly"] = wide
        sheets["Monthly long"] = ml.set_index("Review")
    sheets["Checks"] = checks.set_index("Year") if len(checks) else pd.DataFrame({"Year": []}).set_index("Year")
    sheets["Sources used"] = srcs.set_index("Year")
    sheets["By review"] = byrev.sort_values(["Review", "Year", "Column"]).set_index("Review")
    sheets["Releases"] = releases.set_index("Review")
    lines = notes_lines(releases, len(ann))
    titles = [l for l in lines if l.isupper() and l]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, lines, titles + [lines[0]])
    log(f"Saved {args.out}: {len(ann)} years ({ann.index.min()}-{ann.index.max()}), "
        f"{int((checks['Status'].str.startswith('FAILED')).sum()) if len(checks) else 0} failed checks")
    gh = os.environ.get("GITHUB_OUTPUT")
    if gh:
        with open(gh, "a") as f:
            f.write("updated=1\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
