"""
Parser for the Energy Commission of Ghana's weekly Wholesale Electricity Market (WEM) statistics PDFs: the table
'Power Plant generation (GWh)' (one row per plant, one column per day of the week, plus a Total column and a Total row).

Works on word boxes (pymupdf page.get_text('words'): x0, y0, x1, y1, text), so it can be tested without the PDFs.
If the table is rotated 90 degrees on a page (plants across, days down), the coordinates are swapped and the same
code is used.
"""
import re
from collections import defaultdict

import pandas as pd

DATE_RE = re.compile(r"^(\d{1,2})-([A-Za-z]{3})-(\d{2})$")
NUM_RE = re.compile(r"^-?\d+(?:\.\d+)?$")
MONTHS = {m: i + 1 for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}


def to_date(tok):
    m = DATE_RE.match(tok)
    if not m or m.group(2).lower() not in MONTHS:
        return None
    try:
        return pd.Timestamp(2000 + int(m.group(3)), MONTHS[m.group(2).lower()], int(m.group(1)))
    except ValueError:
        return None


def _cx(w):
    return (w[0] + w[2]) / 2


def _cy(w):
    return (w[1] + w[3]) / 2


def _parse_oriented(words):
    """words in (x across, y down) orientation; the day header row has >= 5 dates and a 'Total' word on its line."""
    dates = [w for w in words if DATE_RE.match(w[4]) and to_date(w[4]) is not None]
    rows = defaultdict(list)
    for w in dates:
        rows[round(_cy(w) / 4)].append(w)
    for key in sorted(rows):
        hdr = rows[key]
        if len(hdr) < 5:
            continue
        hy = sum(_cy(w) for w in hdr) / len(hdr)
        tot = [w for w in words if w[4] == "Total" and abs(_cy(w) - hy) < 5 and _cx(w) > max(_cx(h) for h in hdr)]
        if not tot:
            continue
        return hdr, hy, tot[0]
    return None


def parse_words(words):
    """-> (DataFrame plants x dates in GWh, Series day totals from the table's Total row or None, warnings list).
    Raises ValueError when no generation table is found."""
    orient = _parse_oriented(words)
    flip = False
    if orient is None:   # rotated table: swap axes (x <-> y)
        sw = [[w[1], w[0], w[3], w[2], w[4]] for w in words]
        orient = _parse_oriented(sw)
        if orient is None:
            raise ValueError("no generation table")
        words, flip = sw, True
    hdr, hy, totw = orient
    hdr = sorted(hdr, key=_cx)
    cols = [(_cx(w), to_date(w[4])) for w in hdr]
    left = min(w[0] for w in hdr)
    pitch = min(b[0] - a[0] for a, b in zip(cols, cols[1:])) if len(cols) > 1 else 40
    total_x = _cx(totw)
    below = [w for w in words if _cy(w) > hy + 4]
    lines = defaultdict(list)
    for w in sorted(below, key=_cy):
        # cluster by vertical centre (rows are >= 10 pt apart)
        for k in lines:
            if abs(k - _cy(w)) <= 4:
                lines[k].append(w)
                break
        else:
            lines[_cy(w)].append(w)
    data, totals, warns = {}, None, []
    for k in sorted(lines):
        ws = sorted(lines[k], key=_cx)
        label_ws = [w for w in ws if w[2] < left - 6 and not NUM_RE.match(w[4])]
        nums = [w for w in ws if NUM_RE.match(w[4]) and w[0] >= left - pitch / 2]
        label = " ".join(w[4] for w in label_ws).strip()
        if not label and not nums:
            continue
        if not label or len(nums) < len(cols):
            if data and label.lower().startswith(("disclaimer", "the ")):   # past the table (footer text)
                break
            continue
        # values are matched to the day headers by ORDER along the row (some issues print the header row with a
        # different column spacing from the data rows): n days, optionally followed by the row total
        nums = sorted(nums, key=_cx)
        n = len(cols)
        if len(nums) not in (n, n + 1):
            warns.append(f"row '{label}' has {len(nums)} numbers for {n} days")
            continue
        vals = {cols[j][1]: float(nums[j][4]) for j in range(n)}
        if len(nums) == n + 1:
            vals["Total"] = float(nums[n][4])
        if label.lower() == "total":
            totals = pd.Series({d: vals.get(d) for _, d in cols}, dtype=float)
            break
        data[label] = {d: vals.get(d) for _, d in cols}
    if not data:
        raise ValueError("table header found but no plant rows")
    df = pd.DataFrame(data).T
    df.columns = pd.DatetimeIndex(df.columns)
    return df, totals, warns   # columns in header (left-to-right) order
