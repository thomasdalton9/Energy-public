"""
Israel natural gas sales by field, QUARTERLY, from the field owners' own reports (no company publishes monthly volumes):

  Leviathan (100%) - NewMed Energy financial statements (English 'ACCESSIBLE' PDFs listed on
      newmedenergy.com/investor-relations/financial-reports/): table "gas quantities (100%) sold from the Leviathan
      reservoir ... according to the customers' geographic location", Israel | Jordan | Egypt | Total, BCM rounded to 0.1,
      one row per quarter for the report year and the year before (from Q1 2020).
  Tamar (100%) - Tamar Petroleum reports (every file on its site is listed by the keyless WordPress media API
      tamarpetroleum.co.il/wp-json/wp/v2/media): table "broken down by domestic market and export, in each of the quarters"
      (BCM, 2 decimals, Q1/2024 on). Export is one figure (Egypt and Jordan together). The site holds no English report
      with this table before 2024, so Tamar - and the national total - start in Q1 2024.
  Karish (+ Karish North, Katlan; Energean 100%) - Energean plc regulatory announcements (RNS) as republished on
      investegate.co.uk (energean.com blocks GitHub): 'Israel average W.I. production (kboed)' rows with "(inc. X bcm of
      gas)" for year-to-date periods (three, six, nine and twelve months). Quarters are DIFFERENCES of those year-to-date
      figures (Q2 = H1 - Q1 ...); a quarter whose two year-to-date figures are not both published stays blank.

All volumes are SALES (gross field, 100%), which differ slightly from production. The national quarterly total is the sum
of the three fields and is only given for quarters all three have. Bcf/d = BCM x 35.3147 / days in the quarter.
Check (Oct 2026 probe): 2025 Leviathan 10.9 + Tamar 10.05 + Karish 5.6 = 26.6 BCM vs the Natural Gas Authority's 26.8.

SCHEDULE: quarterly (15 Jan/Apr/Jul/Oct): the owners publish 6-10 weeks after each quarter, so each run picks up the
quarter before. Incremental: every report read is listed on the 'Reports' sheet and not downloaded again; a later report
replaces the figures of the same quarter (revisions, rounding).
Probe: discovery_archive/middle_east/ISRAEL_GAS_COMPANY_REPORTS_PROBE.py. Workflow: .github/workflows/israel_gas_fields.yml.
"""

import argparse
import io
import os
import re
import sys
from datetime import date, datetime, timezone

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept-Language": "en-US,en;q=0.9"}
NEWMED_LIST = "https://newmedenergy.com/investor-relations/financial-reports/"
TAMAR_MEDIA = "https://www.tamarpetroleum.co.il/wp-json/wp/v2/media?per_page=100&page={}"
ENOG_LIST = "https://www.investegate.co.uk/company/ENOG?page={}"
ENOG_PAGES = 12          # announcement list pages read per run (newest first; ~20 a page)
FIRST_YEAR = 2020
BCF_PER_BCM = 35.3147

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "israel_gas_fields_quarterly.xlsx")


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def get(url, **kw):
    r = requests.get(url, headers=UA, timeout=90, **kw)
    r.raise_for_status()
    return r


def pdf_text(content):
    from pypdf import PdfReader
    return "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(content)).pages)


def qdate(year, q):
    return pd.Timestamp(year=int(year), month=3 * (int(q) - 1) + 1, day=1)


# ------------------------------------------------------------------ Leviathan (NewMed)

def newmed_reports():
    html = get(NEWMED_LIST).text
    urls = set(re.findall(r'href=["\'](https?://[^"\']*newmedenergy\.com/wp-content/uploads/(20\d\d)/\d\d/[^"\']+\.pdf)', html, re.I))
    out = [u for u, y in urls if int(y) >= FIRST_YEAR + 1 and re.search(r"FS", u.rsplit("/", 1)[-1], re.I)
           and not re.search(r"bond|heb|[֐-׿]", u, re.I)]
    return sorted(out)


ROW_RE = re.compile(r"^\s*Q\s*([1-4])\s+(\d+\.\d)\s+(\d+\.\d)\s+(\d+\.\d)\s+(\d+\.\d)\b")
TOT_RE = re.compile(r"^\s*(?:Total|annual price)[^\d]*(\d+\.\d)\s+(\d+\.\d)\s+(\d+\.\d)\s+(\d+\.\d)\b")


def parse_leviathan(text):
    """[(quarter start, israel, jordan, egypt, total)] from the 'geographic location' table(s): a year heading
    ('2025***', '(BCM)* -2022', '2023 - (BCM)*'), quarter rows, then a Total row (split over two lines in annual
    reports: 'Total/Average' / 'annual price 1.8 2.9 6.2 10.9 5.58') that the quarters must add up to."""
    starts = [m for m in re.finditer(r"geographic\s*location", text) if "Leviathan" in text[max(0, m.start() - 500): m.start()]]
    if not starts:
        return []
    lines = text[starts[0].end(): starts[0].end() + 4000].splitlines()
    rows, year, block = [], None, []

    def flush(total=None):
        nonlocal rows, block
        if block and total is not None and abs(sum(b[4] for b in block) - total) > 0.25:
            raise RuntimeError(f"Leviathan {block[0][0].year}: quarters {sum(b[4] for b in block):.1f} != total {total}")
        rows += block
        block = []

    for line in lines:
        if re.search(r"Cost of natural gas|production\s+costs|Depreciation", line):
            break
        r = ROW_RE.match(line)
        t = TOT_RE.match(line)
        y = re.search(r"\b(20[12]\d)\b", line)
        if r and year:
            block.append((qdate(year, r.group(1)), *(float(r.group(i)) for i in range(2, 6))))
        elif t and block:
            flush(float(t.group(4)))
            year = None
        elif y and not r and not re.search(r"Q[1-4]/|approx|million|\$\s*\d", line):
            flush()
            year = int(y.group(1))
    flush()
    return rows


# ------------------------------------------------------------------ Tamar (Tamar Petroleum)

def tamar_reports():
    out = []
    for page in range(1, 20):
        r = requests.get(TAMAR_MEDIA.format(page), headers=UA, timeout=90)
        if r.status_code != 200:
            break
        items = r.json()
        for it in items:
            u = it.get("source_url") or ""
            name = u.rsplit("/", 1)[-1]
            if u.lower().endswith(".pdf") and re.search(r"FS|Financial", name, re.I) \
                    and not re.search(r"[֐-׿]|immediate|reserve|NSAI", name, re.I):
                y = re.search(r"/uploads/(20\d\d)/", u)
                if y and int(y.group(1)) >= 2024:
                    out.append(u)
        if len(items) < 100:
            break
    return sorted(set(out))


TAMAR_ROW = re.compile(r"Q([1-4])/(20\d\d)\s+(\d+\.\d\d)\s+(\d+\.\d\d)\s+(\d+\.\d\d)")


def parse_tamar(text):
    """[(quarter start, domestic, export, total)] from the 'domestic market and export' quantity table(s)."""
    t = re.sub(r"\s+", " ", text)
    rows = []
    for m in re.finditer(r"broken down by domestic market and export, in each of the quarters[^(]{0,120}\(in BCM\)", t):
        seg = t[m.end(): m.end() + 900]
        seg = seg.split("revenues")[0]
        for r in TAMAR_ROW.finditer(seg):
            dom, exp, tot = float(r.group(3)), float(r.group(4)), float(r.group(5))
            if abs(dom + exp - tot) > 0.03:
                raise RuntimeError(f"Tamar Q{r.group(1)}/{r.group(2)}: {dom} + {exp} != {tot}")
            rows.append((qdate(r.group(2), r.group(1)), dom, exp, tot))
    return rows


# ------------------------------------------------------------------ Karish (Energean RNS)

ENOG_KEEP = re.compile(r"trading-statement|operational-update|results|half-year|full-year|interim|preliminary|"
                       r"production-and|9-months|nine-months", re.I)


def energean_announcements(seen):
    out = []
    for page in range(1, ENOG_PAGES + 1):
        r = requests.get(ENOG_LIST.format(page), headers=UA, timeout=90)
        if r.status_code != 200:
            break
        links = re.findall(r'href=["\']((?:https://www\.investegate\.co\.uk)?/announcement/rns/energean--enog/([^"\']+))["\']', r.text)
        if not links:
            break
        for href, slug in links:
            url = href if href.startswith("http") else "https://www.investegate.co.uk" + href
            if ENOG_KEEP.search(slug) and url not in seen:
                out.append(url)
    return sorted(set(out))


PERIODS = [  # (regex giving the year, months covered)
    (r"Three[- ]months? (?:to|ended) 31 March,? (20\d\d)|Q1 (20\d\d)", 3),
    (r"Six[- ]months? (?:to|ended) 30 June,? (20\d\d)|H1 (20\d\d)|Half[- ]year (?:to 30 June )?(20\d\d)", 6),
    (r"Nine[- ]months? (?:to|ended) 30 September,? (20\d\d)|9M (20\d\d)", 9),
    (r"(?:FY|Full[- ]year|Year (?:to|ended) 31 December,?) (20\d\d)", 12),
]
_INC = r"(?: \(inc\. (\d+(?:\.\d+)?) bcm of (?:sales )?gas\))?"
ISRAEL_ROW = re.compile(r"Israel(?: average W\.I\. production)?(?: \(kboed\))?(?: \(\d\))? \| (\d+(?:\.\d+)?)" + _INC
                        + r" \| (\d+(?:\.\d+)?)" + _INC + r" \|")
# 'Q1 2026: ... 0.96 bcm of gas in Q1 2026 and 1.19 bcm in Q1 2025'
PERIOD_BCM = re.compile(r"(\d+\.\d+) bcm(?: of (?:sales )?gas)? in (Q1|H1|9M|FY) (20\d\d)")
# Energean Israel accounts: '... for six months ended 30 June 2025 totaled approximately 2.3 bcm (billion cubic metres)
# and for six months ended 30 June 2024 totaled approximately 2.5 bcm'; '... between 26 October 2022 and 31 December
# 2022 totaled approximately 0.x'
ACCOUNTS_BCM = re.compile(r"(six months ended 30 June|nine months ended 30 September|three months ended 31 March|"
                          r"year ended 31 December|between 26 October 2022 and 31 December)?,?\s*(20\d\d)\)?,?\s+"
                          r"totaled approximately\s+(\d+(?:\.\d+)?)\s*(?:bcm)?", re.I)


def rns_text(html):
    t = re.sub(r"<(?:td|th)[^>]*>", " | ", html)
    t = re.sub(r"<[^>]+>", " ", t)
    t = re.sub(r"&nbsp;|&#160;", " ", t).replace("&amp;", "&").replace("&#8209;", "-").replace("&#8211;", "-")
    return re.sub(r"\s+", " ", t)


def parse_karish(text):
    """[(year, months, kboed, bcm or None)] year-to-date figures, current period and the one a year earlier."""
    out = []
    for m in ISRAEL_ROW.finditer(text):
        before = text[max(0, m.start() - 2500): m.start()]
        if not re.search(r"kboed|production", before[-400:] + m.group(0), re.I):
            continue   # an Israel row of a cost / capex table, not production
        heads = []
        window = before[-600:]
        for pat, months in PERIODS:
            for h in re.finditer(pat, window):
                if re.match(r"\s*(?:Kboed\s*)?guidance", window[h.end(): h.end() + 25], re.I):
                    continue   # 'FY 2023 guidance' column, not a reported period
                heads.append((h.start(), int(next(g for g in h.groups() if g)), months))
        if not heads:
            continue
        # the row's first value belongs to the table's first period heading ('current | year before | ...')
        pos, _, months = min(heads)
        yr = max(y for p, y, mo in heads if mo == months and p - pos < 300)
        out.append((yr, months, float(m.group(1)), float(m.group(2)) if m.group(2) else None))
        out.append((yr - 1, months, float(m.group(3)), float(m.group(4)) if m.group(4) else None))
    for s in PERIOD_BCM.finditer(text):
        out.append((int(s.group(3)), {"Q1": 3, "H1": 6, "9M": 9, "FY": 12}[s.group(2)], None, float(s.group(1))))
    if re.search(r"natural gas sold|quantit(?:y|ies) of (?:natural )?gas", text, re.I):
        for s in ACCOUNTS_BCM.finditer(text):
            if "bcm" not in text[s.end() - 4: s.end() + 30] and "billion cubic" not in text[s.end(): s.end() + 40]:
                continue
            lead = (s.group(1) or "").lower()
            months = 6 if "six" in lead else 9 if "nine" in lead else 3 if "three" in lead else 12
            out.append((int(s.group(2)), months, None, float(s.group(3))))
    return out


def karish_quarters(ytd):
    """Quarterly BCM from year-to-date BCM: Q1 = 3M, Q2 = 6M - 3M, Q3 = 9M - 6M, Q4 = 12M - 9M."""
    rows = []
    ytd = dict(ytd)
    ytd.setdefault((2022, 9), 0.0)   # Karish first gas 26 Oct 2022: nothing sold before Q4 2022
    for (yr, months), v in sorted(ytd.items()):
        if yr < 2022:
            continue
        if v is None or pd.isna(v):
            continue
        prev = 0.0 if months == 3 else ytd.get((yr, months - 3))
        if prev is None or pd.isna(prev):
            continue
        q = round(v - prev, 3)
        if q < 0:
            raise RuntimeError(f"Karish {yr} {months}M: year-to-date {v} below the previous {prev}")
        rows.append((qdate(yr, months // 3), q))
    return rows


# ------------------------------------------------------------------ archive

def read_sheet(path, sheet, cols):
    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame(columns=cols)
    return df.drop(columns=[c for c in df.columns if str(c).startswith("Unnamed")])


NOTES_LINES = [
    "UNITS",
    "BCM per quarter (billion cubic metres) of natural gas SOLD from each field, 100% (gross) of the field, as the owners "
    "report it. 'Bcf/d' columns = BCM x 35.3147 / days in the quarter. Sales differ slightly from production.",
    "",
    "FIELDS AND SOURCES",
    "Leviathan: NewMed Energy financial statements (table of quantities sold by customers' location: Israel, Jordan, "
    "Egypt; rounded to 0.1 BCM; Egypt-bound gas that crosses Jordan counts as Egypt). Tamar: Tamar Petroleum reports "
    "(domestic market and export, 2 decimals; export not split by country). Karish (incl. Karish North and Katlan): "
    "Energean plc announcements via investegate.co.uk: year-to-date gas in BCM; each quarter is the difference of two "
    "year-to-date figures, so it carries their rounding (0.1 BCM) and is blank if either is not published.",
    "",
    "COVERAGE",
    "Leviathan from Q1 2020; Tamar from Q1 2024 (the owner's site has no English report with the quarterly table "
    "earlier); Karish from the announcements found. The national total is given only for quarters all three fields "
    "have. Reported 6-10 weeks after each quarter; the pull runs on 15 Jan/Apr/Jul/Oct.",
    "",
    "CHECK",
    "Natural Gas Authority (gov.il) annual review: 2025 total 26.8 BCM (Leviathan 11, Tamar 10, Karish 5.8), exports "
    "12.4 (Egypt 8.9, Jordan 3.5). Sum of the company figures for 2025: 10.9 + 10.05 + 5.6 = 26.6 BCM.",
    "",
    "SOURCE LINKS",
    f"NewMed: {NEWMED_LIST} | Tamar Petroleum: https://www.tamarpetroleum.co.il/ | Energean via "
    "https://www.investegate.co.uk/company/ENOG",
]
NOTES_TITLES = {"UNITS", "FIELDS AND SOURCES", "COVERAGE", "CHECK", "SOURCE LINKS"}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()

    lev = read_sheet(args.out, "Leviathan", ["date", "Israel_bcm", "Jordan_bcm", "Egypt_bcm", "Total_bcm", "Source"])
    tam = read_sheet(args.out, "Tamar", ["date", "Domestic_bcm", "Export_bcm", "Total_bcm", "Source"])
    kar = read_sheet(args.out, "Karish YTD", ["Year", "Months", "Israel_kboed", "Gas_bcm", "Source"])
    reps = read_sheet(args.out, "Reports", ["Source", "Field", "Read_utc", "Rows"])
    reps = reps[reps["Field"] != "Karish"] if len(reps) else reps   # Karish announcements are re-read every run
    seen = set(reps["Source"])
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M")
    new_reports, failures = [], []

    def upsert(df, rows, cols, src):
        add = pd.DataFrame(rows, columns=cols).assign(Source=src)
        add["date"] = pd.to_datetime(add["date"])
        if not len(df):
            return add
        df["date"] = pd.to_datetime(df["date"])
        return pd.concat([df[~df["date"].isin(add["date"])], add], ignore_index=True)

    # Leviathan: oldest report first, so a later report's figures replace an earlier one's
    for url in sorted(newmed_reports(), key=lambda u: re.search(r"/uploads/(\d{4}/\d\d)/", u).group(1)):
        if url in seen:
            continue
        try:
            rows = parse_leviathan(pdf_text(get(url).content))
        except Exception as e:  # noqa: BLE001
            failures.append(f"Leviathan {url}: {type(e).__name__}: {e}")
            continue
        if rows:
            lev = upsert(lev, rows, ["date", "Israel_bcm", "Jordan_bcm", "Egypt_bcm", "Total_bcm"], url)
        new_reports.append((url, "Leviathan", now, len(rows)))
        log(f"  Leviathan {url.rsplit('/', 1)[-1]}: {len(rows)} quarter row(s)")

    for url in sorted(tamar_reports(), key=lambda u: re.search(r"/uploads/(\d{4}/\d\d)/", u).group(1)):
        if url in seen:
            continue
        try:
            rows = parse_tamar(pdf_text(get(url).content))
        except Exception as e:  # noqa: BLE001
            failures.append(f"Tamar {url}: {type(e).__name__}: {e}")
            continue
        if rows:
            tam = upsert(tam, rows, ["date", "Domestic_bcm", "Export_bcm", "Total_bcm"], url)
        new_reports.append((url, "Tamar", now, len(rows)))
        log(f"  Tamar {url.rsplit('/', 1)[-1]}: {len(rows)} quarter row(s)")

    for url in energean_announcements(set()):
        try:
            found = parse_karish(rns_text(get(url).text))
        except Exception as e:  # noqa: BLE001
            failures.append(f"Karish {url}: {type(e).__name__}: {e}")
            continue
        for yr, months, kb, bcm in found:
            mask = (kar["Year"] == yr) & (kar["Months"] == months) if len(kar) else None
            if mask is not None and mask.any():
                i = kar.index[mask][0]
                if bcm is not None:
                    old = kar.loc[i, "Gas_bcm"]
                    if pd.notna(old) and abs(old - bcm) > 0.15:
                        log(f"  Karish {yr} {months}M: {old} -> {bcm} bcm ({url})")
                    if pd.isna(old) or len(str(bcm)) > len(str(old)):   # keep the more precise figure
                        kar.loc[i, ["Gas_bcm", "Source"]] = [bcm, url]
                if kb is not None and pd.isna(kar.loc[i, "Israel_kboed"]):
                    kar.loc[i, "Israel_kboed"] = kb
            else:
                kar = pd.concat([kar, pd.DataFrame([{"Year": yr, "Months": months, "Israel_kboed": kb, "Gas_bcm": bcm,
                                                     "Source": url}])], ignore_index=True)
        if url not in seen:
            new_reports.append((url, "Karish", now, len(found)))
        log(f"  Karish {url.rsplit('/', 2)[-2]}: {[(y, m, k, b) for y, m, k, b in found]}")

    if new_reports:
        reps = pd.concat([reps, pd.DataFrame(new_reports, columns=["Source", "Field", "Read_utc", "Rows"])],
                         ignore_index=True)
    kar = kar.sort_values(["Year", "Months"]).reset_index(drop=True)
    ytd = {(int(r.Year), int(r.Months)): r.Gas_bcm for r in kar.itertuples()}
    kq = pd.DataFrame(karish_quarters(ytd), columns=["date", "Karish_bcm"]).set_index("date")

    q = pd.DataFrame(index=pd.DatetimeIndex([], name="date"))
    if len(lev):
        lv = lev.set_index(pd.to_datetime(lev["date"])).sort_index()
        q = q.join(lv[["Total_bcm", "Israel_bcm", "Jordan_bcm", "Egypt_bcm"]].add_prefix("Leviathan_"), how="outer")
    if len(tam):
        tm = tam.set_index(pd.to_datetime(tam["date"])).sort_index()
        q = q.join(tm[["Total_bcm", "Domestic_bcm", "Export_bcm"]].add_prefix("Tamar_"), how="outer")
    q = q.join(kq, how="outer")
    q = q.rename(columns={"Leviathan_Total_bcm": "Leviathan_bcm", "Tamar_Total_bcm": "Tamar_bcm"})
    for c in ["Leviathan_bcm", "Tamar_bcm", "Karish_bcm"]:
        if c not in q:
            q[c] = float("nan")
    q["Total_bcm"] = q[["Leviathan_bcm", "Tamar_bcm", "Karish_bcm"]].sum(axis=1, min_count=3)
    days = (q.index + pd.offsets.QuarterBegin(startingMonth=1) - q.index).days
    for c in ["Leviathan_bcm", "Tamar_bcm", "Karish_bcm", "Total_bcm"]:
        q[c.replace("_bcm", "_Bcf_d")] = (q[c] * BCF_PER_BCM / days).round(3)
    q.index.name = "date"

    for f in failures:
        log("FAILED: " + f)
    xlsx_notes.write_workbook(args.out, {"Quarterly": q.sort_index(),
                                         "Leviathan": lev.sort_values("date").set_index("date"),
                                         "Tamar": tam.sort_values("date").set_index("date"),
                                         "Karish YTD": kar.set_index("Year"),
                                         "Reports": reps.set_index("Source")}, NOTES_LINES, NOTES_TITLES)
    full = q.dropna(subset=["Total_bcm"])
    print(f"Saved {args.out}: {len(new_reports)} new report(s); quarters with all three fields: "
          f"{full.index.min():%Y-%m} to {full.index.max():%Y-%m}" if len(full) else
          f"Saved {args.out}: {len(new_reports)} new report(s); no quarter has all three fields yet")
    if failures:
        sys.exit(f"{len(failures)} report(s) could not be parsed (listed above); everything else is saved")


if __name__ == "__main__":
    main()
