"""
ONE-OFF backfill (Oct 2026): Tamar gas field quarterly SALES history, 100% of the field, BCM, Q1 2016 - Q4 2023, written
into the 'Tamar' sheet of israel_gas_fields_quarterly.xlsx (the store of middle_east/ISRAEL_GAS_FIELDS.py, whose regular
pull only reads Tamar Petroleum's English reports from 2024 on). Rows from Q1 2024 on are never touched.

SOURCES (all Tamar Petroleum Ltd, partner in Tamar with 16.75%; every figure is for 100% of the field):
  * Tamar Petroleum's own site (keyless WordPress media list wp-json/wp/v2/media): English quarterly / annual reports
    Q2 2017 - Q2 2020 and the English periodic report for 2021, Hebrew quarterly reports Q3 2020 - Q3 2021.
  * TASE MAYA (company 1689 'תמר פטרוליום', found through /api/v1/reports/1575974 -> reporterId; the report list
    comes from /api/v1/reports/<id>/siblings in a Playwright page, see
    discovery_archive/middle_east/ISRAEL_TAMAR_HISTORY_PROBE.py): the Hebrew reports for Q1 2022 - FY 2024 that the
    company's site does not hold. The PDFs download with plain requests from mayafiles.tase.co.il/rpdf/..., so this
    script needs no browser; the report ids are listed in SOURCES below.

WHAT THE REPORTS GIVE
  * Every report's Board of Directors table row 'Gas sales in BCM' / 'מכירת גז ב- BCM': year-to-date (current and
    the year before), the quarter itself (current and the year before) and the previous full year; rounded to 0.1 BCM
    up to 2022, 0.01 from 2023. Annual reports give the year and the year before.
  * From Q3 2023 a quarterly table 'רבעון N YYYY  domestic  export  total' (BCM, 2 decimals): Q1 2022 - Q4 2023.
  * Domestic / export splits in the notes as prose: H1 and 9M 2020/2021 (0.1 BCM), Q1 2021 / Q1 2022, H1 and 9M 2022
    (0.01), the year 2020 and 2021 from the English 2021 periodic report (domestic, Jordan, Egypt; 0.01) and 2022 from the
    FY 2024 report's annual table. Before 2020 there is no domestic/export split in any report (Tamar's only export
    then was ~0.15 BCM a year to two plants in Jordan), and Q1/Q2 2020 have none either: those splits stay blank.

METHOD
  Each figure is read from the downloaded PDF text: the table rows by a regex (mapped to periods by the report's own
  period), the prose figures by a regex that must find the quoted numbers in that report (the run stops if any is
  missing). For each quarter and each of total / domestic / export the most precise candidate wins (more decimals;
  on a tie the quarter's own published figure, then the later report):
    - the quarter's own figure (table or prose),
    - a difference of two published year-to-date figures (Q2 = H1 - Q1, Q3 = 9M - H1, Q4 = FY - 9M; Q1 = H1 - Q2 when
      no Q1 figure exists), as in the Karish logic of ISRAEL_GAS_FIELDS.py.
  Where both splits are more precise than the total, total = domestic + export. Nothing is interpolated: a quarter or
  split with no published figures stays absent. Source = the report URL (two URLs, 'A minus B', for a difference).

CHECKS (logged; a failed one stops the run before anything is written)
  * domestic + export within 0.02 BCM of the total (0.10 where a figure is rounded to 0.1, as in 2020-21), else the
    split is left blank;
  * the four quarters of each year within 0.15 BCM of the company's own annual total;
  * annual totals against the Natural Gas Authority's annual reviews (gov.il: 2020 8.2, 2022 10.1, 2023 9.1, 2024
    10.1 BCM); a gap above 0.3 BCM is logged.
Result (Oct 2026): 2016 9.3, 2017 9.6 (company 9.7), 2018 10.4 (10.3), 2019 10.4, 2020 8.25, 2021 8.66, 2022 10.25,
2023 9.12 BCM; gov.il gaps 0.05 / 0.15 / 0.02.

Run once (12 Oct 2026) by discovery_archive/workflows/israel_tamar_history.yml, then ISRAEL_GAS_FIELDS.py rebuilds 'Quarterly'.
"""

import argparse
import io
import logging
import os
import re
import sys
import time
from datetime import datetime, timezone

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)

import xlsx_notes                                  # noqa: E402
import ISRAEL_GAS_FIELDS as igf                    # noqa: E402

logging.getLogger("pypdf").setLevel(logging.ERROR)
UA = igf.UA
SITE = "https://www.tamarpetroleum.co.il/wp-content/uploads/"


def maya(rid):
    lo = (rid - 1) // 1000 * 1000 + 1
    return f"https://mayafiles.tase.co.il/rpdf/{lo}-{lo + 999}/P{rid}-00.pdf"


# (url, report year, months covered, prose facts). Prose fact = (regex that must match the report text,
# [(period, kind, value), ...]); period = ("Q", year, quarter) or ("YTD", year, months).
SOURCES = [
    (SITE + "2020/07/Tamar-Financial-Report-q2-2017-Eng.pdf", 2017, 6, []),
    (SITE + "2020/07/2017_Delek_Q3_En.pdf", 2017, 9, []),          # Tamar Petroleum's Q3 2017 report despite the name
    (SITE + "2020/07/Annual_Report_2017.pdf", 2017, 12, []),
    (SITE + "2020/07/Tamar-Financial-Report-q1-2018.pdf", 2018, 3, []),
    (SITE + "2020/07/Financial_Statements_Q2_2018.pdf", 2018, 6, []),
    (SITE + "2020/07/Financial_Statements_Q3_2018.pdf", 2018, 9, []),
    (SITE + "2020/07/Annual_Report_2018.pdf", 2018, 12, []),
    (SITE + "2020/07/Quarterly_Report_Q1_-_2019.pdf", 2019, 3, []),
    (SITE + "2020/07/Interim_Report_Q2_-_2019.pdf", 2019, 6, []),
    (SITE + "2020/07/Interim_Report_Q3_-_2019.pdf", 2019, 9, []),
    (SITE + "2020/08/Annual_Report_2019.pdf", 2019, 12, []),
    (SITE + "2020/09/Interim_Report_Q1_-_2020.pdf", 2020, 3, []),
    (SITE + "2020/11/Interim_Report_Q2_-_2020.pdf", 2020, 6, []),
    (SITE + "2020/11/דוח-רבעון-3-לשנת-2020.pdf", 2020, 9, []),
    (SITE + "2021/06/Annual_Report_2020.pdf", 2020, 12, []),
    (SITE + "2021/05/דוח-רבעון-1-לשנת-2021.pdf", 2021, 3, []),
    (SITE + "2021/08/דוח-רבעון-2-חצי-שנתי-לשנת-2021.pdf", 2021, 6, [
        (r"30 ביוני ?2021 הסתכמה בכ ?- ?2\.9 BCM וכ ?- ?0\.5 BCM.{0,100}?3\.3 BCM וכ ?- ?0\.1 BCM",
         [(("YTD", 2021, 6), "dom", 2.9), (("YTD", 2021, 6), "exp", 0.5),
          (("YTD", 2020, 6), "dom", 3.3), (("YTD", 2020, 6), "exp", 0.1)])]),
    (SITE + "2021/11/דוח-רבעון-3-לשנת-2021.pdf", 2021, 9, [
        (r"30 בספטמבר ?2021 הסתכמה בכ ?- ?5\.5 BCM וכ ?- ?0\.7 BCM.{0,100}?5\.5 BCM וכ ?- ?0\.3 BCM",
         [(("YTD", 2021, 9), "dom", 5.5), (("YTD", 2021, 9), "exp", 0.7),
          (("YTD", 2020, 9), "dom", 5.5), (("YTD", 2020, 9), "exp", 0.3)])]),
    # English periodic report 2021, section 6.6: Jordan 0.23 / 0.18, Egypt 0.25 / 0.77, domestic 7.77 / 7.71,
    # total 8.25 / 8.66 BCM in 2020 / 2021 (export = Jordan + Egypt)
    (SITE + "2022/10/Tamar_Petroleum_Ltd_2021.pdf", 2021, 12, [
        (r"0\.23 BCM and approximately 0\.18 BCM.{0,250}?0\.25 BCM in 2020 and to approximately 0\.77 BCM in 2021"
         r".{0,250}?7\.77 BCM and approximately 7\.71 BCM.{0,120}?8\.25 BCM and approximately 8\.66 BCM",
         [(("YTD", 2020, 12), "dom", 7.77), (("YTD", 2020, 12), "exp", 0.48), (("YTD", 2020, 12), "tot", 8.25),
          (("YTD", 2021, 12), "dom", 7.71), (("YTD", 2021, 12), "exp", 0.95), (("YTD", 2021, 12), "tot", 8.66)])]),
    (maya(1452131), 2022, 3, [
        (r"ברבעון הראשון לשנת 2022 הסתכמה בכ ?- ?2\.33 BCM וכ ?- ?0\.28 BCM.{0,200}?1\.49 BCM ו ?- ?0\.24 BCM",
         [(("Q", 2022, 1), "dom", 2.33), (("Q", 2022, 1), "exp", 0.28),
          (("Q", 2021, 1), "dom", 1.49), (("Q", 2021, 1), "exp", 0.24)])]),
    (maya(1471381), 2022, 6, [
        (r"30 ביוני 2022 הסתכמה בכ ?- ?4\.45 BCM וכ ?- ?0\.76 BCM",
         [(("YTD", 2022, 6), "dom", 4.45), (("YTD", 2022, 6), "exp", 0.76)]),
        (r"BCM ?5\.21 \(100% ?\), לעומת כמות כוללת בהיקף של כ ?- ?3\.39 BCM.{0,250}?BCM ?2\.60 \(100% ?\) לעומת"
         r" כמות כוללת בהיקף של כ ?- ?BCM ?1\.66",
         [(("YTD", 2022, 6), "tot", 5.21), (("YTD", 2021, 6), "tot", 3.39),
          (("Q", 2022, 2), "tot", 2.60), (("Q", 2021, 2), "tot", 1.66)])]),
    (maya(1487746), 2022, 9, [
        (r"30 בספטמבר ?2022 הסתכמה בכ ?- ?6\.98 BCM וכ ?- ?0\.95 BCM",
         [(("YTD", 2022, 9), "dom", 6.98), (("YTD", 2022, 9), "exp", 0.95)]),
        (r"BCM ?7\.93 \(100% ?\), לעומת כמות כוללת בהיקף של כ ?- ?6\.15 BCM.{0,250}?BCM ?2\.72 \(100% ?\) לעומת"
         r" כמות כוללת בהיקף של כ ?- ?BCM ?2\.76",
         [(("YTD", 2022, 9), "tot", 7.93), (("YTD", 2021, 9), "tot", 6.15),
          (("Q", 2022, 3), "tot", 2.72), (("Q", 2021, 3), "tot", 2.76)])]),
    (maya(1523849), 2023, 3, []),
    (maya(1542518), 2023, 6, []),
    (maya(1559274), 2023, 9, []),     # quarterly table Q1-Q3 2023 and Q1-Q3 2022
    (maya(1581015), 2023, 12, []),    # quarterly table Q1-Q4 2023
    # FY 2024 report: quarterly table 2024 and 2023, annual domestic/export table 2024 / 2023 / 2022
    (maya(1652267), 2024, 12, [
        (r"לשוק המקומי 6\.68 6\.57 8\.70 כמות גז ?טבעי לייצוא 3\.41 2\.55 1\.55 10\.09 9\.12 10\.25",
         [(("YTD", 2022, 12), "dom", 8.70), (("YTD", 2022, 12), "exp", 1.55), (("YTD", 2022, 12), "tot", 10.25),
          (("YTD", 2023, 12), "dom", 6.57), (("YTD", 2023, 12), "exp", 2.55), (("YTD", 2023, 12), "tot", 9.12)])]),
]
LAST_YEAR = 2023                      # the regular pull owns Q1 2024 on
GOVIL = {2020: 8.2, 2022: 10.1, 2023: 9.1, 2024: 10.1}   # Natural Gas Authority annual reviews, Tamar total BCM

SALES_ROW = re.compile(r"(?:Gas sales in BCM|מכירת ?גז ?ב- ?BCM)\s?\d?\s+((?:\d+\.\d+\s+){1,4}\d+\.\d+)")
QUARTER_ROW = re.compile(r"רבעון ?([1-4])(?: ?(20\d\d))? (\d+\.\d\d) (\d+\.\d\d) (\d+\.\d\d)")


def log(msg):
    print(msg, flush=True)


def download(url):
    for i in range(4):
        try:
            r = requests.get(url, headers=UA, timeout=240)
            if r.status_code == 200 and r.content.startswith(b"%PDF"):
                return r.content
            log(f"  {url}: HTTP {r.status_code}")
        except requests.RequestException as e:
            log(f"  {url}: {type(e).__name__}")
        time.sleep(10 * (i + 1))
    raise RuntimeError(f"could not download {url}")


def decimals(v):
    s = f"{v}"
    return len(s.split(".")[1]) if "." in s else 0


def report_facts(text, year, months, prose):
    """[(period, kind, value, decimals)] read from one report's text."""
    t = re.sub(r"\s+", " ", text)
    out = []
    rows = SALES_ROW.findall(t)
    if rows:
        vals = rows[0].split()
        dec = max(len(v.split(".")[1]) for v in vals)
        if months == 12:
            periods = [("YTD", year, 12), ("YTD", year - 1, 12)]
        elif months == 3:
            periods = [("Q", year, 1), ("Q", year - 1, 1), ("YTD", year - 1, 12)]
        else:
            periods = [("YTD", year, months), ("YTD", year - 1, months), ("Q", year, months // 3),
                       ("Q", year - 1, months // 3), ("YTD", year - 1, 12)]
        if len(vals) != len(periods):
            raise RuntimeError(f"{year} {months}M: sales row {vals} does not have {len(periods)} figures")
        out += [(p, "tot", float(v), dec) for p, v in zip(periods, vals)]
    for m in QUARTER_ROW.finditer(t):
        y = int(m.group(2)) if m.group(2) else year
        dom, exp, tot = (float(m.group(i)) for i in (3, 4, 5))
        if abs(dom + exp - tot) > 0.021:
            raise RuntimeError(f"quarter table Q{m.group(1)} {y}: {dom} + {exp} != {tot}")
        out += [(("Q", y, int(m.group(1))), k, v, 2) for k, v in (("dom", dom), ("exp", exp), ("tot", tot))]
    for pat, facts in prose:
        if not re.search(pat, t):
            raise RuntimeError(f"{year} {months}M: prose figures not found: {pat[:80]}")
        dec = max(decimals(v) for _p, _k, v in facts)   # 2.60 prints as 2.6: the group's precision counts
        out += [(p, k, v, dec) for p, k, v in facts]
    if not out:
        raise RuntimeError(f"{year} {months}M: no figures found")
    return out


def build(facts):
    """facts: [(period, kind, value, decimals, source, rank)] -> {(year, q): {kind: (value, dec, source, how)}}."""
    best = {}
    for p, k, v, d, src, rank in facts:
        key = (p, k)
        old = best.get(key)
        if old and abs(old[0] - v) > 0.1 + 1e-9:
            log(f"  NOTE {p} {k}: {old[0]} ({old[2].rsplit('/', 1)[-1]}) vs {v} ({src.rsplit('/', 1)[-1]})")
        if old is None or (d, rank) > (old[1], old[3]):
            best[key] = (v, d, src, rank)

    def get(p, k):
        if p[0] == "YTD" and p[2] == 3 and (p, k) not in best:
            p = ("Q", p[1], 1)
        return best.get((p, k))

    years = sorted({p[1] for (p, _k) in best})
    out = {}
    for y in years:
        if y > LAST_YEAR:
            continue
        for q in range(1, 5):
            row = {}
            for k in ("tot", "dom", "exp"):
                cands = []
                own = best.get((("Q", y, q), k))
                if own:
                    cands.append((own[0], own[1], 1, own[2], "reported"))
                if q > 1:
                    a, b = get(("YTD", y, 3 * q), k), get(("YTD", y, 3 * (q - 1)), k)
                    if a and b:
                        cands.append((a[0] - b[0], min(a[1], b[1]), 0, f"{a[2]} minus {b[2]}",
                                      f"{3 * q}M - {3 * (q - 1)}M"))
                else:
                    a, b = get(("YTD", y, 6), k), best.get((("Q", y, 2), k))
                    if a and b:
                        cands.append((a[0] - b[0], min(a[1], b[1]), 0, f"{a[2]} minus {b[2]}", "H1 - Q2"))
                if cands:
                    v, d, _, src, how = max(cands, key=lambda c: (c[1], c[2]))
                    row[k] = (round(v, 2), d, src, how)
            if "dom" in row and "exp" in row:
                s = round(row["dom"][0] + row["exp"][0], 2)
                tol = 0.021 if min(row["dom"][1], row["exp"][1]) >= 2 and row.get("tot", (0, 2))[1] >= 2 else 0.101
                if "tot" in row and abs(s - row["tot"][0]) > tol:
                    log(f"  {y} Q{q}: domestic {row['dom'][0]} + export {row['exp'][0]} = {s} vs total "
                        f"{row['tot'][0]}: split left blank")
                    del row["dom"], row["exp"]
                elif "tot" not in row or min(row["dom"][1], row["exp"][1]) > row["tot"][1]:
                    row["tot"] = (s, min(row["dom"][1], row["exp"][1]), row["dom"][2], "domestic + export")
            if "tot" in row:
                out[(y, q)] = row
    return out, best


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=igf.DEFAULT_OUT)
    ap.add_argument("--dry-run", action="store_true", help="derive and check only, do not write the workbook")
    args = ap.parse_args()

    facts = []
    for rank, (url, year, months, prose) in enumerate(SOURCES):
        text = igf.pdf_text(download(url))
        got = report_facts(text, year, months, prose)
        log(f"{year} {months:>2}M {url.rsplit('/', 1)[-1]}: {[(p, k, v) for p, k, v, _d in got]}")
        facts += [(p, k, v, d, url, rank) for p, k, v, d in got]

    quarters, best = build(facts)
    errors = []
    log("\nQUARTERS (BCM; how derived)")
    for (y, q), row in sorted(quarters.items()):
        log(f"  {y} Q{q}: total {row['tot'][0]:.2f} ({row['tot'][3]})"
            + (f", domestic {row['dom'][0]:.2f} ({row['dom'][3]}), export {row['exp'][0]:.2f} ({row['exp'][3]})"
               if "dom" in row else ""))
    log("\nANNUAL CHECK (sum of quarters vs company annual total vs gov.il)")
    for y in sorted({y for y, _q in quarters}):
        qs = [quarters.get((y, q)) for q in range(1, 5)]
        fy = best.get((("YTD", y, 12), "tot"))
        if not all(qs):
            log(f"  {y}: quarters {[q for q in range(1, 5) if (y, q) in quarters]} only")
            continue
        s = round(sum(r["tot"][0] for r in qs), 2)
        line = f"  {y}: quarters {s:.2f}"
        if fy:
            line += f", company {fy[0]}"
            if abs(s - fy[0]) > 0.15 + 1e-9:
                errors.append(f"{y}: quarters {s} vs company annual {fy[0]}")
        if y in GOVIL:
            line += f", gov.il {GOVIL[y]}"
            if abs(s - GOVIL[y]) > 0.3:
                line += "  GAP > 0.3 BCM"
        log(line)
    if errors:
        sys.exit("CHECK FAILED: " + "; ".join(errors))
    if args.dry_run:
        return

    def source(r):
        if "dom" not in r or r["tot"][2] == r["dom"][2] == r["exp"][2]:
            return r["tot"][2]
        return f"total: {r['tot'][2]} | domestic: {r['dom'][2]} | export: {r['exp'][2]}"

    rows = [{"date": igf.qdate(y, q), "Domestic_bcm": r["dom"][0] if "dom" in r else None,
             "Export_bcm": r["exp"][0] if "exp" in r else None, "Total_bcm": r["tot"][0], "Source": source(r)}
            for (y, q), r in sorted(quarters.items())]
    new = pd.DataFrame(rows)
    sheets = {s: igf.read_sheet(args.out, s, []) for s in ("Quarterly", "Leviathan", "Tamar", "Karish YTD", "Reports")}
    tam = sheets["Tamar"]
    tam["date"] = pd.to_datetime(tam["date"])
    keep = tam[tam["date"] >= pd.Timestamp(LAST_YEAR + 1, 1, 1)]
    tam = pd.concat([new, keep], ignore_index=True).sort_values("date")
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M")
    reps = sheets["Reports"]
    reps = reps[~reps["Source"].isin([u for u, *_ in SOURCES])]
    used = pd.Series([f[4] for f in facts]).value_counts()
    reps = pd.concat([reps, pd.DataFrame([(u, "Tamar (history backfill)", now, int(used.get(u, 0)))
                                          for u, *_ in SOURCES], columns=["Source", "Field", "Read_utc", "Rows"])],
                     ignore_index=True)
    xlsx_notes.write_workbook(args.out, {"Quarterly": sheets["Quarterly"].set_index("date"),
                                         "Leviathan": sheets["Leviathan"].set_index("date"),
                                         "Tamar": tam.set_index("date"),
                                         "Karish YTD": sheets["Karish YTD"].set_index("Year"),
                                         "Reports": reps.set_index("Source")}, igf.NOTES_LINES, igf.NOTES_TITLES)
    log(f"\nSaved {args.out}: {len(new)} Tamar quarters {new['date'].min():%Y-%m} to {new['date'].max():%Y-%m} "
        f"written, {len(keep)} stored quarters from {LAST_YEAR + 1} kept")


if __name__ == "__main__":
    main()
