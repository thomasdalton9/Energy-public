"""
Cameroon monthly energy balance (SONATREL - ENEO "PV de validation du bilan energetique"), published by ARSEL.

Source: ARSEL (Agence de Regulation du Secteur de l'Electricite), https://arsel-cm.org/bilan-energetique-mensuel/
One PDF (~20 pages, text tables) per month. Per plant connection node, MWh injected to the transmission grid
(Songloulou, Edea 3, Lagdo, Limbe HFO; IPPs Nachtigal, Memve'ele, Lom Pangar, Kribi, Dibamba), substation off-take
(approximately sales to ENEO distribution), HT customers and auxiliaries.

Discovery: scrape the bilan-energetique-mensuel page and the WordPress REST media list for PDFs whose name contains
"bilan"; the month is read from the title on page 1 of each PDF (file names are irregular / contain typos).
Incremental: the committed workbook is the history store; only months not yet saved (plus the last two, in case ARSEL
replaces a file) are downloaded and parsed.

    python3 rest_of_world/cameroon_arsel_energy_balance.py --out "output/Data and Chart Outputs/cameroon_arsel_energy_balance.xlsx"
"""
import argparse
import io
import os
import re
import sys
import unicodedata

import pandas as pd
import requests

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import xlsx_notes  # noqa: E402

PAGE = "https://arsel-cm.org/bilan-energetique-mensuel/"
MEDIA = "https://arsel-cm.org/wp-json/wp/v2/media?per_page=100&mime_type=application/pdf&search=bilan&page={}"
HEADERS = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120 Safari/537.36"}
REVISION_MONTHS = 2
MONTHS = {"janvier": 1, "fevrier": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6, "juillet": 7, "aout": 8,
          "septembre": 9, "octobre": 10, "novembre": 11, "decembre": 12}

# plant -> (keyword in the node text, table). Edea 3 / Songloulou / Limbe / Lagdo are in the "centrales" table (E),
# IPPs in the IPP table that follows it.
CENTRALES = [("Songloulou", "songloulou"), ("Edea 3", "actif + d90"), ("Lagdo", "lagdo"), ("Limbe HFO", "limbe hfo"),
             ("Oyomabang", "oyomabang")]
IPPS = [("Nachtigal", "nachtigal"), ("Memve'ele", "memve"), ("Lom Pangar", "lom pangar"), ("Kribi (gas)", "kribi"),
        ("Dibamba", "dibamba")]
CLIENTS = [("CIMAF", "cimaf"), ("CIMENCAM", "cimencam"), ("DANGOTE", "dangote"), ("ALUCAM", "alucam"),
           ("PROMETAL", "prometal")]


def ascii_lower(s):
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().lower()


def get(url, **kw):
    return requests.get(url, headers=HEADERS, timeout=(10, 90), **kw)


def discover():
    """-> {url: last_modified_hint}; https-normalised PDF links that look like monthly energy balances."""
    urls = {}
    try:
        html = get(PAGE).text.replace("\\/", "/")
        for u in re.findall(r'https?://[^"\'\s<>\\]+\.pdf', html, re.I):
            urls[u.replace("http://", "https://")] = ""
    except Exception as exc:
        print("page scrape failed:", exc)
    for pg in range(1, 6):
        try:
            r = get(MEDIA.format(pg))
            if r.status_code != 200:
                break
            items = r.json()
            for it in items:
                u = (it.get("source_url") or "").replace("http://", "https://")
                if u:
                    urls[u] = it.get("modified", "")
            if len(items) < 100:
                break
        except Exception as exc:
            print("media list failed:", exc)
            break
    out = {u: m for u, m in urls.items() if re.search(r"bilan", ascii_lower(os.path.basename(u)))}
    print(f"discovered {len(out)} bilan PDFs")
    return out


def num(tokens):
    """Trailing French-formatted number from a token list (['1', '065,00'] -> 1065.00); None when no decimal comma."""
    if not tokens or not re.fullmatch(r"\d+,\d+", tokens[-1]):
        return None
    ip, fp = tokens[-1].split(",")
    s, last, i = ip, len(ip), len(tokens) - 2
    # absorb thousands groups: the group just taken has exactly 3 digits and the previous token is 1-3 digits
    while i >= 0 and last == 3 and re.fullmatch(r"\d{1,3}", tokens[i]):
        s = tokens[i] + s
        last = len(tokens[i])
        i -= 1
    return float(s + "." + fp)


def parse_pdf(content, hint=None):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        pages = [(p.extract_text() or "") for p in pdf.pages]
    text = "\n".join(pages)
    m = re.search(r"energetique\s+d(?:e\s+|\W\s*)([a-z]+)\s+(20\d\d)", ascii_lower(pages[0]).replace("\n", " "))
    month = None
    if m and ascii_lower(m.group(1)) in MONTHS:
        month = pd.Timestamp(int(m.group(2)), MONTHS[ascii_lower(m.group(1))], 1)
    if month is None and hint:
        base = re.sub(r"[^a-z0-9]", "", ascii_lower(os.path.basename(hint)))
        ym = re.search(r"(20\d\d)", base)
        mm = [v for k, v in MONTHS.items() if k in base]
        if ym and len(mm) == 1:
            month = pd.Timestamp(int(ym.group(1)), mm[0], 1)
    sect = None
    ipp_after_E = False
    rows = {"A": [], "C": [], "D": [], "E": [], "IPPi": [], "IPPs": []}
    printed = {}
    for raw in text.split("\n"):
        line = raw.strip()
        low = ascii_lower(line)
        if re.match(r"^3\s*-\s*durees", low):
            break
        if re.match(r"^a\s*-\s*bilan soutirages postes sources", low):
            sect = "A"
        elif re.match(r"^b\s*-\s*bilan soutirages auxiliaires", low) or (low.startswith("b-") and "auxiliaires" in low):
            sect = "B"
        elif re.match(r"^c\s*-\s*bilan soutirages au niveau des centrales", low):
            sect = "C"
        elif re.match(r"^d\s*-\s*bilan soutirages consommateurs", low):
            sect = "D"
        elif re.match(r"^e\s*-\s*bilan des injections des centrales", low):
            sect = "E"
        elif re.match(r"^f\s*-\s*bilan des injections au niveau des postes", low):
            sect = "F"
        elif re.match(r"^n.? reseau ipps", low):
            sect = "IPPi" if ipp_after_E else "IPPs"
            continue
        if sect == "E":
            ipp_after_E = True
        mt = re.match(r"^TOTAL\s+(.*?)\s+((?:\d+ )*\d+,\d+)$", line)
        if mt:
            printed[(sect, ascii_lower(mt.group(1)))] = num(mt.group(2).split())
            continue
        mr = re.match(r"^\d+\s+(RIS|RIN)\s+(.*)$", line)
        if mr and sect in rows:
            toks = mr.group(2).split()
            v = num(toks)
            if v is not None:
                rows[sect].append((mr.group(1), ascii_lower(mr.group(2)), v))
    return month, rows, printed, text


def agg(rows, mapping):
    out = {}
    for name, key in mapping:
        out[name] = sum(v for _, node, v in rows if key in node)
    return out


def summarise(rows, printed):
    inj = agg(rows["E"], CENTRALES)
    inj.update(agg(rows["IPPi"], IPPS))
    cen_total = sum(v for *_, v in rows["E"])
    ipp_total = sum(v for *_, v in rows["IPPi"])
    inj_rec = {k: round(v, 3) for k, v in inj.items()}
    inj_rec["Total plant injections"] = round(cen_total + ipp_total, 3)
    off = {"Off-take RIS (south grid)": sum(v for n, _, v in rows["A"] if n == "RIS"),
           "Off-take RIN (north grid)": sum(v for n, _, v in rows["A"] if n == "RIN")}
    off["Off-take total"] = off["Off-take RIS (south grid)"] + off["Off-take RIN (north grid)"]
    clients = {k: round(v, 3) for k, v in agg(rows["D"], CLIENTS).items()}
    clients["HT customers total"] = round(sum(v for *_, v in rows["D"]), 3)
    chk = {"Centrales parsed (MWh)": cen_total, "Centrales printed total (MWh)": printed.get(("E", "injection centrales")),
           "IPP parsed (MWh)": ipp_total, "IPP printed total (MWh)": printed.get(("IPPi", "injections des ipps")),
           "HT customers parsed (MWh)": clients["HT customers total"],
           "HT customers printed total (MWh)": printed.get(("D", "soutirage des clients"))}
    cp = chk["Centrales printed total (MWh)"]
    chk["Centrales check"] = "ok" if cp is None or abs(cp - cen_total) < 1 else "MISMATCH"
    ip = chk["IPP printed total (MWh)"]
    chk["IPP check"] = "ok" if ip is None or abs(ip - ipp_total) < 1 else "MISMATCH"
    hp = chk["HT customers printed total (MWh)"]
    chk["HT customers check"] = "ok" if hp is None or abs(hp - clients["HT customers total"]) < 1 else "MISMATCH"
    return inj_rec, {k: round(v, 3) for k, v in off.items()}, clients, chk


def load_existing(path):
    out = {}
    if os.path.exists(path):
        for sh in ("Plant injections", "Substation off-take", "HT customers", "Parse checks", "Source files"):
            try:
                out[sh] = pd.read_excel(path, sheet_name=sh, index_col=0)
            except Exception:
                pass
    return out


NOTES = [
    "Cameroon monthly energy balance (ARSEL / SONATREL - ENEO)",
    "",
    "What it is",
    "Monthly validation minutes (PV de validation du bilan energetique) of the transmission energy balance, signed by SONATREL "
    "(grid operator) and ENEO (utility), published by ARSEL, the electricity regulator. Values are MWh at the metering node.",
    "Sheets: Plant injections (MWh injected per plant to the grid; IPP = independent producers), Substation off-take (MWh drawn at "
    "transmission substations: RIS = southern interconnected grid, RIN = northern grid; approximates distribution-level sales "
    "before distribution losses), HT customers (direct high-voltage customers), Parse checks (parsed sums vs totals printed in the PDF), "
    "Source files (URL, month, size, last-modified).",
    "Hydro: Songloulou, Edea 3, Lagdo, Nachtigal, Memve'ele, Lom Pangar. Thermal: Kribi (gas), Dibamba, Limbe HFO, Oyomabang. "
    "Edea 3 = the four Edea 3 nodes (Ndjock Nkong, Mangombe 1+2, Alucam). Solar is not reported as a separate plant.",
    "",
    "Caveats",
    "Publication lag is 8+ months: the twelve months of 2025 were all posted in February 2026 and no 2026 month had been posted by "
    "October 2026. File names are irregular and contain typos, so months are read from the title on page 1 of each PDF; new files are "
    "found by scraping the ARSEL page and the WordPress media list (names containing 'bilan').",
    "Plant injections are at the grid connection (net of plant auxiliaries). Off-take is transmission-substation withdrawal, not retail sales. "
    "Behind-the-meter / off-grid generation and the Eneo-owned isolated networks are not covered. Values are as validated by SONATREL/ENEO "
    "and can be revised. Any month where Parse checks shows MISMATCH should be treated with caution (a PDF row layout parsed differently).",
    "The ARSEL Power BI dashboards (production, energies produites, transport) are image-only snapshots and are not used.",
    "",
    "Source",
    "ARSEL (Cameroon electricity regulator), monthly energy balance validation minutes: https://arsel-cm.org/bilan-energetique-mensuel/",
]
TITLES = {"What it is", "Caveats", "Source"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/cameroon_arsel_energy_balance.xlsx")
    ap.add_argument("--rebuild", action="store_true")
    a = ap.parse_args()
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    old = {} if a.rebuild else load_existing(a.out)
    have = set(pd.to_datetime(old["Plant injections"].index)) if "Plant injections" in old else set()
    recent = set(sorted(have)[-REVISION_MONTHS:])
    known_urls = set(old["Source files"]["url"]) if "Source files" in old and "url" in old["Source files"] else set()

    found = discover()
    inj_new, off_new, cl_new, chk_new, src_new = {}, {}, {}, {}, {}
    for url in sorted(found):
        # Skip files already parsed (unless they are among the most recent saved months).
        if url in known_urls and not a.rebuild:
            prev = old["Source files"][old["Source files"]["url"] == url]
            pm = pd.to_datetime(prev.index[0]) if len(prev) else None
            if pm is not None and pm not in recent:
                continue
        try:
            r = get(url)
            if r.status_code != 200 or not r.content.startswith(b"%PDF"):
                print("skip", url, r.status_code)
                continue
            month, rows, printed, _ = parse_pdf(r.content, url)
            if month is None or not rows["E"]:
                print("could not parse month/tables:", url)
                continue
            inj, off, cl, chk = summarise(rows, printed)
            inj_new[month], off_new[month], cl_new[month], chk_new[month] = inj, off, cl, chk
            src_new[month] = {"url": url, "bytes": len(r.content), "last_modified": r.headers.get("Last-Modified", ""),
                              "listing_modified": found[url]}
            print(month.strftime("%Y-%m"), inj["Total plant injections"], chk["Centrales check"], chk["IPP check"])
        except Exception as exc:
            print("failed", url, exc)

    if not inj_new and not old:
        print("nothing parsed and no history; no workbook written")
        return

    def merge(name, new):
        df_new = pd.DataFrame.from_dict(new, orient="index") if new else pd.DataFrame()
        df_old = old.get(name, pd.DataFrame())
        if len(df_old):
            df_old.index = pd.to_datetime(df_old.index)
        df = pd.concat([df_old[~df_old.index.isin(df_new.index)], df_new]) if len(df_new) else df_old
        df.index = pd.to_datetime(df.index)
        df.index.name = "month"
        return df.sort_index()

    sheets = {n: merge(n, d) for n, d in [("Plant injections", inj_new), ("Substation off-take", off_new),
                                          ("HT customers", cl_new), ("Parse checks", chk_new),
                                          ("Source files", src_new)]}
    for n, df in sheets.items():
        df.index = df.index.strftime("%Y-%m")
        df.index.name = "month"
    notes = NOTES + ["", f"Months covered: {sheets['Plant injections'].index.min()} to {sheets['Plant injections'].index.max()}"]
    xlsx_notes.write_workbook(a.out, sheets, notes, TITLES)
    print("wrote", a.out)


if __name__ == "__main__":
    main()
