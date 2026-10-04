"""
Cambodia electricity generation by type and imports by country, ANNUAL, from the Electricity Authority of Cambodia
(EAC, the power-sector regulator): https://eac.gov.kh/site/annualreport?lang=en
Found via discovery_archive/asia/CAMBODIA_DISCOVERY1-4.py. Nothing official is published monthly or daily: EAC, EDC
(Electricite du Cambodge), the Ministry of Mines and Energy and Open Development Cambodia only publish annual figures
(EDC's own annual reports stop at 2018).

Two EAC publications are read:
  'Report on Power Sector of the Kingdom of Cambodia' (one English PDF per year, 2003 on, published about ten months
      after the year: Annual-Report-YYYY-en.pdf). Its Annex 2 tables give, for the year and the year before:
        2(a) generation in Cambodia and imports from Vietnam / Thailand / Laos at HV and MV, million kWh
        2(b) energy sent out by type of licensee (IPPs, EDC, other licensees) - from 2024 EDC's 'Hydro Power Plant
             from Laos' (a dedicated Lao plant wired to the Cambodian grid) is counted here as generation
        2(c) installed capacity (kW) and energy sent out by type: hydro, diesel/HFO, biomass, coal, solar
  'Salient Features of Power Development' (salient_feature_YYYY_en.pdf, published each December): capacity and energy
      by source and imports by country for last year, the current year (near-final) and the plan for next year. Used
      only for years with no annual report yet (it adds the latest year about ten months earlier); it also gives
      rooftop solar, which the annual report leaves out.

Writes output/Data and Chart Outputs/cambodia_power_generation.xlsx:
  Daily     one row per YEAR (dated 1 January), MWh: Hydro, Gas (0), Wind (0), Solar (licensed plants), Coal, Oil
            (diesel/HFO), Bioenergy, Other (0); Total_MWh = their sum = domestic generation; Imports_MWh and
            Imports_Vietnam / _Thailand / _Laos_MWh (Laos incl. EDC's dedicated Lao hydro plant, shown alone as
            Imports_Laos_EDC_hydro_MWh), Imports_HV_MWh / Imports_MV_MWh, Supply_MWh (domestic + imports), and the
            licensee split IPP_MWh / EDC_MWh / Other_licensees_MWh; Solar_rooftop_MWh (salient features, not in Total)
  Monthly   (capacity; named for the standard capacity layout) one row per year: <Fuel>_MW at the end of the year
  Basis     per year: which publication the row comes from
  Raw       every value read, by publication (the history store: editions already read are not downloaded again)

Incremental: the EAC listing page is read each run; a PDF is downloaded only when it is new or its Last-Modified
differs from the one recorded on the Units sheet. Runs on the 1st and 15th.

    python3 asia/CAMBODIA_EAC.py [--out path]
"""
import argparse
import io
import os
import re
import sys
import time

import pandas as pd
import pdfplumber
import pypdfium2 as pdfium
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

BASE = "https://eac.gov.kh"
LIST = BASE + "/site/annualreport?lang=en"
AR_URL = BASE + "/uploads/annual_report/english/Annual-Report-{y}-en.pdf"
SF_URL = BASE + "/uploads/salient_feature/english/salient_feature_{y}_en.pdf"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36"}
T = (20, 240)
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "cambodia_power_generation.xlsx")
FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Oil", "Bioenergy", "Other"]
ANNEX2 = r"Annex\s*2\s*\(?[abc]\)?"
STAMP = "Read: "   # Units-sheet line prefix: 'Read: <file> | Last-Modified: <date>'


def out(*a):
    print(*a, flush=True)


def num(s):
    """'1,234.56' / '1769.96' / '-' / '' -> float or None."""
    if s is None:
        return None
    s = str(s).replace(",", "").replace("\n", " ").strip()
    m = re.fullmatch(r"-?\d+(?:\.\d+)?", s)
    return float(s) if m else None


def years_in(cell):
    return [int(y) for y in re.findall(r"\b((?:19|20)\d{2})\b", str(cell or ""))]


# ---------------------------------------------------------------- annual report (Annex 2 tables)
def type_key(label):
    s = label.lower()
    for pat, k in [(r"hydro", "Hydro"), (r"diesel|hfo|steam|back-?up|operating|fuel", "Oil"),
                   (r"bio|wood", "Bioenergy"), (r"coal", "Coal"), (r"solar", "Solar"), (r"wind", "Wind"),
                   (r"^\W*total", "Total")]:
        if re.search(pat, s):
            return k
    return None


def licensee_key(label):
    s = label.lower()
    for pat, k in [(r"laos", "EDC_Laos_hydro"), (r"independent", "IPP"), (r"other|consolidated", "Other_licensees"),
                   (r"cambodge|back-?up|operating", "EDC"), (r"^\W*total", "Total")]:
        if re.search(pat, s):
            return k
    return None


def import_key(label):
    s = " ".join(label.lower().split())
    m = re.search(r"import from (vietnam|thailand|laos?) at (hv|mv)", s)
    if m:
        return f"imp_{m.group(1).title().replace('Lao', 'Laos').replace('Laoss', 'Laos')}_{m.group(2).upper()}"
    if "generation in cambodia" in s:
        return "gen_in_cambodia"
    if re.match(r"^\W*total", s):
        return "available_total"
    return None


def parse_table(tb):
    """Annex table -> (kind, {(field, year): value}). Year columns come from the header row with the most year
    cells (proportion columns, 'in % for 2024', are skipped); the last two are energy (previous, current year),
    the two before them installed capacity where present."""
    best, ycols = None, {}
    for r in tb[:4]:
        yc = {j: years_in(c)[0] for j, c in enumerate(r) if years_in(c) and not re.search(r"%|proportion", str(c), re.I)}
        if len(yc) > len(ycols):
            best, ycols = r, yc
    if len(ycols) < 2:
        return None, {}
    cols = sorted(ycols)
    energy = cols[-2:]
    cap = cols[-4:-2] if len(cols) >= 4 else []
    first = min(cols)
    rows = [r for r in tb if r is not best]
    labels = [" ".join(str(c) for c in r[:first] if c and num(c) is None) for r in rows]
    alltext = " ".join(labels).lower()
    if "import from" in alltext:
        kind, keyf = "imports", import_key
    elif "hydro" in alltext and "independent" not in alltext:
        kind, keyf = "types", type_key
    elif "independent" in alltext:
        kind, keyf = "licensees", licensee_key
    else:
        return None, {}
    vals = {}
    for r, lab in zip(rows, labels):
        k = keyf(lab)
        if not k:
            continue
        for j in energy:
            v = num(r[j]) if j < len(r) else None
            if v is not None:
                f = k if kind == "imports" else (f"gen_{k}" if kind == "types" else f"lic_{k}")
                vals[(f, ycols[j])] = vals.get((f, ycols[j]), 0.0) + v   # Diesel back-up + operating rows add up
        for j in cap if kind != "imports" else []:
            v = num(r[j]) if j < len(r) else None
            if v is not None:
                f = f"cap_{k}" if kind == "types" else f"lcap_{k}"
                vals[(f, ycols[j])] = vals.get((f, ycols[j]), 0.0) + v / 1000.0   # kW -> MW
    return kind, vals


NUM = re.compile(r"^-?\d[\d,]*(?:\.\d+)?$")


def parse_annex_text(text, edition):
    """Text fallback for one Annex 2 page (pypdfium2 text, one table row per line): the energy columns are the
    edition year and the year before; a row holds [counts], [capacity prev, cur, %], energy prev, cur, %."""
    vals, kind, carry = {}, None, ""
    min_full = {"imports": 3, "types": 6, "licensees": 8}
    for ln in text.splitlines():
        ln = ln.replace("\x12", " ").strip()
        if re.search(r"Summary Information", ln, re.I):
            kind = ("imports" if re.search(r"Import", ln, re.I) else "types" if re.search(r"Generation Type", ln, re.I)
                    else "licensees" if re.search(r"Sent", ln, re.I) else None)
            carry = ""
            continue
        if not kind:
            continue
        toks = re.sub(r"^\d{1,2}\s+(?=[A-Za-z])", "", ln).split()
        nums = [num(t) for t in toks if NUM.match(t)]
        label = " ".join(t for t in toks if not NUM.match(t))
        if not nums:
            carry = (carry + " " + label).strip()[-120:]
            continue
        label, carry = (carry + " " + label).strip(), ""
        keyf = {"imports": import_key, "types": type_key, "licensees": licensee_key}[kind]
        k = keyf(label)
        if not k or len(nums) < 2:
            continue
        f = k if kind == "imports" else (f"gen_{k}" if kind == "types" else f"lic_{k}")
        pairs = [(edition, nums[-2])] + ([(edition - 1, nums[-3])] if len(nums) >= min_full[kind] else [])
        for y, v in pairs:
            vals[(f, y)] = vals.get((f, y), 0.0) + v
        if kind == "types" and len(nums) >= 4:
            capcur = nums[-5] if len(nums) >= 6 else nums[-4]
            vals[(f"cap_{k}", edition)] = vals.get((f"cap_{k}", edition), 0.0) + capcur / 1000.0
            if len(nums) >= 6:
                vals[(f"cap_{k}", edition - 1)] = vals.get((f"cap_{k}", edition - 1), 0.0) + nums[-6] / 1000.0
    return vals


def parse_annual_report(content, edition):
    vals = {}
    doc = pdfium.PdfDocument(content)   # fast text pass to find the Annex 2 pages; pdfplumber reads only those
    texts = {i: doc[i].get_textpage().get_text_range() for i in range(len(doc))}
    doc.close()
    pages = [i for i, t in texts.items() if re.search(ANNEX2, " ".join(t.split()[:40]), re.I)]
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for i in pages:
            found = 0
            for tb in pdf.pages[i].extract_tables():
                kind, v = parse_table(tb)
                if kind:
                    out(f"    p{i + 1}: {kind} {len(v)} values")
                    found += len(v)
                    for k, x in v.items():
                        vals.setdefault(k, x)
            if not found:   # 2019 / 2020 print these pages rotated: pdfplumber reads them reversed
                v = parse_annex_text(texts[i], edition)
                out(f"    p{i + 1}: text fallback {len(v)} values")
                for k, x in v.items():
                    vals.setdefault(k, x)
    return [{"Publication": f"Annual report {edition}", "Edition": edition, "Field": f, "Year": y, "Value": v}
            for (f, y), v in sorted(vals.items(), key=lambda kv: (kv[0][1], kv[0][0]))]


# ---------------------------------------------------------------- salient features (text table)
SF_LINES = [(r"^\+\s*Hydro", "gen_Hydro"), (r"^\.?\s*Solar Power Station", "gen_Solar"),
            (r"^\.?\s*Rooftop", "gen_Solar_rooftop"), (r"^\+\s*Biomass", "gen_Bioenergy"), (r"^\+\s*Coal", "gen_Coal"),
            (r"^\+\s*Fuel Oil", "gen_Oil"), (r"^-\s*Thailand", "imp_Thailand"), (r"^-\s*Vietnam", "imp_Vietnam"),
            (r"^-\s*Laos", "imp_Laos")]


def parse_salient(content, edition):
    rows = []
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for pg in pdf.pages[:4]:
            t = pg.extract_text() or ""
            if "Different Sources of Power" not in t:
                continue
            lines = t.splitlines()
            yrs = None
            for ln in lines:
                y = years_in(ln)
                if yrs is None and len(y) >= 2 and re.search(r"plan", ln, re.I):
                    yrs = y[:2]   # 'YYYY YYYY Plan for YYYY': actual years only
                    continue
                if not yrs:
                    continue
                for pat, f in SF_LINES:
                    if re.search(pat, ln.strip(), re.I):
                        nums = [num(x) for x in ln.split() if "%" not in x and num(x) is not None]
                        if len(nums) >= 4:   # cap, energy for each of the two years (plan year after)
                            for k, y in enumerate(yrs):
                                rows.append({"Publication": f"Salient features {edition}", "Edition": edition,
                                             "Field": f, "Year": y, "Value": nums[2 * k + 1]})
                                rows.append({"Publication": f"Salient features {edition}", "Edition": edition,
                                             "Field": f.replace("gen_", "cap_").replace("imp_", "capimp_"),
                                             "Year": y, "Value": nums[2 * k]})
                        break
    return rows


# ---------------------------------------------------------------- fetch with Last-Modified check
def recorded_stamps(path):
    try:
        u = pd.read_excel(path, sheet_name="Units", header=None).iloc[:, 0].dropna().astype(str)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return {}
    st = {}
    for v in u:
        if v.startswith(STAMP) and " | Last-Modified: " in v:
            name, lm = v[len(STAMP):].split(" | Last-Modified: ", 1)
            st[name.strip()] = lm.strip()
    return st


def head_lm(url):
    """Last-Modified of a published PDF, None if it is not there. HEAD first; a streamed GET (headers only) if the
    server refuses HEAD or answers it oddly."""
    for method in ("head", "get"):
        for i in range(2):
            try:
                r = requests.request(method, url, headers=H, timeout=(20, 60), allow_redirects=True, stream=True)
                r.close()
                if r.status_code == 200 and "pdf" in (r.headers.get("content-type") or ""):
                    return r.headers.get("last-modified") or "unknown"
                if r.status_code == 404:
                    return None
                out(f"  {method.upper()} {url.rsplit('/', 1)[-1]}: HTTP {r.status_code} {r.headers.get('content-type')}")
            except requests.RequestException as e:
                out(f"  {method.upper()} {url.rsplit('/', 1)[-1]}: {e}")
            time.sleep(5)
    return None


def download(url):
    for i in range(3):
        try:
            r = requests.get(url, headers=H, timeout=T)
            if r.status_code == 200 and r.content[:4] == b"%PDF":
                return r.content
            out(f"  {url}: HTTP {r.status_code}")
            return None
        except requests.RequestException as e:
            out(f"  {url}: {e}")
    return None


def editions():
    ar, sf = set(), set()
    for i in range(3):
        try:
            r = requests.get(LIST, headers=H, timeout=(20, 90))
            ar = {int(y) for y in re.findall(r"Annual-Report-(\d{4})-en\.pdf", r.text)}
            sf = {int(y) for y in re.findall(r"salient_feature_(\d{4})_en\.pdf", r.text)}
            if ar:
                break
            out(f"listing page: HTTP {r.status_code}, {len(r.text)} chars, no report links")
        except requests.RequestException as e:
            out(f"listing page: {e}")
        time.sleep(10)
    this = pd.Timestamp.today().year
    if not ar:   # listing unreadable: try every edition's file name (2003 = the first English report)
        ar = set(range(2003, this + 1))
    for y in (this - 1, this):   # the listing links only the newest salient features; probe the file names too
        if y not in ar:
            ar.add(y)
        if y not in sf:
            sf.add(y)
    return sorted(ar), sorted(sf)


# ---------------------------------------------------------------- build
def pick(raw, field, year, prefix):
    """Latest edition's value for (field, year) among publications starting with prefix."""
    r = raw[(raw.Field == field) & (raw.Year == year) & raw.Publication.str.startswith(prefix)].dropna(subset=["Value"])
    return float(r.sort_values("Edition").Value.iloc[-1]) if len(r) else None


def build(raw):
    ar = raw[raw.Publication.str.startswith("Annual report")]
    sf = raw[raw.Publication.str.startswith("Salient")]
    ar_years = sorted(set(ar[ar.Field.str.startswith("gen_") & (ar.Field != "gen_in_cambodia")].Year))
    rows, caps, basis = {}, {}, {}
    for y in sorted(set(ar_years) | set(sf.Year)):
        d = pd.Timestamp(y, 1, 1)
        if y in ar_years:
            src = "Annual report"
            g = {f: pick(raw, f"gen_{f}", y, src) for f in FUELS}
            laos_plant = pick(raw, "lic_EDC_Laos_hydro", y, src)
            if laos_plant and g["Hydro"]:
                g["Hydro"] -= laos_plant   # EDC's Lao plant: generated in Laos -> imports, not Cambodian hydro
            imp = {c: sum(v for v in (pick(raw, f"imp_{c}_HV", y, src), pick(raw, f"imp_{c}_MV", y, src)) if v)
                   or None for c in ("Vietnam", "Thailand", "Laos")}
            hv = [pick(raw, f"imp_{c}_HV", y, src) for c in ("Vietnam", "Thailand", "Laos")]
            mv = [pick(raw, f"imp_{c}_MV", y, src) for c in ("Vietnam", "Thailand", "Laos")]
            if laos_plant:
                imp["Laos"] = (imp["Laos"] or 0) + laos_plant
            lic = {k: pick(raw, f"lic_{k}", y, src) for k in ("IPP", "EDC", "Other_licensees")}
            ed = int(ar[(ar.Year == y) & ar.Field.isin([f"gen_{f}" for f in FUELS])].Edition.max())
            basis[d] = (f"EAC Report on Power Sector {ed} (Annex 2)" +
                        (f"; hydro excludes EDC's dedicated Lao hydro plant ({laos_plant:,.0f} GWh), counted as "
                         "imports from Laos" if laos_plant else ""))
            cap = {f: pick(raw, f"cap_{f}", y, src) for f in FUELS}
            laos_cap = pick(raw, "lcap_EDC_Laos_hydro", y, src) if laos_plant else None
            if laos_cap and cap["Hydro"]:
                cap["Hydro"] -= laos_cap
            rooftop = pick(raw, "gen_Solar_rooftop", y, "Salient")
        else:
            src = "Salient"
            g = {f: pick(raw, f"gen_{f}", y, src) for f in FUELS}
            imp = {c: pick(raw, f"imp_{c}", y, src) for c in ("Vietnam", "Thailand", "Laos")}
            hv, mv, lic, laos_plant = [None], [None], {}, None
            ed = int(sf[sf.Year == y].Edition.max())
            basis[d] = (f"EAC Salient Features of Power Development {ed} (provisional until the annual report); "
                        "hydro includes EDC's dedicated Lao hydro plant (not split out in this publication)")
            cap = {f: pick(raw, f"cap_{f}", y, src) for f in FUELS}
            rooftop = pick(raw, "gen_Solar_rooftop", y, src)
        if not any(v for v in g.values()):
            continue
        row = {f"{f}_MWh": (g[f] or 0.0) * 1000 for f in FUELS}
        row["Total_MWh"] = sum(row.values())
        iv = {c: (v * 1000 if v is not None else None) for c, v in imp.items()}
        row["Imports_MWh"] = sum(v for v in iv.values() if v is not None) if any(v is not None for v in iv.values()) else None
        for c in ("Vietnam", "Thailand", "Laos"):
            row[f"Imports_{c}_MWh"] = iv[c]
        row["Imports_Laos_EDC_hydro_MWh"] = laos_plant * 1000 if laos_plant else None
        row["Imports_HV_MWh"] = sum(v for v in hv if v) * 1000 + (laos_plant or 0) * 1000 if any(hv) else None
        row["Imports_MV_MWh"] = sum(v for v in mv if v) * 1000 if any(mv) else None
        row["Supply_MWh"] = row["Total_MWh"] + row["Imports_MWh"] if row["Imports_MWh"] is not None else None
        for k in ("IPP", "EDC", "Other_licensees"):
            row[f"{k}_MWh"] = lic[k] * 1000 if lic.get(k) is not None else None
        row["Solar_rooftop_MWh"] = rooftop * 1000 if rooftop else None
        rows[d] = row
        c = {f"{f}_MW": cap[f] for f in FUELS if cap.get(f)}
        if c:
            c["Total_MW"] = sum(c.values())
            caps[d] = c
    daily = pd.DataFrame.from_dict(rows, orient="index").sort_index().round(0)
    capdf = pd.DataFrame.from_dict(caps, orient="index").sort_index().round(1)
    bas = pd.DataFrame({"Basis": pd.Series(basis)}).sort_index()
    for f in (daily, capdf, bas):
        f.index.name = "date"
    return daily, capdf, bas


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    stamps = recorded_stamps(args.out)
    try:
        raw = pd.read_excel(args.out, sheet_name="Raw")
    except (FileNotFoundError, ValueError, OSError):
        raw = pd.DataFrame(columns=["Publication", "Edition", "Field", "Year", "Value"])
    ar_eds, sf_eds = editions()
    out(f"annual-report editions listed: {ar_eds}; salient features: {sf_eds}")
    new = []
    for kind, eds, url_t, parser in (("Annual report", ar_eds, AR_URL, parse_annual_report),
                                     ("Salient features", sf_eds, SF_URL, parse_salient)):
        for y in eds:
            url = url_t.format(y=y)
            name = url.rsplit("/", 1)[-1]
            if name in stamps:   # read before (with or without values)
                lm = head_lm(url) if y >= max(eds) - 1 else stamps[name]   # only recent editions get revised
                if lm in (None, stamps[name]):
                    continue
            else:
                lm = head_lm(url)
                if lm is None:
                    continue
            out(f"  {name}: downloading (Last-Modified {lm})")
            b = download(url)
            if not b:
                continue
            rows = parser(b, y)
            out(f"  {name}: {len(rows)} values")
            if rows:
                raw = raw[raw.Publication != f"{kind} {y}"]
                new.extend(rows)
            stamps[name] = lm   # recorded even when nothing was found, so it is not downloaded every run
    if new:
        raw = pd.concat([raw, pd.DataFrame(new)], ignore_index=True)
    if raw.empty:
        raise SystemExit("No EAC data read")
    raw = raw.sort_values(["Publication", "Year", "Field"]).reset_index(drop=True)
    raw["Edition"] = raw["Edition"].astype(int)
    raw["Year"] = raw["Year"].astype(int)
    daily, capdf, bas = build(raw)
    # check: the type split should add up to EAC's 'generation in Cambodia'
    for y, r in daily.iterrows():
        g = pick(raw, "gen_in_cambodia", y.year, "Annual report")
        if g:
            laos = (r["Imports_Laos_EDC_hydro_MWh"] or 0) if pd.notna(r["Imports_Laos_EDC_hydro_MWh"]) else 0
            diff = (r["Total_MWh"] + laos) / 1000 - g
            if abs(diff) > max(1.0, 0.01 * g):
                out(f"  WARNING {y.year}: types sum {(r['Total_MWh'] + laos) / 1000:,.1f} GWh vs generation in Cambodia "
                    f"{g:,.1f}")
    out((daily[["Total_MWh", "Imports_MWh", "Supply_MWh"]] / 1000).round(1).to_string())
    out(daily.tail(3).T.to_string())
    files = [f"{STAMP}{n} | Last-Modified: {lm}" for n, lm in sorted(stamps.items())]
    notes = [
        "UNITS",
        "Daily: one row per YEAR (dated 1 January), MWh in the year (EAC publishes million kWh = GWh, x1,000). Hydro, "
        "Solar (licensed solar plants), Coal, Oil (diesel / heavy fuel oil, back-up and operating), Bioenergy; Gas, "
        "Wind and Other are 0 (none in Cambodia). Total_MWh = their sum = domestic generation sent out by licensees.",
        "Imports_MWh = Imports_Vietnam + Imports_Thailand + Imports_Laos (HV grid links plus MV border supplies; "
        "Imports_HV_MWh / Imports_MV_MWh split them), not in Total_MWh. Supply_MWh = Total_MWh + Imports_MWh (EAC's "
        "'total energy available').",
        "Imports_Laos_EDC_hydro_MWh: from 2024 EAC counts EDC's dedicated hydro plant in Laos (460 MW, wired to the "
        "Cambodian grid) as EDC generation; here it is moved to imports from Laos so Hydro_MWh is Cambodian plants "
        "only (until 2023 EAC itself listed this energy as 'import from Laos at HV').",
        "IPP_MWh / EDC_MWh / Other_licensees_MWh: energy sent out by type of licensee (EDC excludes the Lao plant).",
        "Solar_rooftop_MWh: rooftop solar (EAC salient features, from 2024), not in Total_MWh.",
        "Monthly: installed capacity at the end of each year, MW (one row per year; sheet named for the standard "
        "capacity layout); Hydro_MW also leaves out EDC's Lao plant where EAC counts it.",
        "Basis: the publication each year's row comes from. Raw: every value read (GWh / MW) by publication.",
        "",
        "COVERAGE",
        f"Annual, {daily.index.min():%Y} to {daily.index.max():%Y}. Years with an EAC annual report use it (the latest "
        "edition that covers the year, so revisions are picked up); the latest year comes from the salient features "
        "until its annual report appears (about ten months after the year). Imports by country are in the annual "
        "reports from the 2010s editions on; earlier rows have generation by type only.",
        "No official monthly or daily series is published for Cambodia (checked: EAC, EDC, MME, NCC, Open Development "
        "Cambodia, ASEAN Centre for Energy - discovery_archive/asia/CAMBODIA_DISCOVERY*.py).",
        "Each PDF is downloaded only when new or when its Last-Modified changes:",
        *files,
        "",
        "SOURCE",
        "Electricity Authority of Cambodia (EAC): Report on Power Sector of the Kingdom of Cambodia (annual, Annex 2 "
        "tables) and Salient Features of Power Development in the Kingdom of Cambodia, "
        "https://eac.gov.kh/site/annualreport?lang=en",
    ]
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "Monthly": capdf, "Basis": bas,
                                         "Raw": raw.set_index("Publication")}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}: {len(daily)} years {daily.index.min():%Y}..{daily.index.max():%Y}")


if __name__ == "__main__":
    main()
