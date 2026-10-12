"""
Monthly installed generation capacity for France, Spain, Portugal and Italy from the official sources. Nothing is
copied from ENTSO-E and no month is invented: a period a source does not cover is absent.

  output/Data and Chart Outputs/france_power_capacity_monthly.xlsx
  output/Data and Chart Outputs/spain_power_capacity_monthly.xlsx
  output/Data and Chart Outputs/portugal_power_capacity_monthly.xlsx      (renewables only: PARTIAL)
  output/Data and Chart Outputs/italy_power_capacity_monthly.xlsx         (see ITALY below)

Date convention (same as the ENTSO-E annual files, dated 1 January, and the Nordic / GB / CEE workbooks): the row dated
1 May holds the capacity installed at 00:00 on 1 May, i.e. at the END of April. Total_MW excludes pumped storage and
batteries, as in the ENTSO-E workbooks. Sheet 'Check_ENTSOE' sets every 1 January row against ENTSO-E A68 (the
<country>_power_capacity.xlsx files).

FRANCE - ODRE / RTE-Enedis-ADEeF 'Registre national des installations de production et de stockage d'electricite'
  (odre.opendatasoft.com, dataset registre-national-installation-production-stockage-electricite-agrege; one edition per
  month 'au 31/07/2026', plus year-end editions ...-3112YY). Every installation with its fuel, technology, power (kW) and
  commissioning date (datemiseenservice, else dateraccordement). Month-end capacity is REBUILT as the sum of the
  installations of the CURRENT edition commissioned before the date. Limits, all measured in 'Check_Editions' against the
  year-end editions: (1) installations closed since are not in the current edition, so earlier months are understated by
  them (thermal about -1 to -2 GW at the start of 2021); (2) installations without any date (old EDF hydro and thermal
  plant, 5.97 + 5.05 GW) count as in service throughout; (3) the register aggregates PV below 36 kW per commune under ONE
  early date, so their growth cannot be dated from the current edition: Solar_agg_interpolated_MW is the exact aggregate of
  each edition (year-end editions 2020-25 and the current one) INTERPOLATED LINEARLY in time between editions, and
  Solar_MW = dated (larger) installations + that aggregate. Year-end editions are exact for their date. Quarterly national
  totals from SDES (statistiques.developpement-durable.gouv.fr, 'tableau de bord solaire / eolien') are on 'SDES_quarterly'
  and checked on 'Check_SDES'. Hydro excludes pumped storage (technology 'Pompage turbinage', PumpedStorage_MW);
  Storage_MW = non-hydraulic storage (batteries); Other = geothermal, marine, waste and unclassified thermal.

SPAIN - Red Electrica REData API (apidatos.ree.es), widget generacion/potencia-instalada-generacion (the old name
  'potencia-instalada' returns HTTP 500), national installed capacity by technology, monthly. See run_es().

PORTUGAL - DGEG 'Estatisticas Rapidas das Renovaveis' monthly PDF (dgeg.gov.pt), table 'Potencia instalada' (MW) of the
  renewable plant fleet: hydro (of which pumped), wind, biomass, municipal waste, biogas, geothermal, photovoltaics
  (incl. self-consumption UPAC, small-production UPP and micro/mini). MONTHLY from January 2024 (the PDF's last column is
  the report month); the Dec PDFs and the year-end columns of the latest PDF give 31 December of 2016-2023. No gas, coal or
  other thermal capacity (REN and DGEG publish those annually only): PARTIAL, renewables only.

ITALY - see run_it().

Usage: python3 SWEIT_CAPACITY_MONTHLY.py [--out-dir DIR] [--only fr,es,pt,it]
"""
import argparse
import hashlib
import io
import json
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone

import numpy as np
import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
SES = requests.Session()
SES.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"})
NOW = datetime.now(timezone.utc)
FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]
STD = [f"{f}_MW" for f in FUELS] + ["Total_MW", "PumpedStorage_MW", "Storage_MW"]
FIRST_ROW = pd.Timestamp("2021-01-01")


# ------------------------------------------------------------------------------------------------ helpers
def fetch(url, tries=4, **kw):
    last = None
    for i in range(tries):
        try:
            r = SES.get(url, timeout=(20, 300), **kw)
            if r.status_code in (429, 500, 502, 503, 504, 202):
                last = f"HTTP {r.status_code}"
                time.sleep(5 * (i + 1))
                continue
            return r
        except requests.RequestException as e:
            last = f"{type(e).__name__}: {e}"
            time.sleep(5 * (i + 1))
    raise RuntimeError(f"giving up on {url}: {last}")


def read_sheet(path, sheet, index_col=None):
    if not os.path.exists(path):
        return None
    try:
        return pd.read_excel(path, sheet_name=sheet, index_col=index_col)
    except Exception as e:  # noqa: BLE001
        print(f"  could not read {sheet} from {os.path.basename(path)}: {type(e).__name__}")
        return None


def load_monthly(path, sheet="Monthly"):
    d = read_sheet(path, sheet, index_col=0)
    if d is None or d.empty:
        return pd.DataFrame()
    d.index = pd.to_datetime(d.index, errors="coerce")
    d = d[d.index.notna()]
    d.index.name = "date"
    return d


def standardise(df, cols):
    """Index = 1st of month; keep `cols` (missing ones NaN), Total_MW = fuels except pumped storage / batteries."""
    out = pd.DataFrame(index=pd.DatetimeIndex(df.index, name="date"))
    for c in cols:
        out[c] = pd.to_numeric(df[c], errors="coerce") if c in df else np.nan
    fuel_cols = [f"{f}_MW" for f in FUELS if f"{f}_MW" in out]
    out["Total_MW"] = out[fuel_cols].sum(axis=1, min_count=1)
    return out.round(1).sort_index()


def merge_rows(old, new):
    """New rows replace old rows of the same date; older dates not in `new` are kept."""
    if old is None or old.empty:
        return new.sort_index()
    keep = old[~old.index.isin(new.index)]
    cols = list(dict.fromkeys(list(new.columns) + [c for c in old.columns if c not in new.columns]))
    return pd.concat([keep.reindex(columns=cols), new.reindex(columns=cols)]).sort_index()


def entsoe_check(out_dir, country, monthly, mapping):
    """1 January rows against ENTSO-E A68 ({our column: ENTSO-E column})."""
    p = os.path.join(out_dir, f"{country}_power_capacity.xlsx")
    e = read_sheet(p, "Monthly")
    if e is None or monthly.empty:
        return pd.DataFrame()
    e["date"] = pd.to_datetime(e["date"])
    e = e.set_index("date")
    rows = []
    for d in e.index:
        if d not in monthly.index:
            continue
        for ours, theirs in mapping.items():
            if ours not in monthly or theirs not in e:
                continue
            a, b = monthly.at[d, ours], e.at[d, theirs]
            if pd.isna(a) or pd.isna(b):
                continue
            rows.append({"date": d, "series": ours, "this_file_MW": round(float(a), 1), "ENTSOE_A68_MW": float(b),
                         "difference_MW": round(float(a) - float(b), 1),
                         "difference_pct": round((float(a) - float(b)) / float(b) * 100, 1) if b else np.nan})
    return pd.DataFrame(rows).set_index("date") if rows else pd.DataFrame()


def releases_frame(rows):
    return pd.DataFrame(rows, columns=["source", "checked_utc", "detail", "sha256"]).set_index("source")


def stored_release(path, source):
    r = read_sheet(path, "Releases", index_col=0)
    if r is None or source not in r.index:
        return ""
    return str(r.at[source, "detail"])


def now_s():
    return NOW.strftime("%Y-%m-%d %H:%M UTC")


def write(path, sheets, notes, titles):
    xlsx_notes.write_workbook(path, sheets, notes, titles)


ENTSOE_MAP = {f"{f}_MW": f"{f}_MW" for f in FUELS + ["PumpedStorage"]}


# ------------------------------------------------------------------------------------------------ France
ODRE = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/"
FR_REG = "registre-national-installation-production-stockage-electricite-agrege"
FR_COLS = "nominstallation,filiere,technologie,codecombustible,puismaxinstallee,datemiseenservice,dateraccordement,regime"
FR_FUEL_ORDER = ["Hydro", "PumpedStorage", "Wind_Onshore", "Wind_Offshore", "Solar", "Nuclear", "Gas", "Coal", "Oil", "Bioenergy", "Storage", "Other"]


def fr_meta(ds):
    j = fetch(f"{ODRE}{ds}").json()
    title = j["metas"]["default"].get("title", "")
    m = re.search(r"au (\d{2})/(\d{2})/(\d{4})", title)
    asof = pd.Timestamp(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None
    return title, str(j["metas"]["default"].get("data_processed")), asof


def fr_load(ds):
    r = fetch(f"{ODRE}{ds}/exports/csv?delimiter=%3B&use_labels=false&select={FR_COLS}")
    d = pd.read_csv(io.BytesIO(r.content), sep=";", low_memory=False)
    return d, hashlib.sha256(r.content).hexdigest()


def _dates(s):
    out = pd.to_datetime(s, format="%d/%m/%Y", errors="coerce")
    miss = out.isna() & s.notna()
    if miss.any():
        out[miss] = pd.to_datetime(s[miss], errors="coerce", format="mixed")
    return out


def fr_prepare(d):
    """Add fuel class, MW, in-service date and the <36 kW aggregate flag to a register edition."""
    d = d.copy()
    d["MW"] = pd.to_numeric(d["puismaxinstallee"], errors="coerce") / 1000.0
    d["start"] = _dates(d["datemiseenservice"]).fillna(_dates(d["dateraccordement"]))
    d["isagg"] = d["nominstallation"].astype(str).str.startswith("Agr")
    fil, tech, cc = d["filiere"].fillna(""), d["technologie"].fillna(""), d["codecombustible"].fillna("")
    f = pd.Series("Other", index=d.index)
    f[fil == "Nucléaire"] = "Nuclear"
    f[fil == "Hydraulique"] = "Hydro"
    f[(fil == "Hydraulique") & (tech == "Pompage turbinage")] = "PumpedStorage"
    f[fil == "Eolien"] = "Wind_Onshore"
    f[(fil == "Eolien") & tech.str.startswith("En mer")] = "Wind_Offshore"
    f[fil == "Solaire"] = "Solar"
    f[(fil == "Solaire") & (tech == "Batterie")] = "Storage"
    f[fil == "Bioénergies"] = "Bioenergy"
    f[fil == "Stockage non hydraulique"] = "Storage"
    th = fil.isin(["Thermique non renouvelable", "Autre"])
    f[th & (cc == "GAZ")] = "Gas"
    f[th & (cc == "CHARB")] = "Coal"
    f[th & (cc == "FIOUL")] = "Oil"
    f[(fil == "Autre") & (tech == "Batterie")] = "Storage"
    d["fuel"] = f
    return d


def fr_totals(d):
    """Edition totals by fuel class (all rows), MW."""
    by = d.groupby("fuel")["MW"].sum()
    out = {k: float(by.get(k, 0.0)) for k in FR_FUEL_ORDER}
    s = d[d["fuel"] == "Solar"]
    out["Solar_agg"] = float(s.loc[s["isagg"], "MW"].sum())
    out["Solar_dated"] = float(s.loc[~s["isagg"], "MW"].sum())
    out["rows"] = int(len(d))
    return out


def fr_rebuild(cur, dates, anchors):
    """Month-end capacity for the row dates `dates`; `anchors` = {asof_date: Solar_agg MW} of the editions."""
    rec = {}
    ax = np.array([(a + pd.Timedelta(days=1)).toordinal() for a in sorted(anchors)], dtype=float)
    ay = np.array([anchors[a] for a in sorted(anchors)], dtype=float)
    for r in dates:
        s = cur[cur["start"].isna() | (cur["start"] < r)]
        s = s[~((s["fuel"] == "Solar") & s["isagg"])]
        by = s.groupby("fuel")["MW"].sum()
        row = {k: float(by.get(k, 0.0)) for k in FR_FUEL_ORDER}
        agg = float(np.interp(r.toordinal(), ax, ay)) if len(ax) else 0.0
        row["Solar_dated"] = row["Solar"]
        row["Solar_agg_interpolated"] = agg
        row["Solar"] = row["Solar_dated"] + agg
        rec[r] = row
    df = pd.DataFrame.from_dict(rec, orient="index")
    df.index.name = "date"
    out = pd.DataFrame(index=df.index)
    out["Hydro_MW"] = df["Hydro"]
    out["Gas_MW"] = df["Gas"]
    out["Wind_MW"] = df["Wind_Onshore"] + df["Wind_Offshore"]
    out["Solar_MW"] = df["Solar"]
    out["Coal_MW"] = df["Coal"]
    out["Nuclear_MW"] = df["Nuclear"]
    out["Oil_MW"] = df["Oil"]
    out["Bioenergy_MW"] = df["Bioenergy"]
    out["Other_MW"] = df["Other"]
    out["PumpedStorage_MW"] = df["PumpedStorage"]
    out["Storage_MW"] = df["Storage"]
    extra = pd.DataFrame({"Wind_Onshore_MW": df["Wind_Onshore"], "Wind_Offshore_MW": df["Wind_Offshore"],
                          "Solar_dated_MW": df["Solar_dated"], "Solar_agg_interpolated_MW": df["Solar_agg_interpolated"]})
    std = standardise(out, STD)
    return pd.concat([std, extra.round(1)], axis=1)


SDES = "https://www.statistiques.developpement-durable.gouv.fr/"
QNAMES = ["premier", "deuxieme", "troisieme", "quatrieme"]


def sdes_find(kind, today):
    """Latest SDES quarterly dashboard page of `kind` ('solaire-photovoltaique' | 'eolien') -> (slug, media id)."""
    y, q = today.year, (today.month - 1) // 3          # current quarter index 0-3; try the one before first
    for _ in range(8):
        q -= 1
        if q < 0:
            q, y = 3, y - 1
        for suffix in ("", "-0"):
            slug = f"tableau-de-bord-{kind}-{QNAMES[q]}-trimestre-{y}{suffix}"
            r = SES.get(SDES + slug, timeout=60)
            if r.status_code != 200:
                continue
            m = re.search(r'href="/media/(\d+)/download[^"]*"[^>]*>\s*Donn', r.text)
            if not m:
                m = re.search(r'href="/media/(\d+)/download', r.text)
            if m:
                return slug, int(m.group(1))
    return None, None


def sdes_series(media_id, sheets):
    """National 'Total France' power (MW) per quarter end from an SDES quarterly workbook -> {sheet: Series}."""
    r = fetch(f"{SDES}media/{media_id}/download?inline")
    x = pd.ExcelFile(io.BytesIO(r.content))
    out = {}
    for sh in sheets:
        name = next((s for s in x.sheet_names if s.lower() == sh.lower()), None)
        if name is None:
            continue
        df = x.parse(name, header=None)
        dates, kinds = df.iloc[2, 3:].tolist(), df.iloc[3, 3:].tolist()
        row = df.index[df.iloc[:, 0].astype(str).str.strip().str.lower().isin(["total france"])]
        if not len(row):
            continue
        vals = df.iloc[row[0], 3:].tolist()
        s = {}
        for dt, k, v in zip(dates, kinds, vals):
            if str(k).strip().lower() == "puissance" and pd.notna(dt) and pd.notna(v):
                s[pd.Timestamp(dt)] = float(v)
        out[sh] = pd.Series(s).sort_index()
    return out, hashlib.sha256(r.content).hexdigest()


def run_fr(out_dir):
    path = os.path.join(out_dir, "france_power_capacity_monthly.xlsx")
    old = load_monthly(path)
    editions = read_sheet(path, "Editions", index_col=0)
    editions = editions if editions is not None else pd.DataFrame()
    if len(editions):
        editions.index = pd.to_datetime(editions.index)
    title, processed, asof = fr_meta(FR_REG)
    detail = f"{title}; processed {processed}"
    prev = stored_release(path, "ODRE register (current edition)")
    sd_prev = stored_release(path, "SDES quarterly")
    rel = []
    if asof is None:
        raise RuntimeError(f"cannot read the as-of date from the register title: {title}")

    # year-end editions (exact for their date): fetched once each and kept on 'Editions'
    for y in range(2020, NOW.year):
        a = pd.Timestamp(y, 12, 31)
        if a in editions.index:
            continue
        ds = f"{FR_REG}-3112{str(y)[2:]}"
        try:
            t2, p2, a2 = fr_meta(ds)
        except Exception as e:  # noqa: BLE001
            print(f"  edition {ds}: not available ({type(e).__name__})")
            continue
        d, sha = fr_load(ds)
        tot = fr_totals(fr_prepare(d))
        tot.update({"edition": t2, "kind": "year-end edition"})
        editions = pd.concat([editions, pd.DataFrame([tot], index=[a])])
        print(f"  edition {a.date()}: {tot['rows']} rows, solar aggregate {tot['Solar_agg']:.0f} MW")
    editions = editions.sort_index()

    changed = (prev != detail) or old.empty
    cur_tot = None
    if changed:
        d, sha = fr_load(FR_REG)
        cur = fr_prepare(d)
        cur_tot = fr_totals(cur)
        cur_tot.update({"edition": title, "kind": "monthly edition (as published)"})
        editions = pd.concat([editions[editions.index != asof], pd.DataFrame([cur_tot], index=[asof])]).sort_index()
        anchors = {a: float(editions.at[a, "Solar_agg"]) for a in editions.index if pd.notna(editions.at[a, "Solar_agg"])}
        last_row = (asof + pd.Timedelta(days=1)).replace(day=1)
        dates = pd.date_range(FIRST_ROW, last_row, freq="MS")
        monthly = fr_rebuild(cur, dates, anchors)
        rel.append(("ODRE register (current edition)", now_s(), detail, sha))
    else:
        print("  register unchanged since the last run; stored months kept")
        monthly = old
        rel.append(("ODRE register (current edition)", now_s(), detail, ""))

    # SDES quarterly national totals (validation, quarterly)
    sdes = read_sheet(path, "SDES_quarterly", index_col=0)
    sd_detail = sd_prev
    try:
        slug_s, mid_s = sdes_find("solaire-photovoltaique", NOW)
        slug_w, mid_w = sdes_find("eolien", NOW)
        if mid_s and mid_w:
            sd_detail = f"solar media {mid_s} ({slug_s}); wind media {mid_w} ({slug_w})"
            if sd_detail != sd_prev or sdes is None:
                ss, sha_s = sdes_series(mid_s, ["solaire PV"])
                ww, sha_w = sdes_series(mid_w, ["éolien_terrestre", "éolien_offshore"])
                q = pd.DataFrame({"Solar_MW": ss.get("solaire PV"), "Wind_Onshore_MW": ww.get("éolien_terrestre"),
                                  "Wind_Offshore_MW": ww.get("éolien_offshore")})
                q["Wind_MW"] = q["Wind_Onshore_MW"] + q["Wind_Offshore_MW"].fillna(0)
                q = q[q.index >= "2020-12-31"]
                q.index = q.index + pd.Timedelta(days=1)          # quarter-end -> row date (1st of next month)
                q.index.name = "date"
                sdes = q.round(1)
                rel.append(("SDES quarterly", now_s(), sd_detail, sha_s + sha_w))
            else:
                rel.append(("SDES quarterly", now_s(), sd_detail, ""))
        else:
            print("  SDES quarterly pages not found; stored table kept")
    except Exception as e:  # noqa: BLE001
        print(f"  SDES step failed: {type(e).__name__}: {e}")
    if sdes is not None and len(sdes):
        sdes.index = pd.to_datetime(sdes.index)
        sdes.index.name = "date"

    # checks
    chk = entsoe_check(out_dir, "france", monthly, ENTSOE_MAP)
    chk_ed = []
    for a, row in editions.iterrows():
        r = a + pd.Timedelta(days=1)
        if r not in monthly.index:
            continue
        for ours, k in (("Hydro_MW", "Hydro"), ("PumpedStorage_MW", "PumpedStorage"), ("Wind_MW", None), ("Solar_MW", "Solar"), ("Nuclear_MW", "Nuclear"),
                        ("Gas_MW", "Gas"), ("Coal_MW", "Coal"), ("Oil_MW", "Oil"), ("Bioenergy_MW", "Bioenergy"), ("Storage_MW", "Storage"), ("Other_MW", "Other")):
            ed = float(row["Wind_Onshore"] + row["Wind_Offshore"]) if k is None else float(row[k])
            mine = float(monthly.at[r, ours])
            chk_ed.append({"date": r, "series": ours, "rebuilt_MW": round(mine, 1), "edition_MW": round(ed, 1), "rebuilt_minus_edition_MW": round(mine - ed, 1),
                           "difference_pct": round((mine - ed) / ed * 100, 1) if ed else np.nan, "edition": row["edition"]})
    chk_ed = pd.DataFrame(chk_ed).set_index("date") if chk_ed else pd.DataFrame()
    chk_sd = []
    if sdes is not None and len(sdes):
        for r, row in sdes.iterrows():
            if r not in monthly.index:
                continue
            for col in ("Solar_MW", "Wind_MW", "Wind_Onshore_MW", "Wind_Offshore_MW"):
                a, b = monthly.at[r, col], row.get(col)
                if pd.isna(a) or pd.isna(b) or not b:
                    continue
                chk_sd.append({"date": r, "series": col, "this_file_MW": round(float(a), 1), "SDES_MW": round(float(b), 1),
                               "difference_MW": round(float(a) - float(b), 1), "difference_pct": round((float(a) - float(b)) / float(b) * 100, 1)})
    chk_sd = pd.DataFrame(chk_sd).set_index("date") if chk_sd else pd.DataFrame()

    notes = ["France - monthly installed generation capacity by fuel, MW, rebuilt from the national production register.",
             "",
             "Source", "ODRE (RTE, Enedis, ADEeF and the other network operators) 'Registre national des installations de production et de stockage d'electricite': " + ODRE + FR_REG,
             "Row dated 1 Mar = capacity at the end of February: the sum of the installations of the CURRENT edition commissioned before that date (same date convention as the ENTSO-E annual files). Power in the register is kW; this file is MW.",
             "", "Limits",
             "Installations closed since are not in the current edition, so earlier months miss them (largest for thermal: see Check_Editions against the year-end editions).",
             "About 226 installations have no date (old EDF hydro and thermal plant, mostly): they count as in service in every month.",
             "PV below 36 kW is aggregated per commune under one early date: its growth cannot be dated from the current edition. Solar_agg_interpolated_MW is the exact aggregate of each published edition (year-end 2020-25 and the current one) interpolated LINEARLY in time between editions; Solar_dated_MW is the dated (larger) installations. Solar_MW is their sum. The aggregate is 18% of solar.",
             "Hydro_MW excludes pumped storage (technology 'Pompage turbinage' = PumpedStorage_MW). Storage_MW = non-hydraulic storage (batteries). Other_MW = geothermal, marine (La Rance), waste and thermal plant with no fuel code. Wind_Onshore_MW / Wind_Offshore_MW split Wind_MW. Total_MW excludes pumped storage and batteries.",
             "SDES_quarterly = national totals (MW) of the SDES 'tableau de bord' solar and wind workbooks at each quarter end (installations connected to the grid); Check_SDES compares them with the rebuilt months.",
             "", "Method",
             "The register edition is replaced monthly (sheet Releases holds its title and processing time); the whole series is rebuilt when it changes. 'Editions' keeps the totals of every edition seen (the year-end editions are fetched once each), so the real-time view is not lost when history is revised."]
    sheets = {"Monthly": monthly}
    if len(chk):
        sheets["Check_ENTSOE"] = chk
    if len(chk_ed):
        sheets["Check_Editions"] = chk_ed
    if sdes is not None and len(sdes):
        sheets["SDES_quarterly"] = sdes
    if len(chk_sd):
        sheets["Check_SDES"] = chk_sd
    sheets["Editions"] = editions
    sheets["Releases"] = releases_frame(rel)
    write(path, sheets, notes, ["Source", "Limits", "Method"])
    print(f"France: {len(monthly)} months {monthly.index.min().date()} - {monthly.index.max().date()}; register {title}")
    return monthly


# ------------------------------------------------------------------------------------------------ Spain
ES_URL = "https://apidatos.ree.es/es/datos/generacion/potencia-instalada-generacion"
ES_MAP = {"Hidráulica": "Hydro_MW", "Nuclear": "Nuclear_MW", "Carbón": "Coal_MW", "Ciclo combinado": "Combined_cycle_MW",
          "Turbina de gas": "Gas_turbine_MW", "Fuel + Gas": "FuelGas_MW", "Motores diésel": "Diesel_engines_MW", "Turbina de vapor": "Steam_turbine_MW",
          "Hidroeólica": "Hydrowind_MW", "Eólica": "Wind_MW", "Solar fotovoltaica": "Solar_PV_MW", "Solar térmica": "Solar_thermal_MW",
          "Otras renovables": "Other_renewables_MW", "Cogeneración": "Cogeneration_MW", "Residuos no renovables": "Waste_nonrenewable_MW",
          "Residuos renovables": "Waste_renewable_MW", "Potencia instalada total": "Total_REE_MW"}


def es_get(a, b, trunc="month"):
    r = fetch(ES_URL, params={"start_date": f"{a}T00:00", "end_date": f"{b}T23:59", "time_trunc": trunc}, headers={"Accept": "application/json"})
    j = r.json()
    if "included" not in j:
        raise RuntimeError(f"REData {a}..{b}: {str(j)[:200]}")
    out = {}
    for it in j["included"]:
        col = ES_MAP.get(it["attributes"].get("title"))
        if col is None:
            print(f"  unmapped REData series {it['attributes'].get('title')!r}")
            continue
        for v in it["attributes"].get("values", []):
            out.setdefault(pd.Timestamp(v["datetime"][:10]), {})[col] = v["value"]
    return out, j["data"]["attributes"].get("last-update")


def run_es(out_dir):
    path = os.path.join(out_dir, "spain_power_capacity_monthly.xlsx")
    old = load_monthly(path)
    last_period = None if old.empty else old.index.max() - pd.offsets.MonthBegin(1)   # last month label stored
    today = pd.Timestamp(NOW.year, NOW.month, 1)
    years = list(range(2021, NOW.year + 1)) if last_period is None else list(range(max(2021, last_period.year - 1), NOW.year + 1))
    data, upd = {}, None
    for y in years:
        end = f"{y}-12-31"
        if y == NOW.year:
            end = (today + pd.offsets.MonthEnd(0)).strftime("%Y-%m-%d")
        try:
            d, upd = es_get(f"{y}-01-01", end)
        except RuntimeError as e:
            if y != NOW.year:
                raise
            print(f"  {e}; retrying up to the previous month")
            d, upd = es_get(f"{y}-01-01", (today - pd.Timedelta(days=1)).strftime("%Y-%m-%d"))
        data.update(d)
    if old.empty or FIRST_ROW not in old.index:
        d, upd = es_get("2020-12-01", "2020-12-31")
        data.update(d)
    df = pd.DataFrame.from_dict(data, orient="index").sort_index()
    df.index = df.index + pd.offsets.MonthBegin(1)      # the month's value is the stock at its end: row = 1st of the next month
    df.index.name = "date"
    df = df[df["Total_REE_MW"].notna()]
    new = pd.DataFrame(index=df.index)
    new["Hydro_MW"] = df["Hydro_MW"]
    new["Gas_MW"] = df[["Combined_cycle_MW", "Gas_turbine_MW"]].sum(axis=1, min_count=1)
    new["Wind_MW"] = df["Wind_MW"]
    new["Solar_MW"] = df[["Solar_PV_MW", "Solar_thermal_MW"]].sum(axis=1, min_count=1)
    new["Coal_MW"] = df["Coal_MW"]
    new["Nuclear_MW"] = df["Nuclear_MW"]
    new["Oil_MW"] = df[["Diesel_engines_MW", "Steam_turbine_MW", "FuelGas_MW"]].sum(axis=1, min_count=1)
    new["Bioenergy_MW"] = df[["Other_renewables_MW", "Waste_renewable_MW"]].sum(axis=1, min_count=1)
    new["Other_MW"] = df[["Cogeneration_MW", "Waste_nonrenewable_MW", "Hydrowind_MW"]].sum(axis=1, min_count=1)
    new["PumpedStorage_MW"] = np.nan
    new["Storage_MW"] = np.nan
    new = standardise(new, STD)
    extra = df[["Solar_PV_MW", "Solar_thermal_MW", "Combined_cycle_MW", "Cogeneration_MW", "Total_REE_MW"]].round(1)
    new = pd.concat([new, extra], axis=1)
    monthly = merge_rows(old, new)
    monthly = monthly[monthly.index >= FIRST_ROW]
    # year-end semantics check: REData's own yearly value against the December monthly value
    yr = []
    for y in range(2021, NOW.year):
        try:
            d, _ = es_get(f"{y}-01-01", f"{y}-12-31", "year")
        except Exception as e:  # noqa: BLE001
            print(f"  yearly check {y}: {e}")
            continue
        for k, v in d.items():
            r = pd.Timestamp(y + 1, 1, 1)
            for col in ("Total_REE_MW", "Wind_MW", "Solar_PV_MW", "Hydro_MW"):
                if col in v and r in monthly.index and col in monthly:
                    yr.append({"date": r, "series": col, "REData_year_value_MW": round(v[col], 1), "REData_December_month_MW": round(float(monthly.at[r, col]), 1),
                               "difference_MW": round(v[col] - float(monthly.at[r, col]), 1)})
    chk_y = pd.DataFrame(yr).set_index("date") if yr else pd.DataFrame()
    chk = entsoe_check(out_dir, "spain", monthly, {"Hydro_MW": "Hydro_MW", "Gas_MW": "Gas_MW", "Wind_MW": "Wind_MW", "Solar_MW": "Solar_MW", "Coal_MW": "Coal_MW",
                                                  "Nuclear_MW": "Nuclear_MW", "Oil_MW": "Oil_MW", "Bioenergy_MW": "Bioenergy_MW", "Total_MW": "Total_MW"})
    notes = ["Spain - monthly installed generation capacity by technology, MW, national electricity system (peninsular + non-peninsular), from Red Electrica's REData API.",
             "",
             "Source", "Red Electrica de Espana (the TSO) REData API, widget generacion/potencia-instalada-generacion ('Estructura de potencia instalada'): " + ES_URL + "?start_date=..&end_date=..&time_trunc=month (the earlier name generacion/potencia-instalada answers HTTP 500). Free, no key.",
             "Row dated 1 Mar = capacity at the end of February (REData labels each value with its month; the sheet Check_REData_year shows that REData's yearly value equals the December monthly value, i.e. the monthly value is the stock at the end of the month).",
             "", "Mapping and limits",
             "Hydro_MW = 'Hidraulica' (pumped storage, 3.3 GW, is NOT in this widget: PumpedStorage_MW is blank; REE reports it annually). Gas_MW = combined cycle + gas turbines; Coal_MW = coal; Oil_MW = diesel engines + steam turbines + fuel/gas (the non-peninsular thermal units); Solar_MW = photovoltaic + solar thermal (REE's photovoltaic includes self-consumption, so it exceeds ENTSO-E's A68 solar); Bioenergy_MW = 'otras renovables' (biogas, biomass, marine, geothermal) + renewable waste; Other_MW = cogeneration (mostly gas, kept separate in Cogeneration_MW) + non-renewable waste + hydro-wind. Total_MW = REE's 'Potencia instalada total' (excludes pumped storage).",
             "A call with a range of more than about 13 months is refused (HTTP 400 'datos no disponibles'): the pull asks one calendar year at a time; the last two years are re-read every run (values are revised), earlier months are kept."]
    sheets = {"Monthly": monthly}
    if len(chk):
        sheets["Check_ENTSOE"] = chk
    if len(chk_y):
        sheets["Check_REData_year"] = chk_y
    sheets["Releases"] = releases_frame([("REData potencia-instalada-generacion", now_s(), f"last-update {upd}", "")])
    write(path, sheets, notes, ["Source", "Mapping and limits"])
    print(f"Spain: {len(monthly)} months {monthly.index.min().date()} - {monthly.index.max().date()} (REData last-update {upd})")
    return monthly


# ------------------------------------------------------------------------------------------------ Portugal
PT_PAGE = "https://www.dgeg.gov.pt/pt/estatistica/energia/publicacoes/estatisticas-rapidas-das-renovaveis/"
PT_LABELS = [("Total Renovável", "Renewables_total_MW"), ("Hídrica", "Hydro_incl_pumped_MW"), ("da qual em bombagem", "PumpedStorage_MW"),
             ("Eólica", "Wind_MW"), ("Biomassa", "Biomass_MW"), ("Resíduos Sólidos Urbanos", "Waste_MW"), ("Biogás", "Biogas_MW"),
             ("Geotérmica", "Geothermal_MW"), ("Fotovoltaica", "Solar_MW")]
PT_MONTHS = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]


def pt_pdf_links():
    """{(year, month): url} for the monthly PDFs listed on the DGEG page, and {year: url} for the year-end ones."""
    h = fetch(PT_PAGE).text
    monthly = {}
    for m in re.finditer(r'href="(/media/[a-z0-9]+/[^"]*?(?:dgeg-arr-|renovaveis-)(\d{4})-?(\d{2})[^"]*?\.pdf)"', h):
        monthly[(int(m.group(2)), int(m.group(3)))] = "https://www.dgeg.gov.pt" + m.group(1)
    return monthly


def pt_parse(content):
    """-> (years header list, {column: [values per year]}) from the 'Potencia instalada' table of a DGEG renewables PDF."""
    import pymupdf
    doc = pymupdf.open(stream=content, filetype="pdf")
    for pg in doc:
        t = pg.get_text()
        if not (re.search(r"pot[eê]ncia\s+instalada", t, re.I) and "Total Renov" in t and "Eólica" in t):
            continue
        toks = [x.strip() for x in t.split("\n") if x.strip()]
        # header years: the first run of 10 consecutive four-digit tokens after 'MW' / 'Potência Instalada (MW)'
        yrs = None
        for i in range(len(toks) - 9):
            run = toks[i:i + 10]
            if all(re.fullmatch(r"(19|20)\d{2}", x) for x in run):
                yrs = [int(x) for x in run]
                start = i + 10
                break
        if yrs is None:
            continue
        n = len(yrs)
        out = {}
        for label, col in PT_LABELS:
            for i in range(start, len(toks)):
                if toks[i] == label:
                    vals, k = [], i + 1
                    while k < len(toks) and len(vals) < n and re.fullmatch(r"-|\d{1,3}(?: \d{3})*", toks[k]):
                        vals.append(np.nan if toks[k] == "-" else float(toks[k].replace(" ", "")))
                        k += 1
                    if len(vals) == n:
                        out[col] = vals
                    break
            if col not in out and col != "PumpedStorage_MW":
                pass
        if "Hydro_incl_pumped_MW" in out and "Wind_MW" in out and "Solar_MW" in out:
            return yrs, out
    return None, None


def run_pt(out_dir):
    path = os.path.join(out_dir, "portugal_power_capacity_monthly.xlsx")
    old = load_monthly(path)
    links = pt_pdf_links()
    if not links:
        raise RuntimeError("no DGEG monthly PDFs found on the page")
    rel_old = read_sheet(path, "Releases", index_col=0)
    seen = {} if rel_old is None else rel_old["detail"].astype(str).to_dict()
    rel = []
    rows = {}
    ye_latest = None
    newest = max(links)
    for (y, m), url in sorted(links.items()):
        key = f"DGEG {y}-{m:02d}"
        recheck = (y, m) >= (newest[0], newest[1] - 1) if newest[1] > 1 else (y, m) >= (newest[0] - 1, 12)
        if key in seen and seen[key] == url and not recheck:
            rel.append((key, "", url, ""))
            continue
        try:
            r = fetch(url)
            yrs, tab = pt_parse(r.content)
        except Exception as e:  # noqa: BLE001
            print(f"  {key}: {type(e).__name__}: {e}")
            continue
        if not tab:
            print(f"  {key}: table not found")
            continue
        sha = hashlib.sha256(r.content).hexdigest()
        rel.append((key, now_s(), url, sha))
        rows[(y, m)] = (yrs, tab)
    # stored series
    data = {}
    if not old.empty:
        for d, row in old.iterrows():
            data[d] = row.to_dict()
    for (y, m), (yrs, tab) in rows.items():
        # current month = last column (report month)
        rd = (pd.Timestamp(y, m, 1) + pd.offsets.MonthBegin(1))
        rec = {c: v[-1] for c, v in tab.items()}
        rec["source"] = f"DGEG report {y}-{m:02d}"
        data[rd] = {**data.get(rd, {}), **rec}
    # year-end values of earlier years from the newest PDF that was read (or the newest stored): Dec of each year -> row 1 Jan next year
    if rows:
        (yy, mm) = max(rows)
        yrs, tab = rows[(yy, mm)]
        for j, yv in enumerate(yrs[:-1]):
            if yv < 2020:
                continue
            rd = pd.Timestamp(yv + 1, 1, 1)
            rec = {c: v[j] for c, v in tab.items()}
            rec["source"] = f"DGEG report {yy}-{mm:02d} (31 Dec {yv})"
            data[rd] = {**data.get(rd, {}), **rec}
    df = pd.DataFrame.from_dict(data, orient="index").sort_index()
    df.index.name = "date"
    df["Hydro_MW"] = df["Hydro_incl_pumped_MW"] - df["PumpedStorage_MW"].fillna(0)
    df["Bioenergy_MW"] = df[["Biomass_MW", "Waste_MW", "Biogas_MW"]].sum(axis=1, min_count=1)
    df["Other_MW"] = df["Geothermal_MW"]
    mon = standardise(df, STD)
    # Total_MW here = renewable fleet only (no gas/coal/oil in the source)
    mon["Total_MW"] = (df["Renewables_total_MW"] - df["PumpedStorage_MW"].fillna(0)).round(1)
    keep = ["Hydro_incl_pumped_MW", "Biomass_MW", "Waste_MW", "Biogas_MW", "Geothermal_MW", "Renewables_total_MW", "source"]
    mon = pd.concat([mon, df[[c for c in keep if c in df]]], axis=1)
    mon["Gas_MW"] = np.nan
    mon["Coal_MW"] = np.nan
    mon["Oil_MW"] = np.nan
    mon["Nuclear_MW"] = 0.0
    mon["Storage_MW"] = np.nan
    mon = mon[mon.index >= FIRST_ROW]
    chk = entsoe_check(out_dir, "portugal", mon, {"Hydro_MW": "Hydro_MW", "PumpedStorage_MW": "PumpedStorage_MW", "Wind_MW": "Wind_MW", "Solar_MW": "Solar_MW",
                                                  "Bioenergy_MW": "Bioenergy_MW"})
    notes = ["Portugal - monthly installed capacity of the RENEWABLE fleet, MW, from DGEG's 'Estatisticas Rapidas das Renovaveis'. PARTIAL: no gas, coal or oil.",
             "",
             "Source", "DGEG (Direcao-Geral de Energia e Geologia) monthly PDF 'Estatisticas Rapidas das Renovaveis', table 'Potencia instalada' (MW): " + PT_PAGE,
             "Row dated 1 Mar = capacity at the end of February (same date convention as the ENTSO-E annual files). Monthly from January 2024 (the PDF's last column is the report month, e.g. 'agosto de 2026'); before that only 31 December of each year (rows dated 1 January), taken from the latest report read, so revised values replace the originals; the months in between are absent, not interpolated.",
             "", "Limits",
             "Renewables only. DGEG's Solar_MW includes self-consumption (UPAC), small-production (UPP) and micro/mini units, so it is about twice ENTSO-E's A68 figure (see Check_ENTSOE). Hydro_MW = DGEG hydro less the pumped-storage part ('da qual em bombagem', published from the 2022 reports; earlier years keep it inside Hydro_MW). Bioenergy_MW = biomass + municipal solid waste + biogas; Other_MW = geothermal (Azores). Total_MW is the DGEG renewable total, NOT a national total.",
             "Gas_MW, Coal_MW, Oil_MW and Storage_MW are blank: REN and DGEG publish thermal capacity annually only (DGEG 'Potencia instalada nas centrais produtoras de eletricidade'), and no monthly raw source was found (probes in discovery_archive/europe/capacity/). Nuclear_MW = 0."]
    sheets = {"Monthly": mon}
    if len(chk):
        sheets["Check_ENTSOE"] = chk
    rel_old_rows = [(k, "", v, "") for k, v in seen.items() if k not in {r[0] for r in rel}]
    sheets["Releases"] = releases_frame(rel_old_rows + rel)
    write(path, sheets, notes, ["Source", "Limits"])
    print(f"Portugal: {len(mon)} rows {mon.index.min().date()} - {mon.index.max().date()}")
    return mon


# ------------------------------------------------------------------------------------------------ Italy
def run_it(out_dir):
    print("  Italy: not built yet")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--only", default="fr,es,pt,it")
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)
    failed = []
    for key, fn in (("fr", run_fr), ("es", run_es), ("pt", run_pt), ("it", run_it)):
        if key not in a.only.split(","):
            continue
        try:
            print(f"== {key}")
            fn(a.out_dir)
        except Exception:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            failed.append(key)
    if failed:
        print("FAILED:", failed)
        sys.exit(1)


if __name__ == "__main__":
    main()
