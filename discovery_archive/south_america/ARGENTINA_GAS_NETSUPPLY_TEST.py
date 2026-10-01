"""
Live test of the net-supply code in south_america/ARGENTINA_GAS.py before it goes into the
scheduled pull (see ARGENTINA_GAS_NETSUPPLY_DISCOVERY*.py for how the sources were found).

  - parses the daily transport report PDF for the 1st and 15th of every month 2021..2026
    (layout changes, parse failures, footnote values) and compares the footnote imports with
    the daily import reports for the same days;
  - fetches GRT, GETD and CLP and the import reports, builds 'Supply net' with the sampled
    days, and compares GRT's domestic injection (mcm/d) with the sampled daily reports - in
    particular whether GRT's 2026 TGN figures hold Escobar LNG (GRT 'Otros Origenes' is 0);
  - prints the balance check against the workbook's sector and export series.
Nothing is written to the repo.

Found (Oct 2026): 134 of 138 sampled days parse (4 days have no report). The report's file date
was the END of the 06:00-06:00 gas day in 2021-2024 and is its START in 2026, so rows are keyed on
the 'Periodo' line; keyed that way, the footnote imports equal the daily import reports to 0.01
mcm/d. About 20% of reports print the (a) Bolivia footnote with overlapping characters that
pdfplumber cannot read; the import report fills them. GRT's domestic injection (TGN+TGS, imports
taken out) matches the sampled daily reports within sampling noise; in Jun-Jul 2026 GRT is ~2 mcm/d
below them, so GRT's 2026 figures do not hold Escobar LNG (its 'Otros Origenes' is 0 in 2026). The
Secretaria de Energia sector series equal ENARGAS GETD (max differences 1-17 million m3/month).
"""
import os
import sys
import time

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "south_america"))
import ARGENTINA_GAS as A  # noqa: E402

pd.set_option("display.width", 300)
pd.set_option("display.max_columns", 40)
pd.set_option("display.max_rows", 400)
BOOK = os.path.join(ROOT, "output", "Data and Chart Outputs", "argentina_gas_monthly.xlsx")


def main():
    s = A.enargas_session()
    s.get(A.ENARGAS + "dod-partes-dist-trans.php", timeout=A.TIMEOUT)
    days = [d for m in pd.date_range("2021-01-01", "2026-09-01", freq="MS") for d in (m, m + pd.Timedelta(days=14))]
    recs, t0 = {}, time.time()
    for d in days:
        try:
            rec = A.fetch_transport_day(s, d)
        except Exception as e:  # noqa: BLE001
            print(d.date(), "ERROR", type(e).__name__, e)
            continue
        recs[d] = rec or {"Injection_total": None}
    t = pd.DataFrame.from_dict(recs, orient="index")
    t = t.apply(lambda c: pd.to_datetime(c, errors="coerce") if c.name == "Gas_day" else pd.to_numeric(c, errors="coerce"))
    print(f"transport sample: {len(t)} days in {time.time() - t0:.0f}s; missing/unparsed: "
          f"{[str(x.date()) for x in t.index[t['Injection_total'].isna()]]}")
    print("missing fields per column:", t.isna().sum().to_dict())
    print(t.round(2).to_string())
    recs_frame = t.copy()

    # text of reports whose Bolivia footnote did not parse
    import io
    import pdfplumber
    for d in list(t.index[t["Incl_Bolivia_NorAndino"].isna() & t["Injection_total"].notna()])[:2]:
        r = s.get(A.PARTE_URL, timeout=A.TIMEOUT, params={"tipo": "transporte", "path": "partes-diarios/transporte",
                                                          "file": d.strftime("%Y%m%d") + ".pdf"})
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            print(f"\n--- unparsed footnote, {d.date()} ---\n{pdf.pages[0].extract_text()}")
    imp = A.fetch_imports_daily(pd.Timestamp("2021-01-01"), pd.Timestamp.today().normalize())
    tg = A.by_gas_day(t)
    k = imp.reindex(tg.index) / 1000
    t = tg
    cmp = pd.DataFrame({"parte_Bol_NorAnd": t["Incl_Bolivia_NorAndino"],
                        "imp_Bol_NorAnd": k["Bolivia"] + k["Norandino"],
                        "parte_Esc_GasAnd": t["Incl_Escobar_GasAndes"], "imp_Esc_GasAnd": k["GNL Escobar"] + k["Gasandes"],
                        "parte_BB": t["Incl_LNG_BahiaBlanca"], "imp_BB": k["GNL B. Blanca"]})
    print("\nfootnote imports vs import reports (mcm/d), same days:")
    print(cmp.round(2).to_string())

    imports_m = A.monthly_imports(imp)
    grt = A.fetch_grt("2021-01-01")
    getd = A.fetch_getd("2021-01-01")
    clp = A.fetch_clp("2021-01-01")
    exports_daily = A.load_exports_daily(BOOK)
    exports, exports_points = A.monthly_exports(exports_daily)
    net = A.supply_net(grt, getd, exports_points, imports_m, clp, pd.DataFrame())
    sd_full = A.supply_daily(recs_frame, imp)
    print("\nSupply daily (sample days):")
    print(sd_full.dropna(subset=["Domestic_injection_mcmd"]).iloc[:, :14].to_string())
    print("\nSupply net:")
    print(net.to_string())

    sd = sd_full
    samp = sd["Domestic_injection_mcmd"].groupby(sd.index.to_period("M")).mean()
    samp.index = samp.index.to_timestamp()
    g = grt / 1000
    tgn_tgs = (net["Domestic_injection"] - net["Domestic_distributor_pipelines"]) / net.index.days_in_month
    v = pd.DataFrame({"GRT_dom_TGN_TGS_mcmd": tgn_tgs, "parte_dom_sample_mcmd": samp.reindex(net.index),
                      "GRT_TGN_Neuquina_mcmd": A.col_like(g, "TGN", "Neuquina").reindex(net.index) / net.index.days_in_month,
                      "LNG_Escobar_mcmd": net["LNG_Escobar"] / net.index.days_in_month,
                      "GRT_LNG_reported_mcmd": net["GRT_LNG_other_origins_reported"] / net.index.days_in_month})
    v["diff"] = v["GRT_dom_TGN_TGS_mcmd"] - v["parte_dom_sample_mcmd"]
    print("\nGRT vs sampled daily transport reports (mcm/d):")
    print(v.round(2).to_string())

    nat = pd.read_excel(BOOK, sheet_name="National", index_col=0)
    nat.index = pd.to_datetime(nat.index)
    print("\nBalance check:")
    print(A.balance_check(nat, getd, exports, net).to_string())
    g2 = getd.reindex(nat.index) / 1000
    print("\nSE sectors vs GETD (million m3), max abs diff by sector:")
    for se, keys in (("residencial", ("Residencial",)), ("industria", ("Industria",)), ("centrales_electricas", ("Centrales",)),
                     ("gnc", ("GNC",)), ("comercial", ("Comercial",)), ("entes_oficiales", ("Entes",))):
        print(f"  {se}: {(nat[se] - A.col_like(g2, *keys)).abs().max():.2f}")


if __name__ == "__main__":
    main()
