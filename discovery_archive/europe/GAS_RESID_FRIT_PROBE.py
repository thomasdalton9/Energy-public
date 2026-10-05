"""Probe (EU27+UK gas residual, France and Italy): ODRE catalogue (gas datasets) with field names of the promising ones, reachability of Italian sources
(MASE/DGSAIE gas balance, Snam, ARERA), and ENTSOG point-level monthly flows 2025-01..2026-09 for FR, IT, ES, UK. Writes results/residual/."""
import json, os, re, sys, time
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "europe"))
import requests, pandas as pd
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results", "residual")
os.makedirs(OUT, exist_ok=True)
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
which = sys.argv[1:] or ["odre", "it", "entsog"]
if "odre" in which:
    B = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets"
    off, ds = 0, []
    while off < 3000:
        r = requests.get(B, params={"limit": 100, "offset": off, "where": "search(title,'gaz') OR search(title,'gas') OR search(title,'GRTgaz') OR search(title,'Teréga') OR search(title,'biométhane') OR search(title,'GNL')"}, timeout=90)
        if not r.ok: print("odre", r.status_code, r.text[:150]); break
        res = r.json().get("results", [])
        if not res: break
        ds += res; off += 100
    lines = []
    for d in ds:
        t = (d.get("metas", {}).get("default", {}) or {})
        lines.append(f"{d['dataset_id']}\t{t.get('title','')}\t{t.get('records_count','')}\t{(t.get('modified') or '')[:10]}")
    open(os.path.join(OUT, "odre_catalogue.tsv"), "w").write("\n".join(lines)); print("odre datasets", len(ds), flush=True)
    pat = re.compile(r"perte|propre|bilan|compress|stock|linepack|flux|physique|interconnexion|export|import|transit|biom|injection|production|gnl|terminal|consommation|livraison|quantit|ecart|écart|fuel", re.I)
    fields = {}
    for d in ds:
        t = (d.get("metas", {}).get("default", {}) or {}).get("title", "")
        if pat.search(t) or pat.search(d["dataset_id"]):
            try:
                r = requests.get(f"{B}/{d['dataset_id']}/records", params={"limit": 2, "order_by": ""}, timeout=60)
                j = r.json() if r.ok else {}
                fields[d["dataset_id"]] = {"title": t, "sample": j.get("results", [])[:2], "status": r.status_code}
            except Exception as e:
                fields[d["dataset_id"]] = {"title": t, "err": type(e).__name__}
    json.dump(fields, open(os.path.join(OUT, "odre_fields.json"), "w"), ensure_ascii=False, indent=1, default=str)
    print("odre promising", len(fields), flush=True)
if "it" in which:
    res = {}
    urls = ["https://dgsaie.mise.gov.it/", "https://dgsaie.mise.gov.it/gas_naturale.php", "https://dgsaie.mise.gov.it/dgsaie_gas_naturale.php", "https://dgsaie.mise.gov.it/bilancio_gas.php",
            "https://dgsaie.mise.gov.it/importazioni_gas.php", "https://www.mase.gov.it/energia/statistiche-energetiche", "https://www.mase.gov.it/energia/gas-naturale",
            "https://www.mase.gov.it/portale/web/guest/gas-naturale", "https://www.snam.it/it/trasporto/dati-operativi/", "https://www.snam.it/en/transportation/operational-data-business/",
            "https://jarvis.snam.it/", "https://www.snam.it/it/home/", "https://www.snam.it/en/transport/transparency/", "https://www.arera.it/dati-e-statistiche",
            "https://www.arera.it/it/dati/gas.htm", "https://www.terna.it/en", "https://www.gse.it", "https://dati.mase.gov.it/", "https://www.mase.gov.it/portale/bilancio-gas-naturale"]
    for u in urls:
        try:
            r = requests.get(u, headers=UA, timeout=40, allow_redirects=True)
            links = sorted(set(re.findall(r'href="([^"]*(?:gas|bilancio|balance|dati|consum)[^"]*)"', r.text, re.I)))[:60]
            res[u] = {"status": r.status_code, "final": r.url, "len": len(r.content), "links": links}
        except Exception as e:
            res[u] = {"err": type(e).__name__}
        print(u, res[u].get("status", res[u].get("err")), flush=True)
    json.dump(res, open(os.path.join(OUT, "it_sources.json"), "w"), indent=1)
if "entsog" in which:
    import ENTSOG_GAS_FLOWS_DAILY as E
    adj = E.adjacency()
    labels = {}
    rows = []
    for p in pd.period_range("2025-01", "2026-09", freq="M"):
        d0, d1 = p.start_time.date(), p.end_time.date()
        try:
            data = E.get_json("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": d0.isoformat(), "to": d1.isoformat(), "limit": -1}).get("operationalData", [])
        except Exception as e:
            print("fail", p, e); continue
        agg = {}
        for r in data:
            op = r.get("operatorKey") or ""
            if op[:2] not in ("FR", "IT", "ES", "UK", "PT"): continue
            v = r.get("value")
            if v in (None, ""): continue
            try: g = float(v) / 1e6 if r.get("unit", "kWh/d") == "kWh/d" else float("nan")
            except Exception: continue
            a = adj.get((r["pointKey"], op, r["directionKey"]), (None, None))
            k = (op, r["pointKey"], r.get("pointLabel"), r["directionKey"], a[0], a[1])
            agg[k] = agg.get(k, 0.0) + g
        for k, g in agg.items():
            rows.append((str(p),) + k + (round(g / 1000.0, 4),))
        print("entsog", p, len(data), flush=True)
    pd.DataFrame(rows, columns=["month", "operator", "pointKey", "label", "dir", "adjType", "adjCountry", "TWh"]).to_csv(os.path.join(OUT, "entsog_points_monthly.csv"), index=False)
