"""Probe 3 (gas residual): (1) GB System Entry Energy per terminal (D+2 and M+15 allocations), monthly sums 2025-01..2026-09; (2) ODRE: GHG emissions by operator,
methane emissions, monthly biomethane injected into the transmission network, LNG terminal stocks; (3) Italian sources: Snam / SISEN-MASE page crawl for gas-balance files.
Writes results/residual/."""
import json, os, re, sys, time
import requests, pandas as pd
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results", "residual")
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
which = sys.argv[1:] or ["gb", "odre", "it"]
if "gb" in which:
    B = "https://data.nationalgas.com"
    H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json, text/plain, */*", "Referer": "https://data.nationalgas.com/find-gas-data/view", "Content-Type": "application/json"}
    cat = open(os.path.join(OUT, "gb_catalogue.txt")).read().split("\n")
    ids = {}
    for ln in cat:
        k, _, v = ln.partition("\t")
        if v.startswith("System Entry Energy") and ("D+2" in v.split("||")[0] or v.split("||")[0].strip().endswith("M+15")) and "test" not in v:
            ids[k] = v.split("||")[0].strip()
    ids.update({"PUBOBJ1278": "Aggregate Physical Energy, Subterminal, M+15", "PUBOBJ1279": "Aggregate Physical Energy, Interconnector M+15",
                "PUBOBJ1280": "Aggregate Physical Energy, Storage Withdrawal, M+15", "PUBOBJ1281": "Aggregate Physical Energy, LNG Importation, M+15"})
    print("gb ids", len(ids), flush=True)
    months = pd.period_range("2025-01", "2026-09", freq="M")
    rows, keys = [], list(ids)
    path = os.path.join(OUT, "gb_entry_energy_monthly.csv")
    for i in range(0, len(keys), 6):
        sub = keys[i:i + 6]
        for q0 in range(0, len(months), 4):
            ms = months[q0:q0 + 4]
            body = {"latestFlag": "Y", "applicableFor": "Y", "dateFrom": ms[0].start_time.date().isoformat(), "dateTo": ms[-1].end_time.date().isoformat(), "dateType": "GASDAY", "ids": ",".join(sub)}
            for a in range(3):
                try:
                    r = requests.post(B + "/api/find-gas-data", json=body, headers=H, timeout=120)
                    if not r.ok: time.sleep(4); continue
                    for it in r.json().get("data", []):
                        v = it.get("value")
                        if v is None: continue
                        try:
                            d = pd.to_datetime(it["applicableFor"], format="%d/%m/%Y"); rows.append((it.get("itemName"), d.strftime("%Y-%m"), d.strftime("%Y-%m-%d"), float(v)))
                        except Exception: pass
                    break
                except Exception as e:
                    print("fail", type(e).__name__, flush=True); time.sleep(4)
        df = pd.DataFrame(rows, columns=["item", "month", "day", "value"])
        df.groupby(["item", "month"]).agg(sum=("value", "sum"), n=("day", "nunique")).reset_index().to_csv(path, index=False)
        print("gb done", i, len(rows), flush=True)
if "odre" in which:
    B = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets"
    res = {}
    for ds in ("emissions-ges-operateurs", "evolution-emissions-directes-de-methane", "evolution-journaliere-stocks-gnl-terminaux-methaniers", "evolution-de-lactivite-aux-points-dechange-de-gaz-peg-sur-le-reseau-grtgaz0"):
        try:
            r = requests.get(f"{B}/{ds}/records", params={"limit": 100, "order_by": "-1"} if False else {"limit": 100}, timeout=60)
            res[ds] = r.json().get("results", []) if r.ok else r.status_code
        except Exception as e:
            res[ds] = str(e)
    try:
        r = requests.get(f"{B}/prod-nat-gaz-horaire-def/records", params={"limit": 3}, timeout=60); res["prod-nat-gaz-horaire-def"] = r.json().get("results", []) if r.ok else r.status_code
    except Exception as e: res["prod-nat-gaz-horaire-def"] = str(e)
    json.dump(res, open(os.path.join(OUT, "odre_misc.json"), "w"), ensure_ascii=False, indent=1, default=str)
    print("odre ok", flush=True)
if "it" in which:
    res = {}
    for u in ["https://www.snam.it/en/our-businesses/transportation/business-information.html", "https://www.snam.it/it/i-nostri-business/trasporto.html", "https://sisen.mase.gov.it/dgsaie/", "https://sisen.mase.gov.it/dgsaie/gas-naturale"]:
        try:
            r = requests.get(u, headers=UA, timeout=40)
            links = sorted(set(re.findall(r'(?:href|src)="([^"]+)"', r.text)))
            keep = [l for l in links if re.search(r"gas|bilanc|balanc|dati|data|oper|trasp|transp|api|\.js|xls|pdf|csv", l, re.I)][:120]
            res[u] = {"status": r.status_code, "len": len(r.text), "links": keep}
            if "sisen" in u:
                for l in links:
                    if l.endswith(".js"):
                        try:
                            js = requests.get(l if l.startswith("http") else "https://sisen.mase.gov.it" + (l if l.startswith("/") else "/dgsaie/" + l), headers=UA, timeout=40).text
                            res[u].setdefault("js_api", []).extend(sorted(set(re.findall(r'["\'](/?(?:dgsaie/)?api/[^"\']{3,80})["\']', js)))[:60])
                        except Exception as e: pass
        except Exception as e:
            res[u] = {"err": type(e).__name__}
        print(u, res[u].get("status", res[u].get("err")), flush=True)
    json.dump(res, open(os.path.join(OUT, "it_sources2.json"), "w"), indent=1)
