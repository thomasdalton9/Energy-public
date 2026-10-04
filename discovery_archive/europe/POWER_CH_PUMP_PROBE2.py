"""Show columns/last rows of BFE ogd35 (monthly Swiss electricity balance) and ogd51 (weekly) to find pumping consumption."""
import io, requests, pandas as pd
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36", "Referer": "https://www.bfe.admin.ch/"}
for u in ["https://www.bfe-ogd.ch/ogd35/ogd35_schweizerische_elektrizitaetsbilanz_monatswerte.csv",
          "https://www.bfe-ogd.ch/ogd51/ogd51_wochenstatistik_elektrizitaetsbilanz.csv"]:
    try:
        r = requests.get(u, headers=H, timeout=90); print(u, r.status_code, len(r.text))
        d = pd.read_csv(io.StringIO(r.text)); print(d.columns.tolist()); print(d.tail(6).to_string())
        for c in d.columns:
            if d[c].dtype == object and d[c].nunique() < 40: print(c, d[c].unique().tolist())
    except Exception as e:
        print(u, "ERR", e)
