"""Probe 3: PSE open-data API: first/last record per entity (history start) and field names of the generation / exchange entities."""
import requests

UA = "Mozilla/5.0"
B = "https://api.raporty.pse.pl/api/"
for ep in ["his-wlk-cal", "kse-load", "his-gen-pal", "pk5l-wp", "gen-jw", "przeplywy-mocy", "gen-paliwo", "his-gen-paliwo", "gen-ze", "poze-redoze", "pdgsz", "kmb-teis", "ceny-rcn", "gen-jwcd", "gen-moc-jw", "zap-kse", "popyt-kse", "his-wlk-cal-ver1"]:
    for q in ({"$first": 1}, {"$first": 1, "$orderby": "business_date desc"}):
        try:
            r = requests.get(B + ep, params=q, headers={"User-Agent": UA}, timeout=40)
            print("E", ep, list(q.values())[-1], r.status_code, r.text[:420].replace("\n", " "), flush=True)
        except Exception as e:  # noqa: BLE001
            print("E ERR", ep, repr(e)[:80], flush=True)
        if r.status_code == 404:
            break
for d in ["2022-01-01", "2023-01-02", "2024-06-14", "2024-06-15", "2025-01-02"]:
    r = requests.get(B + "kse-load", params={"$filter": f"business_date eq '{d}'", "$first": 1}, headers={"User-Agent": UA}, timeout=40)
    print("D kse-load", d, r.status_code, r.text[:120].replace("\n", " "), flush=True)
