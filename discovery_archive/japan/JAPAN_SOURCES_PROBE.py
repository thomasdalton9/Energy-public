"""Japan sources probe (manual): reachability from GitHub Actions + link/CSV structure of the candidate official sources."""
import re
import sys
import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
PAGES = [
    "https://www.tepco.co.jp/forecast/html/area_jukyu-j.html",
    "https://www.tepco.co.jp/forecast/html/area_jukyu_p-j.html",
    "https://www.occto.or.jp/",
    "https://www.occtonet.occto.or.jp/public/dfw/RP11/OCCTO/SD/LOGIN_login",
    "https://www.jepx.jp/electricpower/market-data/spot/",
    "https://www.jepx.org/electricpower/market-data/spot/",
    "https://www.jepx.jp/js/csv_read.php?dir=spot_summary&file=spot_summary_2025.csv",
    "https://www.jepx.jp/electricpower/market-data/spot/csv/spot_summary_2025.csv",
    "https://www.enecho.meti.go.jp/category/electricity_and_gas/electricity_measures/pdf/denryoku_LNG_stock.pdf",
    "https://www.enecho.meti.go.jp/category/electricity_and_gas/electricity_measures/",
    "https://www.enecho.meti.go.jp/statistics/electric_power/ep002/",
    "https://www.enecho.meti.go.jp/statistics/electric_power/ep002/results.html",
    "https://www.enecho.meti.go.jp/statistics/",
    "https://www.customs.go.jp/toukei/info/index_e.htm",
    "https://www.customs.go.jp/toukei/latest/index.htm",
    "https://www.customs.go.jp/toukei/shinbun/trade-st_e/gaiyo.htm",
    "https://www.customs.go.jp/toukei/srch/indexe.htm",
    "https://www.e-stat.go.jp/en/stat-search/files?page=1&toukei=00350300",
    "https://www.federalreserve.gov/releases/h10/hist/dat00_ja.htm",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/gaiyo/resource/",
    "https://www.meti.go.jp/english/statistics/tyo/denryoku/index.html",
    "https://www.enecho.meti.go.jp/en/statistics/",
    "https://www.hepco.co.jp/network/con_service/public_document/supply_demand_results/index.html",
    "https://www.tohoku-epco.co.jp/nw/tohoku_en/juyo/index.html",
    "https://www.kansai-td.co.jp/denkiyoho/area-performance/index.html",
    "https://www.kyuden.co.jp/td_service_wheeling_rule-document_disclosure.html",
    "https://www.yonden.co.jp/nw/supply_demand/data_download.html",
    "https://www.chuden.co.jp/energy/nw_denki/nw_jukyudata/",
    "https://www.energia.co.jp/nw/jukyuu/eria_jukyu.html",
    "https://www.rikuden.co.jp/nw/denki-yoho/jukyu_jisseki.html",
    "https://www.okinawa-epco.co.jp/denki/yoho/",
    "https://www.iea.org/countries/japan",
]
for u in PAGES:
    print("=" * 100)
    print(u)
    try:
        r = requests.get(u, headers=UA, timeout=(10, 60))
    except Exception as e:  # noqa: BLE001
        print("  FAIL", type(e).__name__, str(e)[:150])
        continue
    ct = r.headers.get("content-type", "")
    print(f"  HTTP {r.status_code} {ct} {len(r.content)} bytes final={r.url}")
    if r.status_code != 200:
        continue
    if "html" in ct or r.content[:15].lower().startswith(b"<!doctype") or b"<html" in r.content[:300].lower():
        try:
            txt = r.content.decode(r.apparent_encoding or "utf-8", "replace")
        except Exception:  # noqa: BLE001
            txt = r.text
        links = re.findall(r'href="([^"#]+)"', txt)
        sel = [l for l in dict.fromkeys(links) if re.search(r"\.(csv|xlsx?|pdf|zip)|csv_read|download", l, re.I)]
        print(f"  {len(links)} links; data-like ({len(sel)}):")
        for l in sel[:50]:
            print("   ", l)
    elif "pdf" in ct or u.endswith(".pdf"):
        try:
            import fitz
            d = fitz.open(stream=r.content, filetype="pdf")
            print("  PDF pages", len(d))
            print(d[0].get_text()[:1500])
        except Exception as e:  # noqa: BLE001
            print("  pdf read fail", e)
    else:
        for enc in ("cp932", "utf-8"):
            try:
                t = r.content.decode(enc)
                print(f"  [{enc}] head:")
                for line in t.splitlines()[:6]:
                    print("   ", line[:300])
                break
            except Exception:  # noqa: BLE001
                continue
sys.stdout.flush()
