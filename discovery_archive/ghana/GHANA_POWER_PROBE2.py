"""Ghana probe 2: Energy Commission WEM statistics pages, GRIDCo report/page lists, VRA, PURC."""
import re, requests
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
def get(u, n=400000):
    r = requests.get(u, headers=H, timeout=(10, 30), stream=True)
    c = b""
    for ch in r.iter_content(65536):
        c += ch
        if len(c) > n: break
    return r, c
def show(u, pat=r"(?i)(\.xls|\.xlsx|\.pdf|\.csv|download|wem|generat|dispatch|hydro|akosombo|lake|reservoir|level|statist|load)", maxl=80, text=0):
    print("=====", u, flush=True)
    try:
        r, c = get(u)
        t = c.decode("utf-8", "ignore")
        print(r.status_code, len(c), r.headers.get("content-type"), r.url, flush=True)
        if text:
            body = re.sub(r"<script.*?</script>|<style.*?</style>", " ", t, flags=re.S)
            body = re.sub(r"<[^>]+>", " ", body); body = re.sub(r"\s+", " ", body)
            print("TEXT:", body[:text], flush=True)
        links = sorted(set(re.findall(r'(?:href|src)=["\']([^"\']+)["\']', t)))
        k = [l for l in links if re.search(pat, l)]
        for l in k[:maxl]: print("   ", l, flush=True)
    except Exception as e:
        print("ERR", type(e).__name__, str(e)[:150], flush=True)
EC = "https://www.energycom.gov.gh"
for p in ["/index.php/planning/weekly-wholesale-electricity-market-wem-statistics", "/index.php/planning/wem-report",
          "/index.php/planning/energy-statistics", "/index.php/planning/key-energy-statistics", "/index.php/planning/ipsmp-data", "/index.php/reports"]:
    show(EC + p, text=1500)
for u in ["https://gridcogh.com/general-reports/", "https://gridcogh.com/annual-reports/", "https://gridcogh.com/page-sitemap.xml", "https://gridcogh.com/post-sitemap.xml"]:
    show(u, pat=r"(?i)(report|dispatch|generat|load|hydro|akosombo|bui|data|dashboard|stat|xls|csv)", text=300)
show("https://vra.com/our_mandate/akosombo_hydro_plant.php", text=1500)
show("https://vra.com/resources/annual_reports.php", text=300)
show("https://purc.com.gh/categ/reports/subcategories/electricity", text=300)
for u in ["https://gridcogh.com/wp-json/wp/v2/search?search=dispatch&per_page=50", "https://gridcogh.com/wp-json/wp/v2/search?search=generation&per_page=50",
          "https://gridcogh.com/wp-json/wp/v2/media?search=dispatch&per_page=50", "https://gridcogh.com/wp-json/wp/v2/media?search=xlsx&per_page=50",
          "https://gridcogh.com/wp-json/wp/v2/media?mime_type=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet&per_page=50",
          "https://gridcogh.com/wp-json/wp/v2/media?search=report&per_page=100&_fields=link,title,source_url,mime_type,date"]:
    print("=====", u, flush=True)
    try:
        r, c = get(u, 300000); print(r.status_code, len(c)); print(c.decode("utf-8","ignore")[:1800], flush=True)
    except Exception as e: print("ERR", e)
