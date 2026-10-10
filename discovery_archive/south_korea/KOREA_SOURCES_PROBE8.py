"""
Probe 8 (manual workflow south_korea_probe.yml): for each EPSIS page of interest print the ajax endpoint + parameters of
dataSerchAjax(), the grid headings and select options, then call it for a recent period. Compact.
"""
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8", "X-Requested-With": "XMLHttpRequest"}
S = requests.Session()
S.headers.update(H)
B = "https://epsis.kpx.or.kr"
PAGES = {
    "TRADED_BY_FUEL": "/epsisnew/selectEkmaPtdBftChart.do?menuId=040501",
    "SETTLE_AMOUNT": "/epsisnew/selectEkmaStmBftChart.do?menuId=040601",
    "FUEL_USE": "/epsisnew/selectEkgeFfuChart.do?menuId=060200",
    "MAXMIN": "/epsisnew/selectEkgeEpsAepChart.do?menuId=030200",
    "SMP_HOURLY": "/epsisnew/selectEkmaSmpShdChart.do?menuId=040202",
    "SMP_LAND": "/epsisnew/selectEkmaSmpNsmChart.do?menuId=040203",
    "CAP_PLANT": "/epsisnew/selectEkpoBgtChart.do?menuId=020200",
    "GEN_MARKET_CAP": "/epsisnew/selectEkmaGcpBftChart.do?menuId=040301",
}


def short(s, n):
    return re.sub(r"\s+", " ", s)[:n]


for name, p in PAGES.items():
    r = S.get(B + p, timeout=(10, 60))
    t = r.text
    print("=====", name, r.status_code, len(t), flush=True)
    if r.status_code != 200:
        continue
    i = t.find("function dataSerchAjax")
    f = t[i:i + 2200] if i >= 0 else ""
    m = re.search(r"\$\.ajax\(\{(.*?)success", f, re.S)
    print(" ajax:", short(m.group(1), 900) if m else short(f, 600))
    i = t.find("function gridLayoutMake")
    seg = t[i:i + 9000] if i >= 0 else ""
    cols = re.findall(r'<DataGridColumn(?:Group)?\s[^>]*?(?:dataField="([^"]*)")?[^>]*?headerText="([^"]*)"', seg)
    print(" cols:", [(a, b) for a, b in cols][:40])
    for sid in re.findall(r'<select[^>]+id="(sel\w+)"', t):
        mm = re.search(r'(?s)<select[^>]+id="%s".*?</select>' % sid, t)
        opts = re.findall(r'<option[^>]*value="([^"]*)"[^>]*>([^<]*)</option>', mm.group(0))
        if len(opts) < 14 and sid != "selCount":
            print("  select", sid, opts)
        elif sid != "selCount":
            print("  select", sid, opts[:3], "...", opts[-2:], len(opts))
    if m:
        url = re.search(r"url\s*:\s*'([^']+)'", m.group(1)).group(1)
        for data in ({"beginDate": "202601", "endDate": "202608", "selYear": "N"},
                     {"beginDate": "202601", "endDate": "202608", "selYear": "N", "selMonth": "Y"}):
            try:
                rr = S.post(B + url, data=data, headers={"Referer": B + p}, timeout=(10, 60))
                print(" POST", url, data, rr.status_code, len(rr.text), "::", short(rr.text, 1100))
            except Exception as e:  # noqa: BLE001
                print(" POST ERR", url, type(e).__name__)
