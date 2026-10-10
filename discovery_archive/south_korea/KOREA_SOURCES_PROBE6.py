"""
Probe 6 (manual workflow south_korea_probe.yml): EPSIS column headings of the grids (gridLayoutMake headerText /
dataField), the select options of the capacity-by-fuel page, and how the generation-by-fuel page (060101) loads its data.
Compact output.
"""
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8", "X-Requested-With": "XMLHttpRequest"}
S = requests.Session()
S.headers.update(H)
B = "https://epsis.kpx.or.kr"


def page(p):
    r = S.get(B + p, timeout=(10, 40))
    return r.text


def heads(t, n=40):
    out = []
    for m in re.finditer(r'(?:headerText|dataField)\s*[:=]\s*"([^"]*)"', t):
        out.append(m.group(1))
    return out[:n]


for name, p in [("SMP", "/epsisnew/selectEkmaSmpSmpChart.do?menuId=040201"),
                ("EPSMEP", "/epsisnew/selectEkgeEpsMepChart.do?menuId=030100"),
                ("CAP", "/epsisnew/selectEkpoBftChart.do?menuId=020100"),
                ("GEN", "/epsisnew/selectEkgeGepTotChart.do?menuId=060101")]:
    t = page(p)
    print("=====", name, len(t))
    i = t.find("function gridLayoutMake")
    seg = t[i:i + 6000] if i >= 0 else ""
    print(" layout:", re.sub(r"\s+", " ", " ".join(f'{a}' for a in re.findall(r'(?:headerText|dataField|<DataGridColumn[^>]*|<DataGridColumnGroup[^>]*)', seg)))[:1800])
    print(" layoutraw:", re.sub(r"\s+", " ", seg)[:1800] if name != "GEN" else "")
    for sid in re.findall(r'<select[^>]+id="(sel\w+)"', t):
        m = re.search(r'(?s)<select[^>]+id="%s".*?</select>' % sid, t)
        opts = re.findall(r'<option[^>]*value="([^"]*)"[^>]*>([^<]*)</option>', m.group(0))
        if name in ("CAP",) or len(opts) < 15:
            print("  select", sid, opts[:12], "..." if len(opts) > 12 else "", len(opts))
    if name == "GEN":
        for kw in ("chartRemake", "selBeginDate", ".submit(", "chartData", "var chartData", "layoutStr"):
            j = t.find(kw)
            print(f" GEN {kw!r} at {j}:", re.sub(r"\s+", " ", t[max(0, j - 100):j + 700]) if j >= 0 else None)
