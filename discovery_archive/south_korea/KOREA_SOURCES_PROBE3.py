"""
Probe 3 (manual workflow south_korea_probe.yml): print the inline JavaScript around ajax calls on the EPSIS generation
(selectEkgeGepTotChart), fuel capacity (selectEkpoBftChart), SMP and supply/demand pages, and the full hidden-form
fields, to learn the data endpoints and parameters.
"""
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8"}
S = requests.Session()
S.headers.update(H)
B = "https://epsis.kpx.or.kr"
PAGES = ["/epsisnew/selectEkgeGepTotChart.do?menuId=060101", "/epsisnew/selectEkpoBftChart.do?menuId=020100",
         "/epsisnew/selectEkmaSmpSmpChart.do?menuId=040201", "/epsisnew/selectEkgeEpsMepChart.do?menuId=030100"]
for p in PAGES:
    r = S.get(B + p, timeout=(10, 40))
    t = r.text
    print("=" * 100, "\n", p, r.status_code, len(t), flush=True)
    scripts = re.findall(r"(?s)<script[^>]*>(.*?)</script>", t)
    inline = "\n".join(s for s in scripts if len(s) > 50)
    print("inline script chars:", len(inline))
    for m in re.finditer(r"(?i)(ajax|\.post\(|\.getJSON|XMLHttpRequest|\.submit\(|action\s*=|url\s*:)", inline):
        a = max(0, m.start() - 250)
        print("--- ctx @", m.start(), ":", re.sub(r"\s+", " ", inline[a:m.start() + 700]))
    print("hidden inputs:", re.findall(r'<input[^>]+type="hidden"[^>]*>', t)[:20])
    print("select options (selDate etc):")
    for sid in ("vChartKind", "selDate", "selCount"):
        m = re.search(r'(?s)<select[^>]+id="%s".*?</select>' % sid, t)
        if m:
            print("  ", sid, re.sub(r"\s+", " ", m.group(0))[:600])
