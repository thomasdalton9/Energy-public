"""
Probe 5 (manual workflow south_korea_probe.yml): call the EPSIS .ajax data endpoints found by probe 3 (they answer with
JavaScript that fills gridData) and print the start of each response; for the generation-by-fuel page print its own
dataSerchAjax() so the parameters are known.
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
    print("PAGE", p, r.status_code, len(r.text), flush=True)
    return r.text


def post(path, ref, data, n=1800):
    try:
        r = S.post(B + path, data=data, headers={"Referer": B + ref}, timeout=(10, 60))
        print(f"POST {path} {data} -> {r.status_code} {len(r.text)} chars", flush=True)
        print("   ", re.sub(r"\s+", " ", r.text)[:n], flush=True)
        return r.text
    except Exception as e:  # noqa: BLE001
        print("POST ERROR", path, type(e).__name__, str(e)[:100])


def func(t, name, n=2500):
    i = t.find("function " + name)
    return re.sub(r"\s+", " ", t[i:i + n]) if i >= 0 else None


GEN = "/epsisnew/selectEkgeGepTotChart.do?menuId=060101"
t = page(GEN)
print("GEN dataSerchAjax:", func(t, "dataSerchAjax"))
print("GEN selects:", re.findall(r'<select[^>]+id="(\w+)"', t))
print("GEN ajax urls:", sorted(set(re.findall(r"url\s*:\s*['\"]([^'\"]+)['\"]", t))))
for sid in ("selDateGrid", "selRegion", "selFuel"):
    m = re.search(r'(?s)<select[^>]+id="%s".*?</select>' % sid, t)
    if m:
        print(sid, re.sub(r"\s+", " ", m.group(0))[:500])

t2 = page("/epsisnew/selectEkmaSmpSmpChart.do?menuId=040201")
post("/epsisnew/selectEkmaSmpSmp.ajax", "/epsisnew/selectEkmaSmpSmpChart.do?menuId=040201",
     {"beginDate": "202509", "endDate": "202610", "selYear": "N"})
post("/epsisnew/selectEkmaSmpSmp.ajax", "/epsisnew/selectEkmaSmpSmpChart.do?menuId=040201",
     {"beginDate": "2023", "endDate": "2026", "selYear": "Y"}, 800)
t3 = page("/epsisnew/selectEkgeEpsMepChart.do?menuId=030100")
post("/epsisnew/selectEkgeEpsMep.ajax", "/epsisnew/selectEkgeEpsMepChart.do?menuId=030100",
     {"beginDate": "20260901", "endDate": "20261008", "selYear": "N", "selMonth": "N"}, 1500)
post("/epsisnew/selectEkgeEpsMep.ajax", "/epsisnew/selectEkgeEpsMepChart.do?menuId=030100",
     {"beginDate": "202501", "endDate": "202609", "selYear": "N", "selMonth": "Y"}, 1500)
t4 = page("/epsisnew/selectEkpoBftChart.do?menuId=020100")
post("/epsisnew/selectEkpoBft.ajax", "/epsisnew/selectEkpoBftChart.do?menuId=020100",
     {"selYear": "N", "selRegion": "", "selMemgubun": "", "selTelgramform": "", "selBusiType": ""}, 1500)
for u in sorted(set(re.findall(r"url\s*:\s*['\"](/epsisnew/[^'\"]+\.ajax)['\"]", t))):
    if "Gep" in u:
        print("GEN endpoint", u)
