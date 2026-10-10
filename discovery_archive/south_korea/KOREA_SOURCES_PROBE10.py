"""
Probe 10 (manual workflow south_korea_probe.yml): save raw EPSIS ajax responses and grid headings to
discovery_archive/results/south_korea/ (the workflow commits them) so the pull script can be written against real data.
"""
import json
import os
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8", "X-Requested-With": "XMLHttpRequest"}
S = requests.Session()
S.headers.update(H)
B = "https://epsis.kpx.or.kr"
OUT = "discovery_archive/results/south_korea"
os.makedirs(OUT, exist_ok=True)


def save(name, text):
    with open(os.path.join(OUT, name), "w", encoding="utf-8") as f:
        f.write(text)
    print("saved", name, len(text), flush=True)


def layout(t):
    i = t.find("function gridLayoutMake")
    seg = t[i:i + 12000] if i >= 0 else ""
    return re.findall(r'<DataGridColumn(?:Group)?\s[^>]*?headerText="([^"]*)"[^>]*?(?:dataField="([^"]*)")?|<DataGridColumn\s+dataField="([^"]*)"\s+headerText="([^"]*)"', seg)


def post(path, ref, data):
    r = S.post(B + path, data=data, headers={"Referer": B + ref}, timeout=(10, 120))
    print("POST", path, data, r.status_code, len(r.text), flush=True)
    return r.text


heads = {}
pages = {"smp": "/epsisnew/selectEkmaSmpSmpChart.do?menuId=040201", "mep": "/epsisnew/selectEkgeEpsMepChart.do?menuId=030100",
         "ptd": "/epsisnew/selectEkmaPtdBftChart.do?menuId=040501", "bft": "/epsisnew/selectEkpoBftChart.do?menuId=020100",
         "shd": "/epsisnew/selectEkmaSmpShdChart.do?menuId=040202", "gep": "/epsisnew/selectEkgeGepTotChart.do?menuId=060101",
         "ffu": "/epsisnew/selectEkgeFfuChart.do?menuId=060200", "aep": "/epsisnew/selectEkgeEpsAepChart.do?menuId=030200"}
html = {}
for k, p in pages.items():
    html[k] = S.get(B + p, timeout=(10, 120)).text
    i = html[k].find("function gridLayoutMake")
    seg = html[k][i:i + 14000] if i >= 0 else ""
    heads[k] = re.findall(r'<DataGridColumn(?:Group)?\b[^>]*>', seg)
    if k == "shd":
        heads["shd_kind_inputs"] = re.findall(r'<input[^>]*name="selKind"[^>]*>', html[k])
        heads["shd_kind_labels"] = re.sub(r"\s+", " ", html[k][html[k].find('name="selKind"') - 200:html[k].find('name="selKind"') + 900])
    if k in ("gep", "ffu", "aep"):
        save(f"epsis_{k}_page.html", html[k])
save("epsis_headers.json", json.dumps(heads, ensure_ascii=False, indent=1))

save("ptd_month.js", post("/epsisnew/selectEkmaPtdBft.ajax", pages["ptd"], {"selYear": "N", "selRegion": "1"}))
save("ptd_year.js", post("/epsisnew/selectEkmaPtdBft.ajax", pages["ptd"], {"selYear": "Y", "selRegion": "1"}))
save("bft_month.js", post("/epsisnew/selectEkpoBft.ajax", pages["bft"],
                          {"selYear": "N", "selRegion": "1", "selMemgubun": "", "selTelgramform": "", "selBusiType": ""}))
save("bft_year.js", post("/epsisnew/selectEkpoBft.ajax", pages["bft"],
                         {"selYear": "Y", "selRegion": "1", "selMemgubun": "", "selTelgramform": "", "selBusiType": ""}))
save("smp_month.js", post("/epsisnew/selectEkmaSmpSmp.ajax", pages["smp"], {"beginDate": "201501", "endDate": "202610", "selYear": "N"}))
save("mep_month.js", post("/epsisnew/selectEkgeEpsMep.ajax", pages["mep"],
                          {"beginDate": "201501", "endDate": "202610", "selYear": "N", "selMonth": "Y"}))
save("mep_day_sample.js", post("/epsisnew/selectEkgeEpsMep.ajax", pages["mep"],
                               {"beginDate": "20260101", "endDate": "20261010", "selYear": "N", "selMonth": "N"})[:6000])
for kind in ("1", "2", "3", ""):
    t = post("/epsisnew/selectEkmaSmpShd.ajax", pages["shd"],
             {"beginDate": "20260901", "endDate": "20261010", "selYear": "N", "selMonth": "N", "selKind": kind, "locale": ""})
    save(f"shd_day_kind{kind or 'none'}.js", t[:5000])
