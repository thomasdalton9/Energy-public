"""Probe 2 (manual workflow south_korea_hourly_probe.yml): EPSIS page controls for hourly/daily generation (menus 040501,
030100, 030300, 060101) and the file lists of the data.go.kr datasets 15065269 / 15069337 / 15065387 (all attached files,
years covered). Results -> discovery_archive/results/south_korea/hourly_probe2.txt"""
import os, re
import requests
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36", "Accept-Language": "ko-KR,ko;q=0.9"}
OUT = "discovery_archive/results/south_korea"
os.makedirs(OUT, exist_ok=True)
log = []
def P(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); log.append(s)
S = requests.Session(); S.headers.update(H)
for ds in ["15065269", "15069337", "15065387", "15065386", "15065388"]:
    try:
        t = S.get(f"https://www.data.go.kr/data/{ds}/fileData.do", timeout=(15, 60)).text
    except Exception as e:
        P(ds, "ERR", type(e).__name__); continue
    P("=====", ds, len(t))
    for m in re.finditer(r"atchFileId=(\w+)[^\"'<>]*?fileDetailSn=(\d+)", t):
        P("  link", m.group(0)[:120])
    for m in re.finditer(r"([^<>\"']{3,80}\.(?:csv|zip|xlsx|xls))", t):
        P("  fname", m.group(1).strip())
    for k in ["fileDetailSn", "atchFileId", "파일데이터명"]:
        for m in list(re.finditer(k, t))[:6]:
            P("  ctx", k, re.sub(r"\s+", " ", t[max(0, m.start() - 80): m.start() + 160]))
E = "https://epsis.kpx.or.kr"
for path in ["/epsisnew/selectEkmaPtdBftChart.do?menuId=040501", "/epsisnew/selectEkgeEpsMepChart.do?menuId=030100",
             "/epsisnew/selectEkgeEpsMepRealChart.do?menuId=030300", "/epsisnew/selectEkgeGepTotChart.do?menuId=060101"]:
    try:
        r = S.get(E + path, timeout=(15, 90)); t = r.text
    except Exception as e:
        P("EPSIS", path, "ERR", type(e).__name__); continue
    P("=====", path, r.status_code, len(t))
    for m in re.finditer(r"<select[^>]*id=\"(\w+)\"[^>]*>(.*?)</select>", t, re.S):
        opts = re.findall(r"<option[^>]*value=\"([^\"]*)\"[^>]*>([^<]*)", m.group(2))
        P("  select", m.group(1), opts[:12], "..." if len(opts) > 12 else "")
    for m in re.finditer(r"<input[^>]*type=\"(radio|hidden)\"[^>]*>", t):
        P("  input", re.sub(r"\s+", " ", m.group(0))[:200])
    for m in sorted(set(re.findall(r"/epsisnew/(\w+)\.ajax", t))):
        P("  ajax", m)
    for m in sorted(set(re.findall(r"url\s*:\s*['\"]([^'\"]+\.ajax)", t))):
        P("  ajaxurl", m)
# real-time page: try common names
for name, key in [("selectEkgeEpsMepRealList", "/epsisnew/selectEkgeEpsMepRealChart.do?menuId=030300")]:
    try:
        r = S.post(E + f"/epsisnew/{name}.ajax", data={}, headers={"Referer": E + key, "X-Requested-With": "XMLHttpRequest"}, timeout=(15, 60))
        P("POST", name, r.status_code, len(r.text), r.text[:300].replace("\n", " "))
    except Exception as e:
        P("POST", name, "ERR", type(e).__name__)
open(os.path.join(OUT, "hourly_probe2.txt"), "w", encoding="utf-8").write("\n".join(log))
