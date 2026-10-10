"""Probe (manual workflow south_korea_hourly_probe.yml): KPX hourly generation file datasets on data.go.kr
(15065387 hourly generation, 15065269 regional hourly solar+wind, 15069337 renewables, 15127502 Jeju solar/wind) and EPSIS
hourly pages. Prints file names, download ids, sizes, header/first/last rows. Results -> discovery_archive/results/south_korea/hourly_probe.txt"""
import io, os, re
import requests
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36", "Accept-Language": "ko-KR,ko;q=0.9"}
OUT = "discovery_archive/results/south_korea"
os.makedirs(OUT, exist_ok=True)
log = []
def P(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); log.append(s)
S = requests.Session(); S.headers.update(H)
for ds in ["15065387", "15065269", "15069337", "15127502", "15069334", "15153657"]:
    P("=====", ds)
    for url in (f"https://www.data.go.kr/data/{ds}/fileData.do", f"https://www.data.go.kr/dcat/metadata/{ds}"):
        try:
            r = S.get(url, timeout=(15, 60)); t = r.text
            P(url, r.status_code, len(t))
        except Exception as e:
            P(url, "ERR", type(e).__name__); continue
        if "dcat" in url:
            P(t[:1500]); continue
        for k in ["파일데이터명", "수정일", "등록일", "제공기관", "확장자", "업데이트 주기", "다음 갱신예정일", "전체 행", "데이터 한계"]:
            m = re.search(re.escape(k) + r"[^<]*</[^>]+>\s*<[^>]+>\s*([^<]{1,100})", t)
            P("  ", k, "=", m.group(1).strip() if m else None)
        ids = sorted(set(re.findall(r"atchFileId=(\w+)&(?:amp;)?fileDetailSn=(\d+)", t)))
        P("  download ids", ids)
        for fid, sn in ids[:3]:
            d = f"https://www.data.go.kr/cmm/cmm/fileDownload.do?atchFileId={fid}&fileDetailSn={sn}"
            try:
                rr = S.get(d, timeout=(15, 180))
                raw = rr.content
                P("  DL", d, rr.status_code, len(raw), rr.headers.get("content-disposition"))
                for enc in ("utf-8-sig", "cp949"):
                    try:
                        txt = raw.decode(enc); break
                    except Exception:
                        txt = None
                if txt:
                    lines = txt.splitlines()
                    P("  enc", enc, "lines", len(lines))
                    for l in lines[:4]: P("   H:", l[:300])
                    for l in lines[-3:]: P("   T:", l[:300])
                else:
                    P("  undecodable; head", raw[:100])
            except Exception as e:
                P("  DL ERR", type(e).__name__, e)
# EPSIS: look for hourly generation menus
E = "https://epsis.kpx.or.kr"
for path in ["/epsisnew/selectEkmaPtdBftChart.do?menuId=040501", "/epsisnew/selectEkmaPtdBftHourChart.do?menuId=040502",
             "/epsisnew/selectEkgeEpsMepChart.do?menuId=030100", "/epsisnew/selectEkgeEpsMepRealChart.do?menuId=030300",
             "/epsisnew/selectEkgeEpsPrgChart.do?menuId=030300", "/epsisnew/main.do"]:
    try:
        r = S.get(E + path, timeout=(15, 60)); t = r.text
        P("EPSIS", path, r.status_code, len(t))
        for m in sorted(set(re.findall(r"menuId=(\d{6})[^>]*>\s*([^<]{2,40})<", t)))[:80]:
            P("   menu", m)
    except Exception as e:
        P("EPSIS", path, "ERR", type(e).__name__)
open(os.path.join(OUT, "hourly_probe.txt"), "w", encoding="utf-8").write("\n".join(log))
