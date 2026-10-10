"""
Probe 14 (manual workflow south_korea_probe.yml): data.go.kr FILE datasets of KOGAS found by probe 13 - Korea's natural gas
imports by continent (15088508), monthly natural gas production (15049906), Asian LNG import unit price (15117762): page,
modification date, download link, and the first lines of the file. Saved to discovery_archive/results/south_korea/.
"""
import os
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "ko-KR,ko;q=0.9"}
OUT = "discovery_archive/results/south_korea"
os.makedirs(OUT, exist_ok=True)
S = requests.Session()
S.headers.update(H)
log = []


def say(*a):
    s = " ".join(str(x) for x in a)
    print(s, flush=True)
    log.append(s)


for ident in ("15088508", "15049906", "15117762"):
    u = f"https://www.data.go.kr/data/{ident}/fileData.do"
    r = S.get(u, timeout=(10, 60))
    t = r.text
    say("PAGE", ident, r.status_code, len(t))
    for key in ("수정일", "등록일", "제공기관", "갱신주기", "확장자", "파일데이터명", "다음 갱신예정일"):
        m = re.search(key + r"[^<]*</[^>]+>\s*<[^>]+>\s*([^<]{1,60})", t)
        say("  ", key, ":", m.group(1).strip() if m else None)
    say("  contentUrl/atch:", re.findall(r'"contentUrl"\s*:\s*"([^"]+)"', t)[:3], re.findall(r"atchFileId=([\w]+)", t)[:3],
        re.findall(r"fileDetailSn=(\d+)", t)[:3])
    links = re.findall(r'(?:href|onclick|data-url)="([^"]*(?:fileDownload|download|atchFileId)[^"]*)"', t)
    say("  links:", links[:5])
    ids = re.findall(r"atchFileId=([\w]+)", t)
    cands = []
    if ids:
        cands.append(f"https://www.data.go.kr/cmm/cmm/fileDownload.do?atchFileId={ids[0]}&fileDetailSn=1")
    for m in re.findall(r'"contentUrl"\s*:\s*"([^"]+)"', t)[:1]:
        cands.append(m)
    for c in cands:
        try:
            rr = S.get(c, timeout=(10, 90))
            say("  DOWNLOAD", c[:140], rr.status_code, len(rr.content), rr.headers.get("content-type"),
                rr.headers.get("content-disposition"))
            raw = rr.content
            open(os.path.join(OUT, f"datagokr_{ident}.bin"), "wb").write(raw[:400000])
            for enc in ("utf-8", "euc-kr", "cp949"):
                try:
                    txt = raw.decode(enc)
                    say("  HEAD", enc, ":", txt[:900].replace("\n", " | "))
                    break
                except UnicodeDecodeError:
                    continue
        except Exception as e:  # noqa: BLE001
            say("  DOWNLOAD ERR", type(e).__name__, str(e)[:100])
open(os.path.join(OUT, "probe14.txt"), "w", encoding="utf-8").write("\n".join(log))
