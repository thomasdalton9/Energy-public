"""Probe 4: Kosovo PxWeb municipal population estimate (POST), Montenegro Census 2023 page, Bosnia agency population pages."""
import json, os, re, requests
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "results", "degree_days")
UA = {"User-Agent": "Mozilla/5.0"}
log = []
def get(url, **kw):
    try:
        r = requests.get(url, headers=UA, timeout=120, **kw)
        log.append(f"{r.status_code} {len(r.content)} {r.url[:170]}")
        return r
    except Exception as e:
        log.append(f"ERR {url[:160]} {e}")
def links(url, pat):
    r = get(url)
    if r is not None and r.ok:
        t = r.content.decode("utf-8", "ignore")
        for h, txt in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', t, flags=re.S):
            txt = re.sub(r"<[^>]+>", "", txt).strip()
            if re.search(pat, txt + h, re.I):
                log.append(f"  LINK {h[:150]} :: {re.sub(chr(10), ' ', txt)[:100]}")
        return t
base = "https://askdata.rks-gov.net/api/v1/en/ASKdata/Population/Estimate%2C%20projection%20and%20structure%20of%20population/Population%20estimate/tab001.px"
r = get(base)
if r is not None and r.ok:
    meta = r.json()
    open(os.path.join(OUT, "kosovo_tab001_meta.json"), "w").write(json.dumps(meta, ensure_ascii=False))
    q = {"query": [], "response": {"format": "json"}}
    try:
        rr = requests.post(base, json=q, headers=UA, timeout=120)
        log.append(f"POST {rr.status_code} {len(rr.content)}")
        open(os.path.join(OUT, "kosovo_tab001_data.json"), "wb").write(rr.content)
    except Exception as e:
        log.append(f"POST ERR {e}")
for u in ("https://monstat.org/eng/page.php?id=1758&page_id=1758", "https://monstat.org/eng/page.php?id=1758&pageid=1758", "https://monstat.org/eng/page.php?id=48&page_id=48"):
    t = links(u, r"\.xls|\.pdf|\.zip|municip|settlement|census|2023")
    if t:
        open(os.path.join(OUT, "probe4_monstat_%s.html" % re.sub(r"\W", "", u[-12:])), "w").write(t[:200000])
for u in ("https://bhas.gov.ba/Calendar/Category/14", "http://popis.gov.ba/", "https://bhas.gov.ba/Calendar/Category?id=14&page=1&statGroup=14&tip=1&lang=en", "https://bhas.gov.ba/data/Publikacije/Bilteni/2016/DEM_00_2016_TB_1_BS.pdf"):
    t = links(u, r"\.xls|\.pdf|municip|census|popis|općin|opcin")
    if t and u.endswith("/14"):
        open(os.path.join(OUT, "probe4_bhas.html"), "w").write(t[:200000])
open(os.path.join(OUT, "probe4_log.txt"), "w").write("\n".join(log))
print("\n".join(log))
