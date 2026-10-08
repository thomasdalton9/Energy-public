"""
NEA hunt 6 (manual workflow nea_hunt6.yml): (1) for every capacity release in the press-list JSON, is the table HTML or an image
(era boundary); (2) tesseract (chi_sim) OCR of three table images to judge accuracy.
"""
import re, signal, time, subprocess, os
from urllib.parse import urljoin
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8"}
S = requests.Session(); S.headers.update(H)
class Hard(Exception): pass
def _alarm(*_): raise Hard("hard limit")
signal.signal(signal.SIGALRM, _alarm)
def get(url, quiet=True):
    signal.alarm(40)
    try:
        r = S.get(url, timeout=(10, 25)); signal.alarm(0)
    except (requests.RequestException, Hard) as e:
        signal.alarm(0); print(f"  ERROR {url}: {type(e).__name__}: {str(e)[:100]}", flush=True); return None
    if r.encoding in (None, "ISO-8859-1"): r.encoding = r.apparent_encoding
    time.sleep(0.3); return r
def clean(s): return re.sub("<[^>]+>", "", s or "")

LIST = "https://www.nea.gov.cn/xwfb/"
items = get(LIST + "ds_4f3484af7ea244e7ab18d094856c82a6.json").json()["datasource"]
cap = sorted([i for i in items if re.search(r"全国电力(工业)?统计数据", clean(i["title"]))], key=lambda x: x["publishTime"])
print("capacity releases:", len(cap))
imgs = []
for i in cap:
    u = urljoin(LIST, i["publishUrl"])
    if u.startswith("http://www.nea.gov.cn/20"): u = u.replace("http://", "https://")
    r = get(u)
    if r is None or r.status_code != 200:
        print(i["publishTime"][:10], "ERR", u, r.status_code if r is not None else ""); continue
    t = len(re.findall(r"<table", r.text)); im = re.findall(r'<img[^>]+src="([^"]+)"', r.text)
    im = [x for x in im if not re.search(r"logo|icon|ewm|qr", x, re.I)]
    print(i["publishTime"][:10], clean(i["title"])[:40], "tables", t, "imgs", len(im), u)
    if im: imgs.append((i["publishTime"][:10], urljoin(u, im[0])))
print("=== OCR ===")
subprocess.run("sudo apt-get install -y -q tesseract-ocr tesseract-ocr-chi-sim > /dev/null 2>&1", shell=True)
os.makedirs("/tmp/ocr", exist_ok=True)
pick = [imgs[0], imgs[len(imgs) // 2], imgs[-1]] if imgs else []
for d, u in pick:
    r = get(u)
    if r is None or r.status_code != 200: print("img ERR", u); continue
    ext = "png" if u.lower().endswith("png") else "jpg"
    p = f"/tmp/ocr/{d}.{ext}"; open(p, "wb").write(r.content)
    print("---", d, u, len(r.content), "bytes")
    out = subprocess.run(["tesseract", p, "-", "-l", "chi_sim", "--psm", "6"], capture_output=True, text=True).stdout
    print(out[:2500])
