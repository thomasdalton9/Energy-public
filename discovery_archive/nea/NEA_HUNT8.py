"""
NEA hunt 8 (manual workflow nea_hunt8.yml): save the capacity-table images of the image-era releases (2025-03 ... 2026-08) into
discovery_archive/results/nea/img/ so they can be read by eye, print the HTML table of one 2025 release, and compare tesseract variants.
"""
import re, json, signal, time, subprocess, os
from urllib.parse import urljoin
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8"}
S = requests.Session(); S.headers.update(H)
class Hard(Exception): pass
def _alarm(*_): raise Hard("hard limit")
signal.signal(signal.SIGALRM, _alarm)
def get(url):
    signal.alarm(40)
    try:
        r = S.get(url, timeout=(10, 25)); signal.alarm(0)
    except (requests.RequestException, Hard) as e:
        signal.alarm(0); print(f"  ERROR {url}: {type(e).__name__}: {str(e)[:100]}", flush=True); return None
    if r.encoding in (None, "ISO-8859-1"): r.encoding = r.apparent_encoding
    time.sleep(0.3); return r
def text(html):
    html = re.sub(r"(?s)<(script|style).*?</\1>", " ", html)
    html = re.sub(r"\s+", " ", html)
    html = re.sub(r"</(p|div|tr|li|h\d)>", "\n", html)
    html = re.sub(r"</t[dh]>", " | ", html)
    t = re.sub(r"<[^>]+>", "", html)
    t = re.sub(r"&nbsp;|&emsp;|　", " ", t)
    return re.sub(r"\n\s*\n+", "\n", re.sub(r"[ \t]+", " ", t))

rows = json.load(open("discovery_archive/results/nea/release_index.json"))
os.makedirs("discovery_archive/results/nea/img", exist_ok=True)
subprocess.run("sudo apt-get install -y -q tesseract-ocr tesseract-ocr-chi-sim > /dev/null 2>&1; pip install -q opencv-python-headless numpy", shell=True)
saved = []
for r in rows:
    if r["kind"] != "cap" or r["date"] < "2025-03": continue
    p = get(r["url"])
    if p is None or p.status_code != 200: print("ERR", r["url"]); continue
    ims = [x for x in re.findall(r'<img[^>]+src="([^"]+)"', p.text) if not re.search(r"logo|icon|ewm|qr|1\.gif", x, re.I)]
    t = len(re.findall(r"<table", p.text))
    print(r["date"], "tables", t, "imgs", ims[:2])
    if r["date"] == "2025-07-23": print(text(p.text)[:1800])
    if ims:
        u = urljoin(r["url"], ims[0]); im = get(u)
        if im is not None and im.status_code == 200:
            ext = u.rsplit(".", 1)[-1].lower()
            fn = f"discovery_archive/results/nea/img/{r['date']}.{ext}"
            open(fn, "wb").write(im.content); saved.append(fn)
print("saved", saved)
import cv2, numpy as np
def prep(fn, out):
    g = cv2.imread(fn, cv2.IMREAD_GRAYSCALE)
    g = cv2.resize(g, None, fx=2.5, fy=2.5, interpolation=cv2.INTER_CUBIC)
    b = cv2.threshold(g, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    h = cv2.morphologyEx(b, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (60, 1)))
    v = cv2.morphologyEx(b, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, 60)))
    b = cv2.subtract(b, cv2.add(h, v))
    cv2.imwrite(out, 255 - b)
for fn in saved[:2] + saved[-1:]:
    print("=====", fn)
    cv2img = cv2.imread(fn); print("shape", None if cv2img is None else cv2img.shape)
    pf = "/tmp/prep.png"; prep(fn, pf)
    for name, f, args in [("raw psm6 chi", fn, ["-l", "chi_sim", "--psm", "6"]), ("prep psm6 chi", pf, ["-l", "chi_sim", "--psm", "6"]),
                          ("prep psm4 chi", pf, ["-l", "chi_sim", "--psm", "4"]), ("prep psm6 digits", pf, ["-l", "eng", "--psm", "6", "-c", "tessedit_char_whitelist=0123456789.-*"])]:
        out = subprocess.run(["tesseract", f, "-"] + args, capture_output=True, text=True).stdout
        print("--", name); print(out[:1800])
