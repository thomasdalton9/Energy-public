"""
NEA hunt 5 (manual workflow nea_hunt5.yml): full text, <table>/<img> structure of sample NEA release articles of each era,
to see whether the capacity table of newer articles is HTML or an image.
"""
import re, signal, time
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
    print(f"GET {url} -> {r.status_code} {len(r.content)} bytes", flush=True)
    time.sleep(0.4); return r
def text(html):
    html = re.sub(r"(?s)<(script|style).*?</\1>", " ", html)
    html = re.sub(r"\s+", " ", html)
    html = re.sub(r"</(p|div|tr|li|h\d)>", "\n", html)
    html = re.sub(r"</t[dh]>", " | ", html)
    t = re.sub(r"<[^>]+>", "", html)
    t = re.sub(r"&nbsp;|&emsp;|　", " ", t)
    return re.sub(r"\n\s*\n+", "\n", re.sub(r"[ \t]+", " ", t))
U = ["20250120/097bfd7c1cd3498897639857d86d5dac", "20250623/6a557ee5ad944062a6c01d6e8417474e", "20251226/640306962d7d421b921b902f48a04b47",
     "20260129/6874f211acd0417eab7ac10c3061a7c2", "20260722/3b678308556b4df0a574537b59a28731", "20260825/1430d07985474c5b8110552785f1c74b",
     "20260821/f227e3fb9f05422285e354f4c27aec07", "20260920/bd277235282140f4b43e9e0f5f0af3d9", "20250120/4f7f249bac714e7693adecac996d742f"]
for u in U:
    r = get(f"https://www.nea.gov.cn/{u}/c.html")
    if r is None or r.status_code != 200: continue
    h = r.text
    print("tables:", len(re.findall(r"<table", h)), "imgs:", re.findall(r'<img[^>]+src="([^"]+)"', h)[:12])
    print("attachments:", re.findall(r'href="([^"]+\.(?:xlsx?|pdf|docx?|png|jpg))"', h)[:10])
    print(text(h)[:6000]); print("------")
