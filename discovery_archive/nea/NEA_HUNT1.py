"""
NEA hunt 1 (manual workflow nea_hunt1.yml): where does the script-rendered NEA list page get its data, and what do
sitemap / RSS / robots / Wayback CDX / gov.cn search return? Prints findings only; nothing is saved.
"""
import re, signal, time, json
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8"}
S = requests.Session(); S.headers.update(H)

class Hard(Exception): pass
def _alarm(*_): raise Hard("hard limit")
signal.signal(signal.SIGALRM, _alarm)

def get(url, quiet=False):
    signal.alarm(50)
    try:
        r = S.get(url, timeout=(10, 35)); signal.alarm(0)
    except (requests.RequestException, Hard) as e:
        signal.alarm(0); print(f"  ERROR {url}: {type(e).__name__}: {str(e)[:100]}", flush=True); return None
    if r.encoding in (None, "ISO-8859-1"): r.encoding = r.apparent_encoding
    print(f"GET {url} -> {r.status_code} {len(r.content)} bytes", flush=True)
    time.sleep(0.5)
    return r

print("=== 1. list page raw HTML: scripts and data hints ===")
for u in ["https://www.nea.gov.cn/xwfb/index.htm", "https://www.nea.gov.cn/sjzz/index.htm", "https://www.nea.gov.cn/"]:
    r = get(u)
    if r is None or r.status_code != 200: continue
    h = r.text
    print("  script src:", re.findall(r'<script[^>]+src="([^"]+)"', h)[:30])
    for m in re.findall(r'[\w./:-]*\.json[\w./?=&-]*', h)[:20]: print("  json ref:", m)
    for m in re.findall(r'(?s)<script[^>]*>(.*?)</script>', h):
        if re.search(r'ajax|json|getJSON|fetch|datasource|ds_', m, re.I): print("  inline js:", re.sub(r'\s+', ' ', m)[:700])
    print("  head of body:", re.sub(r'\s+', ' ', re.sub(r'(?s)<script.*?</script>', '', h))[:500])

print("=== 2. robots / sitemap / rss ===")
for u in ["https://www.nea.gov.cn/robots.txt", "https://www.nea.gov.cn/sitemap.xml", "https://www.nea.gov.cn/sitemap_index.xml",
          "https://www.nea.gov.cn/rss.xml", "https://www.nea.gov.cn/xwfb/rss.xml", "https://www.nea.gov.cn/sjzz/rss.xml",
          "https://www.nea.gov.cn/rss/", "https://www.nea.gov.cn/index.xml"]:
    r = get(u)
    if r is not None and r.status_code == 200: print(r.text[:600].replace("\n", " "))

print("=== 3. Wayback CDX ===")
for q in ["url=www.nea.gov.cn/xwfb/index.htm&output=json&limit=400&fl=timestamp,statuscode,length&collapse=timestamp:6",
          "url=www.nea.gov.cn/sjzz/*&output=json&limit=50&fl=original,timestamp&collapse=urlkey",
          "url=nea.gov.cn/xwfb/*&matchType=prefix&output=json&limit=100&fl=original,timestamp&collapse=urlkey"]:
    r = get("https://web.archive.org/cdx/search/cdx?" + q)
    if r is not None and r.status_code == 200: print(r.text[:1500])

print("=== 4. Wayback list-page capture ===")
r = get("https://web.archive.org/web/2025/https://www.nea.gov.cn/xwfb/index.htm")
if r is not None and r.status_code == 200:
    print(re.sub(r'\s+', ' ', re.sub(r'(?s)<script.*?</script>', '', r.text))[:800])
    print(re.findall(r'href="([^"]*c\.html)"', r.text)[:10])

print("=== 5. gov.cn search API ===")
for u in ["https://sousuo.www.gov.cn/search-gov/data?t=zhengce&q=%E5%85%A8%E7%A4%BE%E4%BC%9A%E7%94%A8%E7%94%B5%E9%87%8F&timetype=timeqb&mintime=&maxtime=&sort=score&sortType=1&searchfield=title&pcodeJiguan=&childtype=&subchildtype=&tsbq=&pubtimeyear=&puborg=&pcodeYear=&pcodeNum=&filetype=&p=1&n=5&inpro=&bmfl=&dup=&orpro=&type=gwyzcwjk",
          "https://sousuo.www.gov.cn/sousuo/search.shtml?code=17da70961a7&dataTypeId=107&searchWord=%E5%85%A8%E7%A4%BE%E4%BC%9A%E7%94%A8%E7%94%B5%E9%87%8F"]:
    r = get(u)
    if r is not None and r.status_code == 200: print(r.text[:800])
