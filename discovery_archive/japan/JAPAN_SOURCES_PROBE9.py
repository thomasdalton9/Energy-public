"""Japan sources probe 9 (manual): contents of Chubu's getFilesInfo.php and Kansai's /yamasou/jisseki.json."""
import collections
import json
import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
r = requests.get("https://powergrid.chuden.co.jp/denkiyoho/resource/php/getFilesInfo.php", headers=UA, timeout=40)
print("chubu", r.status_code, len(r.content), r.headers.get("content-type"))
d = r.json()
print(type(d), len(d))
cats = collections.defaultdict(list)
for e in d:
    cats[e.get("category")].append(e)
for k, v in cats.items():
    print("CATEGORY", k, len(v))
    for e in v[:3] + v[-2:]:
        print("   ", json.dumps(e, ensure_ascii=False)[:300])
r = requests.get("https://www.kansai-td.co.jp/yamasou/jisseki.json", headers=UA, timeout=40)
print("kansai", r.status_code, len(r.content), r.headers.get("content-type"))
t = r.content.decode("utf-8", "replace")
print(t[:1800])
print("....", t[-600:])
