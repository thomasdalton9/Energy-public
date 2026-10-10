"""Japan sources probe 11 (manual): Chubu getFilesInfo.php entries that could hold monthly/yearly area supply-demand results."""
import collections
import io
import json
import zipfile
import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
r = requests.get("https://powergrid.chuden.co.jp/denkiyoho/resource/php/getFilesInfo.php", headers=UA, timeout=40)
d = r.json()
print("categories:", dict(collections.Counter(e.get("category") for e in d)))
for e in d:
    c = str(e.get("category"))
    if "エリア" in c or "年別" in c or "eria" in str(e.get("filename")):
        print("  ", json.dumps(e, ensure_ascii=False)[:230])
