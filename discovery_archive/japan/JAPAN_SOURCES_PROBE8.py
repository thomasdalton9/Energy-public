"""Japan sources probe 8 (manual): Kansai file list JSON and Chubu getFilesInfo.php request."""
import json
import re
import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
r = requests.get("https://www.kansai-td.co.jp/yamasou/jisseki.json", headers=UA, timeout=40)
print("kansai json", r.status_code, len(r.content), r.headers.get("content-type"))
print(r.content.decode("utf-8", "replace")[:1500])
for u in ("https://www.kansai-td.co.jp/yamasou/eria_jukyu_202608_06.csv",
          "https://www.kansai-td.co.jp/yamasou/eria_jukyu_202609_06.csv"):
    x = requests.get(u, headers=UA, timeout=40)
    print(u, x.status_code, len(x.content))
t = requests.get("https://powergrid.chuden.co.jp/denkiyoho/resource/js/create-download-link.js", headers=UA, timeout=40).text
i = t.find("getFilesInfo.php")
print("CHUBU JS around getFilesInfo:\n", t[max(0, i - 1500): i + 900])
