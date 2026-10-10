"""Japan sources probe 12 (manual): Kansai area-performance page data sources (JSON paths, file lists)."""
import re
import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
B = "https://www.kansai-td.co.jp"
t = requests.get(B + "/denkiyoho/area-performance/index.html", headers=UA, timeout=40).text
print("json paths:", sorted(set(re.findall(r'data-json-path="([^"]+)"', t))))
print("jsonpath attrs:", sorted(set(re.findall(r'data-jsonpath="([^"]+)"', t))))
print("scripts:", [s for s in re.findall(r'<script[^>]+src="([^"]+)"', t) if "denkiyoho" in s or "area" in s])
for m in list(re.finditer(r"filelist|file-list|csv|CSV", t))[:8]:
    print("  ~", t[max(0, m.start() - 120): m.end() + 160].replace("\n", " ").replace("\t", ""))
for u in ("/interchange/denkiyoho/area-performance/jisseki.json", "/interchange/denkiyoho/area-performance/files.json",
          "/interchange/denkiyoho/area-performance/list.json", "/interchange/denkiyoho/area-performance/filelist.json",
          "/interchange/denkiyoho/area-performance/eria_jukyu_202609_06.csv",
          "/interchange/denkiyoho/area-performance/eria_jukyu_202608_06.csv",
          "/interchange/denkiyoho/area-performance/csv/eria_jukyu_202608_06.csv"):
    r = requests.get(B + u, headers=UA, timeout=40)
    print(u, r.status_code, len(r.content), r.headers.get("content-type"))
    if r.status_code == 200 and "html" not in (r.headers.get("content-type") or ""):
        print("    ", r.content[:500].decode("utf-8", "replace").replace("\n", " "))
