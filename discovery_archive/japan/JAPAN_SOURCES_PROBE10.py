"""Japan sources probe 10 (manual): members of the Kansai monthly *_jisseki.zip and the Chubu monthly *_keito.zip."""
import io
import zipfile
import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def zipinfo(u):
    try:
        r = requests.get(u, headers=UA, timeout=60)
    except Exception as e:  # noqa: BLE001
        print(u, "FAIL", type(e).__name__)
        return
    print(u, r.status_code, len(r.content), r.headers.get("content-type"))
    if r.status_code != 200:
        return
    try:
        z = zipfile.ZipFile(io.BytesIO(r.content))
    except Exception as e:  # noqa: BLE001
        print("   not a zip:", e, r.content[:80])
        return
    names = z.namelist()
    print("   members:", len(names), names[:6])
    for n in z.namelist()[:2]:
        b = z.read(n)
        for enc in ("cp932", "utf-8-sig"):
            try:
                t = b.decode(enc)
                print("   ---", n, enc)
                for line in t.splitlines()[:5]:
                    print("      |", line[:260])
                break
            except Exception:  # noqa: BLE001
                continue


for base in ("https://www.kansai-td.co.jp/yamasou/", "https://www.kansai-td.co.jp/interchange/denkiyoho/area-performance/",
             "https://www.kansai-td.co.jp/denkiyoho/download/"):
    zipinfo(base + "202608_jisseki.zip")
