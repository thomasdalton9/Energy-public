"""Probe 6 (12 Oct 2026): MaStR export download speed: single stream for 40 s, then 8 parallel range streams for 40 s."""
import re, time, requests, threading
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 (research; energy data)"
html = S.get("https://www.marktstammdatenregister.de/MaStR/Datendownload", timeout=60).text
url = re.search(r'href="(https://download\.marktstammdatenregister\.de/Gesamtdatenexport[^"]+\.zip)"', html).group(1)
print(url, flush=True)
def stream(start, end, secs, out, k):
    n = 0; t0 = time.time()
    try:
        with requests.get(url, headers={"Range": f"bytes={start}-{end}", "User-Agent": "Mozilla/5.0"}, stream=True, timeout=60) as r:
            for c in r.iter_content(1 << 20):
                n += len(c)
                if time.time() - t0 > secs: break
    except Exception as e: print("err", k, repr(e)[:80], flush=True)
    out[k] = n
o = {}; stream(0, 3 * 10**9, 40, o, 0); print("single stream MB/s", o[0] / 40 / 1e6, flush=True)
o = {}; ths = [threading.Thread(target=stream, args=(i * 4 * 10**8, (i + 1) * 4 * 10**8 - 1, 40, o, i)) for i in range(8)]
[t.start() for t in ths]; [t.join() for t in ths]
print("8 parallel MB/s total", sum(o.values()) / 40 / 1e6, o, flush=True)
