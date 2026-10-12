"""Probe 14: Italy - Terna monthly report files (AdactoBackend download ids) for installed capacity content.
Probe 13 finding: Terna recordsv2 POST answers HTTP 500 (needs the page's session/anti-forgery context); the capacity datasets of the download centre are annual statistics."""
import re, subprocess, sys
from html import unescape
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0"

def run(cmd, timeout=200):
    return subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)

def page(u):
    o = run(["curl", "-sL", "-m", "90", "-A", UA, u])
    return o.stdout

subprocess.run([sys.executable, "-m", "pip", "install", "-q", "pymupdf"])
import pymupdf
h = page("https://www.terna.it/en/electric-system/statistics/monthly-report")
print("page", len(h))
for m in re.finditer(r'<a [^>]*href="([^"]*image=(\d+)[^"]*)"[^>]*>(.*?)</a>', h, re.S):
    tx = " ".join(unescape(re.sub(r"<[^>]+>", " ", m.group(3))).split())
    print("ANCHOR", m.group(2), m.group(1)[:120], "|", tx[:100])
ids = sorted(set(re.findall(r"image=(\d+)", h)))
for i in ids:
    u = f"https://www.terna.it/DesktopModules/AdactoBackend/API/download/get?image={i}"
    o = run(["curl", "-sL", "-m", "120", "-A", UA, "-D", f"/tmp/h{i}.txt", "-o", f"/tmp/f{i}.bin", u])
    hd = open(f"/tmp/h{i}.txt", errors="ignore").read()
    ct = re.findall(r"(?i)content-type:\s*([^\r\n]+)", hd)[-1:] 
    cd = re.findall(r"(?i)content-disposition:\s*([^\r\n]+)", hd)[-1:]
    import os
    sz = os.path.getsize(f"/tmp/f{i}.bin")
    print("FILE", i, ct, cd, sz)
    if sz > 20000 and open(f"/tmp/f{i}.bin", "rb").read(4) == b"%PDF":
        d = pymupdf.open(f"/tmp/f{i}.bin")
        print("  pdf pages", len(d), "| title/first:", d[0].get_text()[:150].replace("\n", " "))
        for pg in d:
            t = pg.get_text()
            if re.search(r"installed|capacity|potenza", t, re.I):
                k = re.search(r"installed|capacity|potenza", t, re.I).start()
                print("  p", pg.number + 1, t[max(0, k - 200):k + 500].replace("\n", " "))
                break
