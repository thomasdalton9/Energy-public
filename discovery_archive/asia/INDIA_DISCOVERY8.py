"""
India daily renewable generation discovery, round 8 (series labels from the bundle, yearly sums): decrypt ICED's daily generation response.

Round 4: GET https://icedapi.niti.gov.in/energy/electricity/generation/daily?source=all -> 200, 375 kB, body is a
CryptoJS AES string ('U2FsdGVkX1...' = OpenSSL salted, passphrase = the KEY in the site's public JS bundle); the
dashboard decrypts it client side. Later calls timed out (slow API / throttling) - this round waits between calls.
  1 daily generation (source=all): structure, sources, date range, sample values
  2 v1 dailyPeakDemand/energyMet and /demand (Grid-India energy met / peak demand met)
"""
import base64
import hashlib
import json
import re
import time

import requests
import urllib3
from cryptography.hazmat.primitives import padding
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36", "Accept": "application/json, text/plain, */*",
     "Origin": "https://iced.niti.gov.in", "Referer": "https://iced.niti.gov.in/"}
T = (20, 300)


def out(*a):
    print(*a, flush=True)


def key_from_bundle():
    r = requests.get("https://iced.niti.gov.in/", headers=H, timeout=T, verify=False)
    js = re.findall(r'src=["\']([^"\']*main[^"\']*\.js)["\']', r.text)[0]
    src = requests.get("https://iced.niti.gov.in/" + js.lstrip("/"), headers=H, timeout=T, verify=False).text
    return re.search(r'KEY:"([^"]+)"', src).group(1)


def decrypt(token, passphrase):
    raw = base64.b64decode(token)
    assert raw[:8] == b"Salted__"
    salt, ct = raw[8:16], raw[16:]
    d, prev = b"", b""
    while len(d) < 48:   # OpenSSL EVP_BytesToKey, MD5, as CryptoJS uses
        prev = hashlib.md5(prev + passphrase.encode() + salt).digest()
        d += prev
    dec = Cipher(algorithms.AES(d[:32]), modes.CBC(d[32:48])).decryptor()
    p = dec.update(ct) + dec.finalize()
    un = padding.PKCS7(128).unpadder()
    return json.loads((un.update(p) + un.finalize()).decode("utf-8"))


def show(obj, depth=0, name="root"):  # noqa: C901
    pad = "  " * (depth + 1)
    if isinstance(obj, dict):
        out(f"{pad}{name}: dict keys {list(obj)[:40]}")
        if depth < 3:
            for k in list(obj)[:15]:
                show(obj[k], depth + 1, k)
    elif isinstance(obj, list):
        out(f"{pad}{name}: list len {len(obj)}; first {json.dumps(obj[:3])[:600]}; last {json.dumps(obj[-2:])[:600]}")
        if obj and isinstance(obj[0], (dict, list)) and depth < 3:
            show(obj[0], depth + 1, name + "[0]")
    else:
        out(f"{pad}{name}: {json.dumps(obj)[:300]}")


def fetch(url, key):
    t0 = time.time()
    try:
        r = requests.get(url, headers=H, timeout=T, verify=False)
    except Exception as e:  # noqa: BLE001
        out(f"\nGET {url}: {type(e).__name__} after {time.time() - t0:.0f}s")
        return None
    out(f"\nGET {url} -> {r.status_code} {len(r.content)} in {time.time() - t0:.0f}s")
    if not r.ok:
        return None
    body = r.json() if r.text.strip().startswith(('"', "{", "[")) else r.text
    if isinstance(body, str):
        body = decrypt(body, key)
    show(body)
    return body


def describe(name, d):
    """d = [dates, series]: show what the second element holds."""
    dates, rest = d[0], d[1:]
    out(f"{name}: {len(dates)} dates {dates[0]}..{dates[-1]}; {len(rest)} more elements")
    for k, x in enumerate(rest):
        if isinstance(x, list):
            out(f"  [{k + 1}] list len {len(x)}; element types {sorted({type(v).__name__ for v in x})}")
            for j, v in enumerate(x[:12]):
                if isinstance(v, dict):
                    out(f"    [{j}] dict keys {list(v)}: " + json.dumps({kk: (vv[-4:] if isinstance(vv, list) else vv)
                                                                    for kk, vv in v.items()})[:900])
                elif isinstance(v, list):
                    out(f"    [{j}] list len {len(v)}: first {v[:4]} last {v[-4:]}")
                else:
                    out(f"    [{j}] {v!r}"[:200])
        elif isinstance(x, dict):
            out(f"  [{k + 1}] dict keys {list(x)}")
            for kk, vv in list(x.items())[:20]:
                out(f"    {kk}: " + (f"list len {len(vv)} first {vv[:3]} last {vv[-4:]}" if isinstance(vv, list) else repr(vv))[:400])
        else:
            out(f"  [{k + 1}] {x!r}"[:300])


def main():
    r = requests.get("https://iced.niti.gov.in/", headers=H, timeout=T, verify=False)
    js = re.findall(r'src=["\']([^"\']*main[^"\']*\.js)["\']', r.text)[0]
    src = requests.get("https://iced.niti.gov.in/" + js.lstrip("/"), headers=H, timeout=T, verify=False).text
    key = re.search(r'KEY:"([^"]+)"', src).group(1)
    # round 7: generation/daily = {data: [dates, [7 unnamed series]]}, MU/day, 01-04-2015..; labels live in the chart code
    for kw in ("generateDailyGenBannerChart(", "dailyGenData", "generateExcelData()", "daily-gen-generation",
               "chartData[1]", "chartData[1][0]", "chartData[1][6]"):
        for m in list(re.finditer(re.escape(kw), src))[:4]:
            out(f"\n  ctx {kw}: {src[max(0, m.start() - 300):m.end() + 1500]!r}")
    d = None
    for i in range(4):
        d = fetch("https://icedapi.niti.gov.in/energy/electricity/generation/daily?source=all", key)
        if d is not None:
            break
        time.sleep(60)
    if d is None:
        return
    import pandas as pd
    dates, series = d["data"]
    df = pd.DataFrame({f"s{k}": pd.to_numeric(pd.Series(v), errors="coerce") for k, v in enumerate(series)})
    df.index = pd.to_datetime(dates, format="%d-%m-%Y")
    out("first non-zero date per series: " + str({c: str(df.index[df[c] > 0].min())[:10] for c in df}))
    out("yearly sums (TWh = MU/1000):\n" + (df.groupby(df.index.year).sum() / 1000).round(1).to_string())
    out("days per year: " + str(df.groupby(df.index.year).size().to_dict()))
    out("last 12 days:\n" + df.tail(12).to_string())
    out("sample 2025-11-18 and 2026-06-13:\n" + df.loc[["2025-11-18", "2026-06-13"]].to_string())
    dup = df.index[df.index.duplicated()]
    out(f"duplicate dates: {len(dup)}; missing days since 2021: "
        f"{len(pd.date_range('2021-01-01', df.index.max()).difference(df.index))}")


if __name__ == "__main__":
    main()
