"""
One-off probe (prints only): does AEMO keep INT651 STTM ex-ante price history anywhere the daily pull could
backfill from? Checks NEMWEB Reports/CURRENT/STTM (rolling CSV + DayNN zips), Reports/ARCHIVE/STTM (older zips,
possibly nested), and the AEMO STTM data page, and prints the gas-date range of INT651 found in each file.
"""
import io
import re
import zipfile

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
NEM = "https://nemweb.com.au"


def section(t):
    print("\n" + "=" * 25, t, "=" * 25, flush=True)


def get(url, timeout=180):
    try:
        r = requests.get(url, headers=H, timeout=(15, timeout))
        print(f"{r.status_code} {len(r.content):>12,} B  {url}", flush=True)
        return r if r.ok else None
    except Exception as e:  # noqa: BLE001
        print(f"FAIL {type(e).__name__}: {str(e)[:150]}  {url}", flush=True)
        return None


def listing(url):
    """NEMWEB directory listing -> [(name, size, date text)]."""
    r = get(url)
    if r is None:
        return []
    rows = re.findall(r'([A-Za-z]+, [A-Za-z]+ \d+, \d{4}\s+[\d:]+ [AP]M)\s+(\d+|&lt;dir&gt;)\s+<a href="([^"]+)"', r.text, re.I)
    out = [(href.rstrip("/").split("/")[-1] + ("/" if href.endswith("/") else ""), size, when) for when, size, href in rows]
    if not out:  # fallback: bare links
        out = [(h.rstrip("/").split("/")[-1], "?", "") for h in re.findall(r'href="([^"]+)"', r.text, re.I)]
    print(f"   {len(out)} entries")
    return out


def show(entries, pat=None, n=12):
    hit = [e for e in entries if pat is None or re.search(pat, e[0], re.I)]
    for e in hit[:n]:
        print("   ", e)
    if len(hit) > 2 * n:
        print("    ...")
    for e in hit[-n:] if len(hit) > n else []:
        print("   ", e)
    print(f"   ({len(hit)} matching)")
    return hit


def int651_range(content):
    d = pd.read_csv(io.BytesIO(content))
    dt = pd.to_datetime(d["gas_date"], format="%d %b %Y", errors="coerce").dropna()
    return f"{len(d)} rows, gas days {dt.min():%Y-%m-%d}..{dt.max():%Y-%m-%d} ({dt.nunique()} days)" if len(dt) else f"{len(d)} rows, no dates; cols {list(d.columns)[:8]}"


def scan_zip(content, label, depth=0):
    """Print INT651 members (recursing into nested zips) and their date ranges."""
    try:
        z = zipfile.ZipFile(io.BytesIO(content))
    except Exception as e:  # noqa: BLE001
        print(f"   {label}: not a zip ({type(e).__name__})")
        return
    names = z.namelist()
    inner = [n for n in names if n.lower().endswith(".zip")]
    hits = [n for n in names if "int651" in n.lower()]
    print(f"   {'  ' * depth}{label}: {len(names)} members, {len(inner)} nested zips, {len(hits)} INT651; e.g. {names[:4]}")
    for n in hits[:6]:
        try:
            print(f"   {'  ' * depth}  {n}: {int651_range(z.read(n))}")
        except Exception as e:  # noqa: BLE001
            print(f"   {'  ' * depth}  {n}: parse failed {type(e).__name__} {str(e)[:100]}")
    if depth == 0 and inner:
        for n in [inner[0], inner[-1]]:
            scan_zip(z.read(n), n, depth + 1)


section("CURRENT/STTM")
cur = listing(f"{NEM}/Reports/CURRENT/STTM/")
show(cur, r"int651")
dayz = show(cur, r"\.zip$", 6)
r = get(f"{NEM}/Reports/CURRENT/STTM/int651_v1_ex_ante_market_price_rpt_1.csv")
if r is not None:
    print("   rolling CSV:", int651_range(r.content))
for name in sorted({e[0] for e in dayz[:1] + dayz[-1:]}):
    r = get(f"{NEM}/Reports/CURRENT/STTM/{name}")
    if r is not None:
        scan_zip(r.content, name)

section("ARCHIVE/STTM")
arc = []
for base in (f"{NEM}/Reports/ARCHIVE/STTM/", f"{NEM}/Reports/Archive/STTM/"):
    arc = listing(base)
    if arc:
        break
show(arc, None, 10)
zips = [e for e in arc if e[0].lower().endswith(".zip")]
subdirs = [e for e in arc if e[0].endswith("/")]
pick = zips[:1] + zips[len(zips) // 2:len(zips) // 2 + 1] + zips[-1:] if zips else []
for name, size, when in pick:
    r = get(f"{base}{name}", timeout=600)
    if r is not None:
        scan_zip(r.content, f"{name} ({when})")
for name, _, _ in subdirs[:3]:
    show(listing(f"{base}{name}"), None, 5)

section("AEMO STTM data pages")
for url in ("https://aemo.com.au/energy-systems/gas/short-term-trading-market-sttm/data-sttm",
            "https://aemo.com.au/energy-systems/gas/short-term-trading-market-sttm/data-sttm/price-and-withdrawals"):
    r = get(url)
    if r is not None:
        found = sorted(set(re.findall(r'href="([^"]+)"', r.text, re.I)))
        for u in [u for u in found if re.search(r"\.(xlsx?|csv|zip)(\?|$)|price", u, re.I)][:30]:
            print("   link:", u)
