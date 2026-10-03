"""
Sabah / Sarawak discovery, round 5: myenergystats.st.gov.my (the Energy Commission's current statistics site; the
legacy MEIH portal stops at 2021). Read its electricity pages: tables, download links (xlsx / csv / pdf) and any
Sabah / Sarawak / region breakdown.
"""
import re

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 90)
SITE = "https://myenergystats.st.gov.my"


def out(*a):
    print(*a, flush=True)


def main():
    s = requests.Session()
    s.headers.update(H)
    home = s.get(SITE + "/", timeout=T, verify=False)
    pages = sorted(set(u for u in re.findall(r'href="(/[a-z0-9\-_/]+)"', home.text) if not u.startswith(("/o/", "/documents", "/c/"))))
    out(f"home links: {pages}")
    sm = s.get(SITE + "/sitemap.xml", timeout=T, verify=False).text
    subs = [u.replace("&amp;", "&") for u in re.findall(r"<loc>([^<]+)</loc>", sm)]
    locs = []
    for u in subs[:40]:
        try:
            locs += re.findall(r"<loc>([^<]+)</loc>", s.get(u, timeout=T, verify=False).text)
        except Exception:  # noqa: BLE001
            pass
    locs = sorted(set(locs))
    out(f"sitemap pages ({len(locs)}): {locs}")
    want = [u for u in locs + [SITE + p for p in pages] if re.search(r"electric|power|generation|capacity|sabah|sarawak|region", u, re.I)]
    for u in sorted(set(want))[:12]:
        p = s.get(u, timeout=T, verify=False)
        text = re.sub(r"<script.*?</script>|<style.*?</style>", " ", p.text, flags=re.S)
        text = re.sub(r"<[^>]+>|\s+", " ", text)
        out(f"\n######## {u} -> {p.status_code} {len(p.text)}")
        for m in re.finditer(r".{0,120}(Sabah|Sarawak|Peninsular|Semenanjung).{0,160}", text):
            out("  txt: " + m.group(0)[:280])
        files = sorted(set(re.findall(r'href="([^"]+\.(?:xlsx?|csv|pdf)[^"]*)"', p.text, re.I) +
                           re.findall(r'href="(/documents/[^"]+)"', p.text)))
        out(f"  files: {files[:30]}")
        for m in re.findall(r'(?:iframe|embed)[^>]+src="([^"]+)"', p.text)[:5]:
            out(f"  embed: {m}")
        tables = re.findall(r"<table.*?</table>", p.text, re.S)
        out(f"  {len(tables)} tables")
        for tb in tables[:2]:
            rows = [" | ".join(re.sub(r"<[^>]+>|\s+", " ", c).strip() for c in re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", tr, re.S))
                    for tr in re.findall(r"<tr.*?</tr>", tb, re.S)]
            out("    " + "\n    ".join(rows[:15]))


if __name__ == "__main__":
    main()
