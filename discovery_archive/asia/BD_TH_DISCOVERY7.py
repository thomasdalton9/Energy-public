"""
Sabah / Sarawak discovery, round 2: the Energy Commission's Malaysia Energy Information Hub statistics portal
(meih.st.gov.my/statistics). Open the 'Electricity Generation' (and installed capacity / final electricity
consumption) views in one session and print the tables and any region / year / fuel selectors and export links.
"""
import io
import re

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 90)
WANT = re.compile(r"Electricity Generation|Installed Capacity|Final Electricity Consumption|Capacity", re.I)


def out(*a):
    print(*a, flush=True)


def main():
    s = requests.Session()
    s.headers.update(H)
    r = s.get("https://meih.st.gov.my/statistics", timeout=T, verify=False)
    out(f"statistics -> {r.status_code} {len(r.text)}")
    links = []
    for href, text in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', r.text, re.S):
        text = re.sub(r"<[^>]+>|\s+", " ", text).strip()
        if WANT.search(text):
            links.append((text, href.replace("&amp;", "&")))
    out(f"{len(links)} links: {[t for t, _ in links]}")
    for text, href in links[:4]:
        p = s.get(href, timeout=T, verify=False)
        out(f"\n######## {text}: {p.status_code} {len(p.text)} {p.url[:160]}")
        try:
            tables = pd.read_html(io.StringIO(p.text))
        except ValueError:
            tables = []
        out(f"  {len(tables)} tables")
        for i, t in enumerate(tables[:5]):
            with pd.option_context("display.width", 250, "display.max_columns", 30):
                out(f"  table {i} {t.shape}\n{t.head(25).to_string()[:3500]}")
        sel = re.findall(r'<select[^>]*name="([^"]+)"[^>]*>(.*?)</select>', p.text, re.S)
        for name, body in sel:
            opts = re.findall(r'<option[^>]*value="([^"]*)"[^>]*>([^<]*)', body)
            out(f"  select {name}: {opts[:30]}")
        for m in re.findall(r'href="([^"]*(?:export|download|excel|xls|csv)[^"]*)"', p.text, re.I)[:10]:
            out(f"  export link: {m.replace('&amp;', '&')[:200]}")
        forms = re.findall(r'<form[^>]*action="([^"]+)"', p.text)
        out(f"  forms: {[f.replace('&amp;', '&')[:160] for f in forms[:5]]}")
        inputs = re.findall(r'<input[^>]*name="([^"]+)"[^>]*value="([^"]*)"', p.text)
        out(f"  inputs: {inputs[:30]}")


if __name__ == "__main__":
    main()
