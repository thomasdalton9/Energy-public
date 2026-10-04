"""
Grid-operator data discovery, round 3 (after SCADA_DISCOVERY2.py):
  Vietnam   EVN daily posts 'Thong tin chung ve van hanh he thong dien Quoc gia ngay DDMMYYYY': full text of two
            posts (generation by source?), how far back the list pages go (paging), the reservoir page's data call
  Pakistan  NEPRA FPA 'Plant wise Units.xlsx' (hourly MW per plant with fuel): coverage (dates, plants, fuels), monthly
            totals by fuel; which other months are posted (Admission Notices folders), CPPA fuel adjustment pages
  Sri Lanka CEB cebcare GenSum endpoints (JSON)
"""
import io
import re

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36", "Accept": "text/html,application/json,*/*;q=0.8", "Accept-Language": "en-US,en;q=0.9"}
T = (15, 90)


def out(*a):
    print(*a, flush=True)


def get(u, **kw):
    try:
        r = requests.get(u, headers=H, timeout=T, verify=False, **kw)
        out(f"GET {u} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} ERROR {type(e).__name__}: {str(e)[:120]}")
        return None


def text_of(html):
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", html, flags=re.S | re.I)
    return re.sub(r"<[^>]+>|\s+", " ", t)


def vietnam():
    base = "https://www.evn.com.vn"
    lst = get(base + "/vi-VN/news-l/Thong-tin-tom-tat-van-hanh-HTD-Quoc-gia-60-2015")
    posts = sorted(set(re.findall(r'href="(/d/vi-VN/news/Thong-tin-chung[^"]+)"', lst.text)))
    out(f"  {len(posts)} daily posts on page 1: {posts[:3]}")
    for p in posts[:2]:
        x = get(base + p)
        t = text_of(x.text)
        i = t.find("Thông tin chung")
        out("    TEXT: " + t[i:i + 3500])
        imgs = re.findall(r'src="(/userfile[^"]+)"', x.text)
        out(f"    images: {imgs[:4]}")
        out(f"    tables: {len(re.findall(r'<table', x.text))}")
    # paging: look for the pager links / ajax
    pg = re.findall(r'href="([^"]*(?:page|trang|p=)\d+[^"]*)"', lst.text, re.I)
    out(f"  pager links: {pg[:10]}")
    for m in re.finditer(r"(\$\.(?:ajax|get|post)|fetch\()\s*\(?\s*[{'\"][^;]{0,300}", lst.text):
        out("    js: " + m.group(0)[:300])
    for q in ("?page=2", "/2", "?p=2"):
        y = get(base + "/vi-VN/news-l/Thong-tin-tom-tat-van-hanh-HTD-Quoc-gia-60-2015" + q)
        if y is not None and y.ok:
            ps = sorted(set(re.findall(r'Thong-tin-chung-ve-van-hanh-he-thong-dien-Quoc-gia-ngay-(\d{8})', y.text)))
            out(f"    {q}: dates {ps[:3]}..{ps[-3:]}")
    # guess older post urls by id search? try the site search
    s = get(base + "/vi-VN/search?q=Th%C3%B4ng%20tin%20chung%20v%E1%BB%81%20v%E1%BA%ADn%20h%C3%A0nh")
    if s is not None and s.ok:
        ps = sorted(set(re.findall(r'van-hanh-he-thong-dien-Quoc-gia-ngay-(\d{8})', s.text)))
        out(f"    search dates: {ps[:5]}..{ps[-5:]} ({len(ps)})")
    r = get(base + "/vi-VN/thong-tin-ho-thuy-dien/Muc-nuoc-cac-ho-thuy-dien-60-123")
    for m in re.finditer(r"(\$\.(?:ajax|get|post|getJSON)|fetch\(|url\s*:)\s*[^;]{0,250}", r.text):
        out("    js: " + m.group(0)[:250])
    scr = re.findall(r'<script[^>]+src="([^"]+)"', r.text)
    out(f"    scripts: {scr[:25]}")
    i = r.text.find("Mực nước các hồ thủy điện", r.text.find("Trang chủ"))
    out("    html around content: " + r.text[i:i + 2500])


def pakistan():
    x = get("https://nepra.org.pk/Admission%20Notices/2026/09%20Sep/01-%20Plant%20wise%20Units.xlsx")
    d = pd.read_excel(io.BytesIO(x.content))
    d["EventDate"] = pd.to_datetime(d["EventDate"])
    out(f"  rows {len(d)}; dates {d.EventDate.min()}..{d.EventDate.max()}; plants {d.Plant.nunique()}; "
        f"fuels {d.Fuel.value_counts().to_dict()}")
    g = d.groupby([d.EventDate.dt.to_period("M"), "Fuel"])["MW"].sum().unstack() / 1000
    out("  GWh by month and fuel:\n" + g.round(0).to_string())
    out("  plants per fuel: " + str(d.groupby("Fuel").Plant.nunique().to_dict()))
    out("  daily total GWh (first days): " + str((d.groupby("EventDate").MW.sum() / 1000).round(1).head(5).to_dict()))
    for f in ("02-%20Benefitted%20Consumption.xlsx", "03-%20Hourly%20Marginal%20Price.xlsx"):
        y = get("https://nepra.org.pk/Admission%20Notices/2026/09%20Sep/" + f)
        if y is not None and y.ok:
            z = pd.read_excel(io.BytesIO(y.content), nrows=8)
            out(f"  {f}: columns {list(z.columns)}\n{z.to_string()[:1200]}")
    # other months: list the Admission Notices index pages
    for u in ("https://nepra.org.pk/admission-notices.php", "https://nepra.org.pk/Admission%20Notices/2026/09%20Sep/",
              "https://nepra.org.pk/Admission%20Notices/2026/", "https://nepra.org.pk/hearings.php"):
        y = get(u)
        if y is not None and y.ok:
            fl = sorted(set(re.findall(r'href="([^"]*(?:Plant|Hourly|Units|FCA|FPA)[^"]*)"', y.text, re.I)))
            out(f"    files: {fl[:40]}")
    r = get("https://cppa.gov.pk/fuel-adjustment-notifications")
    if r is not None and r.ok:
        fl = sorted(set(re.findall(r'href="([^"]+\.(?:pdf|xlsx?|zip))"', r.text, re.I)))
        out(f"    CPPA FCA files: {fl[:30]}")


def sri_lanka():
    for k in ("GetEnergySummary", "GetEnergySummaryType3", "GetEnergySummaryType4", "GetNightPeakPowerSummaryType3"):
        r = get("https://cebcare.ceb.lk/GenSum/" + k)
        if r is not None and r.ok:
            out("    " + r.text[:1500])


if __name__ == "__main__":
    import sys
    for f in sys.argv[1:] or ["pakistan", "vietnam", "sri_lanka"]:
        out(f"\n==================== {f}")
        try:
            globals()[f]()
        except Exception as e:  # noqa: BLE001
            out(f"!! {f}: {e!r}")
