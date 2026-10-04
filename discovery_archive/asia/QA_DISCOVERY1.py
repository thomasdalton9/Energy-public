"""
QA probe 1 (one-off, read-only): check suspicious Vietnam EVN daily posts against the source text.
  - 31/12/2025 (total 639.6 GWh, peak 35.8 GW - looks like the 1 Jan holiday) and 1/1/2026 (is there a post?)
  - January 2026 (missing from the workbook): which posts does the list carry?
  - 20/9/2025 (hydro parsed as 30.8 GWh vs ~430 neighbours; total check off by 400 GWh)
  - a 2023 / early-2024 post (solar not parsed before 4 Jun 2024): print the source-mix lines
  - 3/6/2024 (stated total present, solar missing)
Prints the 'Co cau san luong' section of each post. Uses asia/VIETNAM_EVN.py's own list / parse helpers.
"""
import os
import re
import sys
from datetime import date

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "asia"))
sys.path.insert(0, ROOT)
import VIETNAM_EVN as V  # noqa: E402

WANT = {date(2025, 12, 31), date(2026, 1, 1), date(2026, 1, 2), date(2026, 1, 15), date(2025, 9, 20),
        date(2025, 9, 19), date(2024, 6, 3), date(2024, 6, 4), date(2023, 8, 15), date(2024, 2, 1)}

posts, raw_links = {}, []
for page in range(1, 140):
    t = V.get(f"{V.LIST}?page={page}")
    if t is None:
        print("page", page, "failed")
        continue
    # re-implement the listing parse on the fetched page so the raw link text can be shown
    for u, title in re.findall(r'<a[^>]+href="(/d/vi-VN/news/[^"]*?van-hanh-he-thong-dien-Quoc-gia[^"]*)"[^>]*>(.*?)</a>',
                               t, re.I | re.S):
        txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", title)).strip()
        m = re.search(r"ngày\s*(\d{1,2})/(\d{1,2})/(\d{4})", txt)
        if m:
            d = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            posts.setdefault(d, (V.BASE + u, txt))
        elif re.search(r"(2026|2025)", u) and ("-12" in u or "2026" in u):
            raw_links.append((page, u, txt))
    if posts and min(posts) < date(2023, 8, 1):
        break
print(f"{len(posts)} dated posts, {min(posts)}..{max(posts)}")
jan = sorted(d for d in posts if d.year == 2026 and d.month == 1)
print("January 2026 posts in the list:", jan)
dec = sorted(d for d in posts if d >= date(2025, 12, 25) and d <= date(2026, 2, 3))
print("Posts 25 Dec 2025 - 3 Feb 2026:", dec)
print("Undated links mentioning 2025/2026 (first 40):")
for r in raw_links[:40]:
    print("  ", r)

for d in sorted(WANT):
    if d not in posts:
        print(f"\n######## {d}: NOT LISTED")
        continue
    url, txt = posts[d]
    print(f"\n######## {d}: {txt}\n{url}")
    page = V.get(url)
    if not page:
        print("  fetch failed")
        continue
    t = V.plain(page)
    i = t.find("Thông tin chung về vận hành")
    seg = t[i:i + 3500] if i >= 0 else t[:3500]
    print(re.sub(r"\n\s*\n+", "\n", seg))
    print("PARSED:", V.parse(page))
