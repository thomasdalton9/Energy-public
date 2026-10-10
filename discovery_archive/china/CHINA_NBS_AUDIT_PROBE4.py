"""
NBS audit probe 4 (GitHub Actions only): table dump of the OLD releases (ids found by probe 3's ID scan, 2015-2020) to
test the pull parsers on pre-2021 layouts.  usage: CHINA_NBS_AUDIT_PROBE4.py GROUP  (A industrial+energy, B others, C 10-day)
Output: discovery_archive/results/china_nbs_audit/olddump_<group>.jsonl
"""
import glob
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "asia"))
import china_nbs_common as nbs  # noqa: E402

RES = os.path.join(HERE, "..", "results", "china_nbs_audit")
group = sys.argv[1]
RX = {"A": r"规模以上工业增加值|能源生产",
      "B": r"居民消费价格|工业生产者出厂价格|产能利用率|工业企业利润|固定资产投资|社会消费品零售总额",
      "C": r"流通领域重要生产资料"}[group]
items = []
for f in sorted(glob.glob(os.path.join(RES, "idscan_*.tsv"))):
    for line in open(f, encoding="utf-8"):
        p = line.rstrip("\n").split("\t")
        if len(p) >= 3 and re.search(RX, p[1]) and not p[1].startswith("ERROR"):
            items.append((int(p[0]), p[1].replace("-国家统计局", ""), p[2]))
items.sort(reverse=True)
print(len(items), "old releases", flush=True)
with open(os.path.join(RES, f"olddump_{group}.jsonl"), "w", encoding="utf-8") as f:
    for i, t, u in items:
        try:
            html = nbs.fetch(u) or ""
        except Exception as e:  # noqa: BLE001
            print("FAIL", i, t, type(e).__name__, flush=True)
            continue
        rows = nbs.table_rows(html)
        text = re.sub(r"<[^>]+>", "\n", html)
        text = [x.strip() for x in text.split("\n") if x.strip()]
        keep = [x for x in text if group == "A" and re.search(r"进口|同比", x)][:25] if group == "A" else []
        f.write(json.dumps({"id": i, "title": t, "url": u, "rows": rows, "text": keep}, ensure_ascii=False) + "\n")
        f.flush()
        print("ok", i, t, len(rows), flush=True)
