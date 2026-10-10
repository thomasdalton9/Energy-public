"""
NBS audit probe 2 (GitHub Actions only): dump every table row (and the body text of prose releases) of
every release of a group, to audit what each release holds vs what the workbooks store.
usage: CHINA_NBS_AUDIT_PROBE2.py GROUP   (A industrial+energy, B cpi/ppi/capacity/profits/fai/retail, C 10-day prices)
Reads discovery_archive/results/china_nbs_audit/release_list.tsv (probe 1) plus the 2021 EXTRA ids of the pull scripts.
Output: discovery_archive/results/china_nbs_audit/dump_<group>.jsonl
"""
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..", "..")
sys.path.insert(0, os.path.join(ROOT, "asia"))
sys.path.insert(0, ROOT)
import china_nbs_common as nbs  # noqa: E402
import china_nbs_output as out  # noqa: E402
import CHINA_NBS_MARKET_PRICES as mp  # noqa: E402
import CHINA_NBS_PPI as ppi  # noqa: E402
import CHINA_NBS_CAPACITY_UTILIZATION as cap  # noqa: E402

RES = os.path.join(HERE, "..", "results", "china_nbs_audit")
group = sys.argv[1]
RX = {
    "A": r"规模以上工业(增加值|生产)|能源生产情况",
    "B": r"居民消费价格|工业生产者出厂价格|工业产能利用率|工业企业利润|固定资产投资|社会消费品零售总额",
    "C": r"流通领域重要生产资料",
}[group]
EXTRA = {"A": out.EXTRA_RELEASES, "B": ppi.EXTRA_RELEASES + cap.EXTRA_RELEASES, "C": mp.EXTRA_RELEASES}[group]
L = [l.rstrip("\n").split("\t") for l in open(os.path.join(RES, "release_list.tsv"), encoding="utf-8")]
items = [(t, h) for t, h in L if re.search(RX, t)]
have = {h for _, h in items}
items += [(t, h) for t, h in EXTRA if h not in have]
print(len(items), "releases", flush=True)


def body(html):
    h = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", html, flags=re.S | re.I)
    h = re.sub(r"</(p|div|h\d|li|tr)>", "\n", h, flags=re.I)
    h = re.sub(r"<[^>]+>", "", h).replace("&nbsp;", " ")
    lines = [re.sub(r"\s+", " ", x).strip() for x in h.split("\n")]
    lines = [x for x in lines if x]
    s = next((i for i, x in enumerate(lines) if x.startswith("字体")), 0)
    e = next((i for i, x in enumerate(lines) if x.startswith("[责任编辑") or x.startswith("相关链接")), len(lines))
    return lines[s + 5:e][:140]


with open(os.path.join(RES, f"dump_{group}.jsonl"), "w", encoding="utf-8") as f:
    for t, h in items:
        try:
            html = nbs.fetch(h) or ""
        except Exception as e:  # noqa: BLE001
            print("FAIL", t, type(e).__name__, flush=True)
            f.write(json.dumps({"title": t, "url": h, "error": type(e).__name__}, ensure_ascii=False) + "\n")
            continue
        rows = nbs.table_rows(html)
        f.write(json.dumps({"title": t, "url": h, "rows": rows, "text": body(html) if len(rows) < 25 or group == "A" else body(html)[:12]},
                           ensure_ascii=False) + "\n")
        f.flush()
        print("ok", t, len(rows), flush=True)
