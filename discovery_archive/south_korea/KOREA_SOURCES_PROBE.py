"""
Probe 1 (manual workflow south_korea_probe.yml): which South Korean energy sources answer from GitHub Actions, and what
they return (status, size, content type, text start, links with energy keywords). Parallel, hard time limits.
"""
import re
import concurrent.futures as cf

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "en-US,en;q=0.9,ko;q=0.8"}
URLS = [
    "https://epsis.kpx.or.kr/",
    "https://epsis.kpx.or.kr/epsisnew/",
    "https://epsis.kpx.or.kr/epsisnew/selectEkmaMain.do?locale=en",
    "https://epsis.kpx.or.kr/epsisnew/selectEkgeEpsMepRealChart.do?menuId=040202",
    "https://epsis.kpx.or.kr/epsisnew/selectEkmaUpsBftChart.do?menuId=040100",
    "https://epsis.kpx.or.kr/epsisnew/selectEkmaGepSmpChart.do?menuId=040200",
    "https://www.kpx.or.kr/",
    "https://new.kpx.or.kr/",
    "https://www.data.go.kr/data/15039544/fileData.do",
    "https://www.data.go.kr/en/data/15039544/fileData.do",
    "https://www.data.go.kr/tcs/dss/selectFileDataDownload.do?publicDataPk=15039544",
    "https://www.kogas.or.kr/",
    "https://www.kogas.or.kr/portal/main/main.do",
    "https://www.kogas.or.kr/portal/contents.do?menuCode=statisticsStatistics",
    "https://www.kogas.or.kr/eng/main.do",
    "https://www.kesis.net/",
    "https://www.kesis.net/main/main.jsp",
    "https://www.keei.re.kr/",
    "https://www.kosis.kr/",
    "https://kosis.kr/openapi/index/index.jsp",
    "https://unipass.customs.go.kr/ets/index_eng.do",
    "https://tradedata.go.kr/",
    "https://stat.kita.net/",
    "https://www.customs.go.kr/english/main.do",
    "https://www.petronet.co.kr/",
    "https://www.petronet.co.kr/main2.jsp",
    "https://www.opinet.co.kr/",
    "https://www.motie.go.kr/",
    "https://www.motie.go.kr/kor/article/ATCL3f49a5a8c",
    "https://www.energy.or.kr/",
    "https://www.kepco.co.kr/",
    "https://home.kepco.co.kr/kepco/KO/ntcob/list.do?boardCd=BRD_000211&menuCd=FN06030101",
    "https://bigdata.kepco.co.kr/",
    "https://www.federalreserve.gov/releases/h10/hist/dat00_ko.htm",
    "https://ecos.bok.or.kr/",
    "https://www.index.go.kr/",
    "https://www.ember-energy.org/data/monthly-electricity-data/",
    "https://ourworldindata.org/grapher/electricity-prod-source-stacked.csv?country=KOR",
    "https://www.energyinst.org/statistical-review",
    "https://www.gie.eu/",
    "https://www.kita.net/",
    "https://www.kcs.go.kr/",
    "https://stat.mofa.go.kr/",
    "https://www.kosis.kr/statisticsList/statisticsListIndex.do?menuId=M_01_01&vwcd=MT_ZTITLE&parmTabId=M_01_01",
    "https://data.kma.go.kr/",
    "https://www.knoc.co.kr/",
    "https://www.gasa.or.kr/",
    "https://www.kgu.or.kr/",
    "https://www.ngvtc.or.kr/",
    "https://www.kemco.or.kr/",
    "https://www.iea.org/countries/korea",
    "https://www.kpx.or.kr/menu.es?mid=a10606030000",
    "https://www.kpx.or.kr/eng/",
]
KW = re.compile(r"(csv|xls|xlsx|download|smp|statist|generation|electric|gas|lng|stock|import)", re.I)


def one(u):
    out = [f"=== {u}"]
    try:
        r = requests.get(u, headers=H, timeout=(10, 25))
        ct = r.headers.get("content-type", "")
        out.append(f"  status {r.status_code} bytes {len(r.content)} type {ct}")
        if r.encoding in (None, "ISO-8859-1"):
            r.encoding = r.apparent_encoding
        t = r.text if "text" in ct or "json" in ct or "html" in ct else ""
        title = re.search(r"<title[^>]*>(.*?)</title>", t, re.S | re.I)
        if title:
            out.append("  title: " + re.sub(r"\s+", " ", title.group(1))[:120])
        body = re.sub(r"(?s)<(script|style).*?</\1>", " ", t)
        body = re.sub(r"<[^>]+>", " ", body)
        out.append("  text: " + re.sub(r"\s+", " ", body)[:300])
        ls = re.findall(r'href="([^"#]+)"', t)
        ls = [x for x in dict.fromkeys(ls) if KW.search(x)][:25]
        out.append("  links: " + " | ".join(ls))
    except Exception as e:  # noqa: BLE001
        out.append(f"  ERROR {type(e).__name__}: {str(e)[:150]}")
    return "\n".join(out)


with cf.ThreadPoolExecutor(12) as ex:
    for res in ex.map(one, URLS):
        print(res, flush=True)
