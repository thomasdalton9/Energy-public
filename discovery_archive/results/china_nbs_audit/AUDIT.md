# NBS workbooks re-audit (10 Oct 2026)

Scope: the six workbooks in `output/Data and Chart Outputs/` pulled from the National Bureau of Statistics (NBS) Chinese release
list `https://www.stats.gov.cn/sj/zxfb/` (`data.stats.gov.cn` still answers 403 from Actions). Method: probes 1-6 in this folder
(`discovery_archive/china/CHINA_NBS_AUDIT_PROBE1-6.py`) opened the real release pages from GitHub Actions:
the full release list (1003 releases, Sep 2021 - Oct 2026), a table dump of every industrial, energy, CPI, PPI, capacity, profit,
investment, retail and 10-day price release on it (`dump_A/B/C.jsonl`), an id scan of the migrated archive
`/sj/zxfb/202302/t20230203_<id>.html` (`idscan_*.tsv`: **everything NBS migrated, Dec 2012 - Oct 2021, is still online**),
a dump of the old (2015-2020) releases to test the parsers (`olddump_*.jsonl`) and a Wayback CDX search (probe 6).
All counts below come from the committed workbooks after the rebuild (run ids in the report).

## 1. Why months look missing (findings)

| Finding | Cause | Evidence |
|---|---|---|
| Every January and February is blank in the monthly NBS series (industrial, energy, clean-energy, retail; 26 cells, 2014-2026) | **Publication practice.** The January survey is not reported ("1月份数据免报" in the release notes) and NBS publishes one combined January-February release. Not split or filled; the two-month figures are on the `Jan-Feb` sheets (13 years, 2014-2026) | releases `YYYY年1—2月份...`; note 3 of the industrial release |
| Solar cells blank Mar 2021 - Feb 2023 (20 months) | **Publication practice.** The product row first appears in the Jan-Feb 2023 release table | all 22 releases before it hold no `太阳能电池` row |
| Service robots blank before Mar 2025 | **Publication practice.** Row first appears in the Jan-Feb 2025 table | same check |
| Industry value-added rows (coal mining, oil and gas extraction) only from Feb 2022 | **Publication practice** (the main-industry list was shorter) | releases Mar-Dec 2021 list 26 rows, from Jan-Feb 2022 29-30 |
| Steam coal 4500/5000/5800 kcal, gasoline 92#, styrene, PVC, bagged cement end Dec 2025 in the 10-day prices | **NBS dropped them.** The 2026 Jan 1-10 release note states: 50 products optimised, "新增乙醇、冰醋酸、磷酸铁锂、多晶硅、白糖、磷肥、钾肥等7种产品，删除苯乙烯、聚氯乙烯、92号汽油、袋装普通硅酸盐水泥、普通混煤、山西大混和大同混煤等7种产品". Every release (200 of 200, 50 price rows each) parses fully - no parser loss | `sample_xun_2026年1月上旬*.txt`, `dump_C.jsonl` |
| Hot-rolled sheet ends Dec 2021 | NBS replaced it by hot-rolled coil from 2022 (earlier note, unchanged) | release tables |
| 10-day prices: 12 periods missing (Spring Festival / Golden Week) | 2025 Oct 1-10 and 2026 Feb 11-20 were **not published** (the next release's note "上期为..." skips them). The other ten (2015-02-11, 2016-02-01, 2017-01-21, 2018-02-11, 2019-02-01, 2021-02-11, 2022-02-01, 2023-01-21, 2024-02-11, 2025-01-21) were published (the next release compares with them) but are **neither on the release list nor in the migrated archive** (every id around them scanned; Wayback CDX found nothing). Left blank | `dump_C.jsonl` notes, `idscan_1901366/1901733.tsv`, `probe6_*` |
| Profits: Jan row never present; Feb-Apr 2013, May 2015, Jul 2019 | January is never released alone (YTD series start with Jan-Feb). 2013 profit releases are prose only (no table). The May-2015 and Jul-2019 profit releases are not in the archive | release counts below |
| Energy production release (能源生产情况): "much more than we store" | The release is **prose, no table**; its monthly levels are the industrial table's (unrounded there). Extra prose items: coal/crude/gas imports (Mar 2022 - Dec 2024 and some earlier months; not in 2025+ releases), daily averages (= month / days), two-year averages (2021-22 only), Qinhuangdao coal prices and Brent (Oct-Dec 2021 only) | `dump_A.jsonl` text, `sample_energy_production_newest.txt` |
| Crude oil, coal, nuclear, wind, solar, aluminium blank in 2013-2016 | **Publication practice / names.** Early tables list fewer products: coal from Apr 2015, nuclear/wind/solar generation and primary aluminium from Mar-Jun 2016 (alumina only to 2015); crude was printed as `天然原油` before 2016 (a parser name gap, fixed) | `olddump_A.jsonl` |
| Capacity utilisation: "every quarter since 2021" | **Complete**: 35 quarters 2017Q4 - 2026Q2, 0 gaps, 19 industries each (NBS's quarterly series online starts 2017Q4) | `china_nbs_capacity_utilization_quarterly.xlsx` |
| PPI "100 columns" | Every row of every release is mapped (0 unmapped rows in 68 releases); the release has no further table. Added the YTD y/y column NBS prints (50 more columns) | `dump_B.jsonl` |
| Schedules | The NBS pulls ran on the 5th/15th/20th/25th/13th while the master rebuilds on the 1st/15th evening (the 15th build saw the previous cycle). Now all on the 1st and 15th, 05:00-06:30 UTC, before the master | workflow files |

## 2. What each release holds versus what was stored

| Release | Content of the release | Stored before | Stored now |
|---|---|---|---|
| Industrial production (规模以上工业增加值) | product table, 38 rows: month value, month y/y, YTD value, YTD y/y; value-added growth (month, YTD) for 3 sectors, high-tech, 4 ownership groups, ~29 industries; product sales rate; export delivery value; seasonally adjusted m/m (only as a revised 13-month vintage table) | 36 product levels, Mar 2021 - Aug 2026 | 38 product levels (+ alumina to 2015) **with month y/y, YTD, YTD y/y** (4 columns each), 30 value-added series x2, sales rate, export value; Sep 2013 - Aug 2026 (130 months + 13 Jan-Feb). 213 columns |
| Energy production (能源生产情况) | prose: coal, crude, refinery runs, gas, power by source (levels, y/y, YTD), daily averages, imports (2022-24) | 11 columns from the industrial table | same 11 + y/y/YTD + 7 energy value-added series (58 columns) + `Imports` sheets (customs flash, 7 series, Jul 2017 - Dec 2024); history from Sep 2013 |
| 10-day prices (流通领域重要生产资料市场价格) | 50 products: price, change (yuan), change % | 58 columns, price only, Jan 2021 - | 62 product columns (4 products to 2019 added) + NBS's change % for each (124 columns), **Jan 2014 - Sep 2026, 447 periods** |
| PPI (工业生产者出厂价格) | m/m, y/y, YTD y/y for 50 series | 100 columns, 68 months | 150 columns, **Jan 2013 - Aug 2026, 164 months, no gap** |
| CPI (居民消费价格) | m/m, y/y, YTD y/y for 46 items (food, housing utilities `水电燃料`, vehicle fuel `交通工具用燃料/能源`, core excl. food and energy ...) | not stored | **new workbook** china_nbs_cpi_monthly.xlsx, 138 columns, Jan 2013 - Aug 2026 (vehicle fuel and utilities from 2016) |
| Capacity utilisation | quarter rate, change vs year ago, YTD rate and change, 19 industries | 19 columns, 22 quarters | 76 columns, 35 quarters from 2017Q4 |
| Industrial profits (工业企业利润) | YTD revenue, cost, profit and growth by sector, ownership and 41 industries; margins and balance-sheet ratios by sector | not stored | **new workbook**, 373 columns, May 2014 - Aug 2026 (134 YTD points) |
| Fixed-asset investment (固定资产投资) | YTD y/y by sector and industry (table), YTD totals, infrastructure/industry/region growth (text) | not stored | **new workbook**, 49 columns, Feb 2013 - Aug 2026 (totals from 2013; the y/y table from 2018) |
| Retail sales (社会消费品零售总额) | month and YTD value and y/y: total, urban/rural, catering, 16 goods categories incl. petroleum products, automobiles | not stored | **new workbook**, 104 columns, Sep 2013 - Aug 2026 (+13 Jan-Feb) |

## 3. Not built (and why)

* PMI (中国采购经理指数, 60 releases): diffusion indices, not energy series - asked of the owner; 70-city housing prices, real-estate
  development, GDP: not energy series.
* Energy release prose items: daily averages (derivable), two-year averages and coal/Brent prices (2021 only, third-party sources).
* Seasonally adjusted m/m of value added (only a revised vintage table in each release).
* NBS database items (626 industrial products, regional output): only on `data.stats.gov.cn` (403).
* The ten unpublished-online 10-day periods listed above.

## 4. Parser bugs fixed in this audit

* Industrial table: `天然原油` (name to 2015), numbered row labels (`一、规模以上工业增加值`), `台/套` robot unit.
* Profits: earlier industry names (`开采辅助活动`, `石油加工、炼焦和核燃料加工业`, `国有及国有控股企业`).
* Archive titles without a year: the year is taken from the neighbouring monthly releases only (annual reports published later had put a
  2015 profit release in 2014 in a first run; that workbook was rebuilt).
* 10-day prices: 4 products of 2014-2019 (gasoline 93#/97#, composite cement 32.5/32.5R) had no column.
