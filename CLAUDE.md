# Energy data repo - working conventions

Standing instructions from the repo owner. Follow these on every change.

## Branches
- Develop on `github-access-energy`; production is `main`. Every change goes to BOTH branches.
- External data sites are mostly blocked from the Claude sandbox: run pulls in GitHub Actions.
- Pulls run in the PUBLIC repo thomasdalton9/Energy-public (free Actions minutes); make code changes there.
  Its main is copied into the private thomasdalton9/Energy (main + github-access-energy) on the 1st and 15th (23:50 UTC),
  by .github/workflows/sync_to_private.yml (needs the PRIVATE_REPO_TOKEN secret). Never copy private-repo
  content into the public repo.

## Spreadsheets and charts
- **Every xlsx we pull must include native Excel charts of its data**, redrawn by the
  script on every run (openpyxl charts inside the workbook, not just a separate PNG).
  - Storage / reservoir / water-level series: AGSI-style water-year chart via
    `water_year_chart.add_water_year_chart()` (5-year min-max band, 5Y average,
    previous and current water year).
  - Other time series: line or stacked chart of the main columns on a chart sheet.
  - Charts have **no borders** (no chart-area or plot-area outline).
  - Chart dates are formatted **mmm/yy** (e.g. Jan/26); annual series show the year.
  - `add_charts.py` holds the per-workbook chart registry; each scheduled workflow runs it after the pull.
- Storage/level/seasonal charts use an **Oct-Sep water year**, not Jan-Dec.
- Outputs: xlsx in `output/Data and Chart Outputs/`; PNG charts in `output/PNG Charts/`; a copy of each producing
  script in `output/Python Scripts/` (refreshed by the workflow on every run).
- Every new or updated PNG chart is committed to `output/PNG Charts/` on main and sent to the owner.
- Units/notes tab via `xlsx_notes.write_workbook()` (atomic write).

## Master workbooks
- `south_america/SOUTH_AMERICA_MASTER.py` -> `south_and_central_america_master.xlsx`: South America plus Central
  America (Guatemala-Panama + Belize; NOT Mexico) and the Caribbean: Trinidad & Tobago, Puerto Rico, Jamaica,
  Dominican Republic. Gas Dashboard + "Dashboard - Power & Hydro", a data tab per
  chart, raw data tabs, Sources tab. New datasets: add them to its DATASETS / RAW_POWER_DATASETS / HYDRO_DATASETS
  lists (and add_charts.py REGISTRY), plus a SOURCES entry (publisher + link).
- `americas/NORTH_AMERICA_MASTER.py` -> `north_america_master.xlsx`: United States, Canada, Mexico, same layout
  (reuses SOUTH_AMERICA_MASTER's code). Gas Dashboard + "Dashboard - Power"; add datasets to its DATASETS /
  RAW_POWER_DATASETS / OTHER_POWER_DATASETS / CAPACITY_DATASETS lists plus a SOURCES entry.
  US power generation comes from EIA-930 (balancing-authority data, current to yesterday) - owner's decision; don't
  switch the US regions to direct ISO feeds.
- `asia/SOUTH_SOUTHEAST_ASIA_MASTER.py` -> `south_southeast_asia_master.xlsx`: South Asia (India, Pakistan, Bangladesh,
  Sri Lanka, Nepal, Bhutan) and Southeast Asia (Thailand, Vietnam, Philippines, Indonesia, Malaysia, Singapore,
  Myanmar, Cambodia, Laos, Brunei, Timor-Leste) - NOT Japan, Taiwan, Korea or China. Same layout. Add datasets to its
  DATASETS / RAW_POWER_DATASETS / OTHER_POWER_DATASETS / HYDRO_DATASETS lists plus a SOURCES entry. Ember fallback:
  monthly (`south_southeast_asia_power_by_type.xlsx`) and, where Ember has no monthly data, yearly
  (`..._annual.xlsx`, EMBER_POWER_BY_TYPE.py --yearly). GSO covers Peninsular Malaysia only, so the regional total
  keeps Ember's national Malaysia (TOTAL_USE_EMBER). Singapore uses raw data (EMC/NEMS metered generation topped up to the EMA/SingStat monthly total;
  singapore_power_generation_daily.xlsx from SINGAPORE_POWER.py). The Ember pull keeps every country, raw ones included,
  as an unused fallback (one whole-file download, about a minute); the master uses the raw workbook wherever it exists.
  Malaysia: GSO is Peninsular only; the master shows Sabah + Sarawak as an ESTIMATE (Ember national minus GSO, labelled) -
  Sarawak Energy blocks GitHub, SESB times out, and the Energy Commission's regional tables stop at 2021.
  Gas: India (PPAC), Singapore (EMA/SingStat + a daily gas-for-power estimate from NEMS CCGT output), Thailand
  (EPPO tables 3.1-1 / 3.2-2: production by field, Myanmar pipeline, LNG, use by sector, monthly from 1986) and
  Bangladesh (Petrobangla daily gas production & distribution report PDFs, from 2021; dated by the START of the 08:00
  gas day - owner's decision, it lines up with the power data), Pakistan (PBS bulletin table 3.2 production by
  province, to Mar 2024 - PBS stopped posting; LNG import value only), Indonesia (Ditjen Migas Buku Statistik Migas,
  half-yearly, ~6-18 month lag) and Vietnam (NSO monthly tables, from 2015). Malaysia (DOSM quarterly only) and the
  Philippines (nothing monthly) have no gas pull.
  Raw power by country (no copied or invented months: a month a raw feed lacks is a gap, and the master's regional
  total fills it from Ember, labelled): India = CEA daily (NPP) for coal/gas/oil/nuclear/hydro + NITI Aayog ICED for
  wind/solar/other RE and demand (INDIA_RE_DAILY.py; Grid-India/MERIT geo-block GitHub); Vietnam = EVN daily posts of
  NSMO data (VIETNAM_EVN.py, from May 2023); Philippines = IEMOP 5-minute dispatch per plant mapped to fuels with DOE's
  plant list (asia/philippines_resource_fuels.csv; IEMOP keeps ~90 days, history grows from Jul 2026); Pakistan =
  CPPA-G / NEPRA monthly FCA filings (national grid) plus K-Electric only for months with a real KE figure
  (KE_included; others count as gaps); Nepal = NEA LDC daily reports (Apr 2023 - Jan 2025; monthly reports back to Jul 2022) + the daily home-page panel
  (NEA posted nothing Feb 2025 - Sep 2026, so that span is an Ember-filled gap);
  Cambodia = EAC annual reports (annual rows); Indonesia, Myanmar, Laos, Brunei, Timor-Leste = Ember annual (no public
  official sub-annual data found; Indonesia: ESDM EBTKE monthly renewable capacity only).
  Charts and the regional total use only months a feed covers (>= 80% of days; add_charts.complete_months).
  Capacity: India (CEA monthly), Bangladesh (BPDB), Sri Lanka (PUCSL), Philippines (DOE annual), Malaysia (GSO),
  Cambodia (EAC); capacity factors are blank above physical ceilings (solar 32%/28%, wind 65%/55%: incomplete plant
  lists). Prices: India IEX, Philippines IEMOP (final prices only on charts), Malaysia SMP, Singapore USEP.
  Reservoirs: India CEA, Thailand RID, Philippines PAGASA, Pakistan IRSA/WAPDA (Tarbela, Mangla), Sri Lanka PUCSL.
  Geo-blocks (check-host test from ~60 countries, discovery_archive/asia/GEO_DISCOVERY*.py): EVN reservoirs
  (hochuathuydien.evn.com.vn) and NSMO answer only from Vietnam; Grid-India and MERIT only from India; BPS only from
  Indonesia; web.pln.co.id no longer resolves anywhere. A self-hosted runner or proxy in the country would unlock
  them (parked by the owner for now).
- `europe/EUROPE_MASTER.py` -> `europe_master.xlsx`: Europe - the ENTSO-E countries (EU/EEA, UK excluded, plus the Balkans) plus
  Ireland EirGrid, Turkey, Cyprus and the Rhine at Kaub. Gas Dashboard (GIE AGSI+ storage and ALSI LNG, Ireland GNI),
  "Dashboard - Power" (generation, capacity, day-ahead prices, net imports, and a supply/demand balance per country: generation +
  net imports + pumped storage/batteries vs load) and "Dashboard - Capacity factors". Country list and bidding zones live in
  `europe/europe_countries.py` (add a country there and the pulls, registry and master pick it up). ENTSO-E is the primary
  power source for Europe (it is the TSOs' own statutory reporting, and matched SMARD/RTE within 0.5% in a settled-week check),
  labelled as such; replace it per country with a national feed where that is better. Great Britain is not in ENTSO-E generation: it is pulled from
  Elexon BMRS + NESO (`GB_POWER_DAILY.py`, incl. embedded wind/solar and interconnector flows) and its gas from National Gas NTS
  (`GB_GAS_NTS_DAILY.py`). GB storage flows come from the nine sites' own daily stock/inflow/outflow on the same portal (`GB_STORAGE_SITES_DAILY.py`, site items only from Aug-Oct 2024; `portal_total_stock` from Oct 2021 = sum of sites): the NTS aggregate storage_withdrawal/injection overstates net withdrawals by ~5-10 TWh a year (it implies +33 TWh net withdrawal 2022-25 while the stock actually rose ~1.5 TWh), so the master uses the site totals, and before Oct 2024 the day-to-day change in the portal's total stock (net only). With true storage the UK balance runs 2.6-4% short (was 1.5-2.7%): the NTS excess had been masking other missing supply. Rough (stopped injecting 2025, empty since Oct 2025) supplied 12 of the 13.4 TWh Oct 2024-Sep 2025 net withdrawal. Ireland's ENTSO-E all-island feed covers only part of demand, so the Europe totals use Ember's monthly data for the
  Republic of Ireland (`EMBER_EUROPE_MONTHLY.py`, labelled fallback); Albania has no Ember rows and ENTSO-E only from May 2026, so it is
  left out of the totals. CGES (Montenegro) is not yet pulled. Months where a country's feed was incomplete (Sweden before Dec 2021) are
  dropped, so the Europe totals start Dec 2021.
  Gas: ENTSOG physical flows (pipeline imports by origin, exports, production, consumption) with ALSI LNG and
  AGSI+ storage give a gas balance per country and for the EU.
  National consumption comes from the gas operators' own series where ENTSOG's country totals are partial (raw data only, no
  Eurostat gap-filling; Eurostat `eurostat_gas_monthly.xlsx` is a validation benchmark and not charted): Germany (THE), France
  (ODRE; industrial + public distribution only - the CCCG power-plant series is a subset of industrial), Spain (Enagas daily from
  2023, Enagas monthly bulletin spread over the days for 2021-22), Denmark (Energinet), Portugal (REN) in
  `TSO_GAS_DEMAND_DAILY.py`; Austria (AGGM), Czechia (NET4GAS CAMS, a border/storage/production system balance, not a consumption
  series), Lithuania (Amber Grid) in `GAS_TSO_CEE_DAILY.py`; Poland (Gaz-System), Romania (Transgaz), Croatia (Plinacro),
  Finland (Gasgrid), Spain monthly in `GAS_TSO_SOUTHEAST_DAILY.py`; Great Britain NTS and Ireland GNI as above. Poland, Romania
  and Croatia stay on ENTSOG in the balances: the operators' exits equal ENTSOG's and the 8-10% gap to Eurostat is domestic
  production consumed off-grid (also absent from ENTSOG production). Italy stays on ENTSOG (within 4% of
  Eurostat; Snam blocks GitHub). The Netherlands takes production and total consumption from CBS StatLine 86103NED (`NETHERLANDS_CBS_GAS.py` -> `netherlands_cbs_gas_monthly.xlsx`, monthly, spread over the days; ENTSOG's Dutch production runs 15-17 TWh a year above CBS and its consumption exits ~6 TWh below; months CBS lacks use ENTSOG times the last-12-month CBS/ENTSOG ratio), which took the DE+NL balance from +4% to about +2%; the rest is mainly Germany's Eurostat-sourced biomethane line (~12 TWh, which THE's consumption may already contain) and ENTSOG own-side differences on the DE-NL border (under 10 TWh, they cancel in the combined block); LNG send-out and net imports match Eurostat/CBS within noise. Bulgaria has no raw national series. ENTSOG border flows (`ENTSOG_GAS_FLOWS_DAILY.py`): a virtual point (VIP) and the physical points it aggregates are taken as the larger of the two, not summed (VIP Brandov = EUGAL + OPAL + Hora Svate Katerina doubled Czech imports in 2025); operators reporting the same gas at one point (both German TSOs at Ueberackern, Wallbach) count once; each country keeps its own side of a border and the other side is used only where the own country publishes no row for those points (Baumgarten has no Austrian-side row; a reported zero or a differently labelled neighbour, e.g. Komotini IGB / Kulata, is not filled - doing so put Bulgaria 55 points out). A `Border flows` sheet (`NL>DE` ...) holds the larger-of-both-sides flow per border. A rule change needs `entsog_gas_flows.yml` with rebuild=true. Norway: Gassco daily flows by
  destination (`NORWAY_GASSCO_DAILY.py`, mcm/d x 11.2 GWh, from Oct 2020) give the Great Britain import line; the rest of the
  St Fergus and Easington terminals is counted as UK production.
  Norway's Gassco flows to Germany, France, Belgium and other also replace ENTSOG's Norway origin in the EU balance (ENTSOG
  captures only ~60% of Norwegian pipeline gas). Denmark's own balance comes from Energinet Gasflow
  (`DENMARK_GASFLOW_DAILY.py`: North Sea + Tyra entries incl. Norwegian gas for Baltic Pipe, biogas, storage, Germany, Sweden,
  Poland; consumption = KWhToDenmark, which already includes the biogas - do not add biogas to it again). Biomethane is its own
  supply line in the gas balances (not in ENTSOG production): France ODRE, Denmark Energinet, Netherlands CBS, Austria AGGM
  (`BIOMETHANE_DAILY.py`, `BIOMETHANE_STATS.py`), other EU27 countries from Eurostat's annual figures held at the last year;
  Great Britain's is not added because the NTS offtake we use as UK consumption excludes embedded gas. Germany and Ireland have no
  operator biomethane feed. ENTSOG does publish the Emden (EPT1) entries (OGE, GUD, GTS; Thyssengas repeats GUD) and the Nord Stream 1 entries at Greifswald (NEL, OPAL; 620 TWh in 2021, 314 in 2022), but lists no far side for them, so the country classification drops them (as with TAP and Moffat): `ENTSOG_POINT_FIXES_DAILY.py` pulls them and `point_fix_args()` adds them to Germany (Greifswald, OGE + GUD) and the Netherlands (GTS) imports, which took the separate DE and NL balances from -58%/-43% to within 3% in 2022 and DE+NL from -29%/-10% to -2%/+2% in 2021/2022 (Nord Stream alone overshoots, because the Gassco-based `emden_gap` estimate, kept only as the fallback when the Emden columns are missing, is 513 TWh in 2022 against ENTSOG's 345; Yamal gas appears only as Mallnow DE<-PL, 224 TWh in 2021, since ENTSOG has no Kondratki/Wysokoje rows). Belgium (Zeebrugge) and France (Dunkerque) capture 94-95%
  of Gassco's flows. France is about +8% and Austria far off for other, untraced reasons.
  Gas balance fixes for points ENTSOG's country classification drops (`ENTSOG_POINT_FIXES_DAILY.py` -> `entsog_point_fixes_daily.xlsx`,
  applied by `point_fix_args()` in the master): Greece's TAP entry at Nea Mesimvria (listed with country GR, so it looked like a flow inside Greece) is now booked by the border-flow pull as AL>GR, so the separate TAP point fix was removed (it counted the gas twice, +16%); Hungary subtracts the "Exit for Blending" from production (imported gas is blended
  with domestic gas and re-enters at the production entry, so production was double-counted; +14% -> ~0); Great Britain adds the Moffat exit
  to exports (ROI share = GNI's Moffat import figure; the rest, Northern Ireland + Isle of Man, less the Carrickfergus exit to the Republic (7-9 TWh a year, already in ENTSOG's UK exports and formerly counted twice), is added to UK consumption since the NTS
  offtake is GB only; ENTSOG omits Moffat because it lists the far side as country UK) and uses National Gas NTS's own storage
  withdrawals/injections (ENTSOG lacks Stublach, Holford, Hill Top entries; error +6% -> -0.4%); France consumption = ODRE offtake + biomethane
  injected into distribution (ODRE equals ENTSOG's distribution + industrial exits, which exclude embedded biomethane). Remaining FR error (~+3%) is network own use/losses
  and ENTSOG missing ~15 TWh of French exports against Eurostat; Hungary 2022 stays -11% (ENTSOG's production entry starts 2023); Greece 2021-22 +5-9%. The EU27 total does not yet
  include these four corrections.
  Austria's balance is AGGM's own market-area series (`GAS_TSO_CEE_DAILY.py`: domestic production, net border entry/exit, storage withdrawal/injection against the end-customer consumption it determines from metering) in `austria_gas_balance()`: ENTSOG left it 12-22% short in 2022-24 (no Austrian production, 5 TWh a year; no Austrian-side Baumgarten row; AGSI's Austrian storage flows incl. Haidach, fed from the German grid, differ from AGGM's by 5-13 TWh a year) and AGGM's identity closes within 0.4% in 2023-26.
  Czechia's imports and exports are floored at NET4GAS's own allocated border entries/exits (`import_floor` / `export_floor`): ENTSOG's VIP Brandov was 14 TWh in 2023 and its physical points sum to 63 TWh against NET4GAS's 79 TWh (balance -23% -> -1.6%); consumption there is NET4GAS's own system balance, so that balance closes largely by construction.
  Residual gas-balance errors after these fixes (2023 / 2024 / 2025): Great Britain -2.1 / -1.2 / -1.6% (-8 to -13 TWh a year; LNG send-out matches National Gas's own LNG importation item to 0.1 TWh, terminal entries match NTS nominations, and IUK + BBL + Moffat exits match NTS interconnector exports; what remains is unexplained, NTS offtake exceeds ENTSOG's exits by 2.5-5 TWh a year and NTS shrinkage is about 2.3 TWh); France +2.5 / +2.2 / +2.9% (ENTSOG exports 2025: Belgium 78.0, Switzerland 82.4, Spain 8.3, Germany 0.1 TWh, no other French border carries data; ODRE has no network losses or own-use series; ENTSOG's Dunkerque entry exceeds Gassco's flow to France by 16 TWh in 2025 and 45 in 2023, so the import side is also uncertain).
  Switzerland power is Swissgrid/BFE (`SWITZERLAND_SWISSGRID_DAILY.py`, replaces ENTSO-E whose Swiss hydro is incomplete): production by carrier is
  gross of pumped-storage output, so pumping consumption and physical imports/exports come from BFE's monthly electricity balance (ogd35,
  spread evenly over the days); the balance then closes within about 3%.
  Netherlands power is CBS StatLine 84575NED (`NETHERLANDS_CBS_POWER.py`: monthly production by source incl. rooftop solar, spread over the days;
  ENTSO-E's Dutch solar is under 1 TWh a year). Load is CBS consumption incl. losses (ENTSO-E's Dutch load is ~10% low in 2021-22, within 1% in
  2024-25), so the Dutch balance closes by construction; cross-border flows stay ENTSO-E (they match CBS imports/exports).
  Power balance audit (supply/load by country-year; `country_balance()` in EUROPE_MASTER.py, causes in its KNOWN_GAPS dict): ENTSO-E reports the
  same Ukraine tie-lines under three zones (UA, UA-IPS, UA-BEI), so `ENTSOE_FLOWS_DAILY.py` merges them (largest value, never sum) when building
  Net imports (`--net-only` recomputes from the saved Borders sheet; this fixed Slovakia 0.90 -> 1.00 and Hungary/Romania/Poland). Remaining gaps are
  documented, not patched: Italy (ENTSO-E generation and load both omit embedded/self-consumed power; no consistent raw pair), Poland (94-95% before 2024, 98-100% since), Romania (96% in 2025), Great Britain (+2-3%, Elexon gross of station
  load vs NESO demand net), Bulgaria (Eurostat generation equals ENTSO-E's, so the 2025 gap is load or export flows), Balkans (BA-ME physical flow does not close either side), Denmark (see KNOWN_GAPS).
  Germany power is Eurostat nrg_cb_pem monthly net generation by fuel (Destatis, all producers incl. industrial self-generation and rooftop PV; `EUROSTAT_POWER_MONTHLY.py` -> `germany_eurostat_power_daily.xlsx`, each month spread over its days, ENTSO-E daily values after Eurostat's latest month about 2.5 months back, pumped storage/load/flows stay ENTSO-E): supply is 98-99.6% of load in 2022-25, was 95-98%.
- Every master has a "Dashboard - Long-term" page (`fundamentals.py`, called from each master): an annual summary per
  country and the region (demand and its 10-year growth vs GDP growth, elasticity, demand and GDP per head, fuel shares,
  wind + solar change, net imports, capacity and fleet utilisation, gas balance and import dependence, IMF 5-year GDP
  and population outlook, degree days) and history charts from 2000 (gas from 1990). Inputs: `long_term_energy.xlsx`
  (`LONG_TERM_ENERGY.py`: Ember yearly power 2000-, Energy Institute Statistical Review gas/oil/coal/LNG 1965-; EI via
  curl_cffi, OWID fallback) and `macro_drivers.xlsx` (`MACRO_DRIVERS.py`: World Bank GDP/population/industry/
  urbanisation/access, IMF WEO growth and population incl. forecasts, CDD/HDD base 18C from NASA POWER daily
  temperature at population-weighted cities, City_T2M_daily is the incremental store). These are compiled annual
  statistics, labelled as such: the long consistent history the raw feeds are too short to give. Country lists per
  region and the name -> ISO3 map are in `fundamentals.py` (REGIONS, ISO3).
- Every dashboard chart shows its source. Prefer raw sources (grid operators, ministries, statistics offices);
  Ember is a fallback only for countries with no raw feed, labelled as such.

## Pull scripts
- Schedules: pulls whose source keeps history run on the 1st and 15th of each month (they backfill every missed
  day), to save Actions minutes. Only sources with no history run daily or more often: the Argentina AIC snapshot
  (argentina_aic_snapshot.yml; the full Argentina hydro pull is 1st/15th), Canada IESO (rolling 'today' XML),
  Ecuador CENACE daily, LNG feedgas (TC keeps no history), Turkey EPIAS ('today' only), Australia gas hub prices
  (au_sttm_prices.yml: AEMO STTM report INT651 holds about a week, DWGM about 14 days), Philippines dam levels
  (philippines_dam_levels.yml: PAGASA posts only today's and yesterday's readings), Pakistan reservoirs
  (pakistan_reservoirs.yml: IRSA keeps only ~8 daily reports), Nepal (nepal_power.yml: NEA's home-page energy panel
  has no history; the NDOR report PDFs are read only on the 1st/15th). Masters rebuild on the
  1st/15th evening, then the private-repo sync (23:50). A new pull follows the 1st/15th schedule unless its
  source has no history.
- Incremental: backfill gaps only, don't re-pull complete history each run.
  The committed workbook IS the history store: each run reads it and fetches only periods not saved yet (plus a
  short revision window). Sources that only publish whole files (e.g. MBIE webtables, Hydro Tasmania) are
  downloaded only when a new release/Last-Modified appears (recorded on the Units sheet).
- South America gas demand by sector: data from 2021 only.
- Ireland gas: GNI transparency pages from 2026-03-31 on (open data covers earlier);
  ENTSOG is for validation only.

## Housekeeping
- One-off discovery/inspect/probe/test scripts go in `discovery_archive/`
  (with their manual-only workflows in `discovery_archive/workflows/`).
