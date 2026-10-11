# Gas demand by sector outside the Americas / ANZ / Singapore / Ireland: survey

Date: 2026-10-11. Branch: `worktree-agent-a0a06e8e8f62443d0`.

## Read this first: what is verified and what is not

**Nothing in this survey was verified live.** Every probe from the Claude sandbox through the agent proxy was
refused (`CONNECT` answered 403, a host allow-list policy) for every host tried (60 URLs): Eurostat, ENTSOG, JODI, PPAC,
METI, e-Stat, KOGAS, NBS, EPPO, OGRA, Petrobangla, gov.uk, NESO, National Gas, BNetzA, THE, ODRE (GRTgaz),
Snam, CBS, Fluxys, SSB, BOTAS, EPIAS, EPDK, Rosstat, GASTAT, FCSC, CAPMAS, NUPRC, Sonelgaz, Ember, Energy
Institute and IEA. Only `raw.githubusercontent.com` answered (OWID energy CSV, not useful for sectors). The probe
is `GAS_SECTOR_WORLD_PROBE.py` (this folder): 60 URLs blocked, 1 reachable.

So the "URL / frequency / history / sector breakdown" columns below come from the author's background knowledge of
each publisher, not from fetching it. Treat each as a lead to confirm. The manual workflow
`discovery_archive/workflows/world_gas_sector_probe.yml` re-runs the probe from GitHub Actions (where these hosts
are reachable) and dry-runs the two new pulls; paste its log back here to turn the "sandbox: no" column into real
status codes.

Volumes are rough 2023 gas consumption in bcm, order of magnitude only, to rank by size.

Availability grade: **A** raw, free, machine-readable, monthly or better, with a sector split. **B** raw and free
but needs a key, a PDF, only part of the sector split, or annual. **C** no raw sector series found; only totals
(JODI/EI/Ember fallback) or none.

## Ranked table (volume x availability)

| # | Country / region | bcm | Grade | Best raw source and URL (unverified) | Frequency / history | Sector split | Sandbox |
|---|---|---|---|---|---|---|---|
| 1 | **EU-27 + Norway** (+ UK to 2019) | ~330 | **A** | Eurostat `nrg_bal_c` (annual) and `nrg_cb_gasm` (monthly) via `https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_bal_c?siec=G3000&unit=TJ&geo=DE`; same for `nrg_cb_gasm` | Annual from 1990 (lag ~11 months); monthly from 2008 (lag ~2 months) | Annual: power/heat, industry, transport, households, commerce+public, agriculture, non-energy, energy sector, losses (nrg_bal_c). Monthly: whichever balance items member states report (gross inland deliveries, power input, possibly final-use groups); read from the response | No (403); **built** |
| 2 | Russia | ~470 | C | Rosstat / Minenergo monthly production only; sector data restricted since 2022 | n/a | none | No |
| 3 | China | ~400 | C (B annual) | NBS monthly energy release (production, `data.stats.gov.cn`), NDRC monthly "apparent consumption" total; sector split only in the annual China Energy Statistical Yearbook / NBS energy balance | Monthly total; sector annual, ~1-year lag | Annual only (industry, transport, residential, commerce, power); repo already pulls NBS production | No |
| 4 | Iran | ~240 | C | NIGC / Ministry of Petroleum: no machine-readable series found; JODI monthly total only | JODI monthly | none | No |
| 5 | Saudi Arabia | ~120 | C | GASTAT annual energy statistics; JODI monthly total; Aramco annual report | Annual | none usable | No |
| 6 | **Japan** | ~95 | B | METI "Gas Business Statistics" (city gas sales by household / commercial / industrial / other, monthly) and METI power survey (LNG burnt by utilities), both distributed through e-Stat `https://www.e-stat.go.jp/en` (API needs a free `appId`) | Monthly, history 10+ years, lag ~2 months | City gas by use + LNG for power; not total gas by sector | No |
| 7 | UAE / Qatar / Oman / Kuwait | ~70 / 45 / 20 / 25 | C | FCSC, MEW, KAPSARC (earlier discovery found only power data); JODI total | Annual | none | No |
| 8 | **India** | ~65 | A- | PPAC "Natural gas consumption" (sector-wise, MMSCM): `https://ppac.gov.in/natural-gas/consumption` (file URLs not known) | Monthly, from ~2010 | Fertiliser, power, CGD, refineries, petrochemicals, sponge iron, other | No; **built (layout-tolerant)** |
| 9 | UK | ~60 | A- | National Gas Data Item Explorer `https://data.nationalgas.com/` (daily NTS demand: power stations, industrial, LDZ offtake); DESNZ Energy Trends table 4.1 quarterly (`https://www.gov.uk/government/statistics/gas-section-4-energy-trends`); NESO data portal `api.neso.energy` | Daily (National Gas, recent years); quarterly DESNZ from 1998 | Power / industry / LDZ (domestic+commercial) daily; DESNZ adds sector detail quarterly | No; next candidate |
| 10 | Egypt | ~60 | C (B annual) | Ministry of Petroleum monthly bulletin (PDF), CAPMAS annual yearbook (gas by sector); JODI monthly total | Annual / PDF | Power, industry, household in annual yearbook | No |
| 11 | Turkey | ~50 | B+ | EPDK monthly natural gas market report and BOTAS monthly consumption bulletin (power / industry / residential); EPIAS gas transparency platform `https://seffaflik.epias.com.tr/` (repo already uses EPIAS for power) | Monthly; EPIAS daily | Power, industry, residential, fertiliser | No |
| 12 | South Korea | ~50 | B | KOGAS monthly sales by use (power, city gas split residential / commercial / industrial) via data.go.kr API (key) or KESIS `https://www.kesis.net/` | Monthly | Power vs city gas by segment | No |
| 13 | Thailand | ~50 | A- | EPPO natural gas consumption by sector `https://www.eppo.go.th/index.php/en/en-energystatistics/ng-statistic` (xls) | Monthly, from the 1990s | Power, GSP/petrochemical feed, industry, NGV, other | No; next candidate |
| 14 | Uzbekistan / Kazakhstan / Turkmenistan | ~50 / 20 / 40 | C | National statistics bureaux (stat.gov.kz bulletin has a gas balance); no sector API found | Monthly bulletins | partial | No |
| 15 | Algeria | ~45 | C | Sonelgaz / Ministry of Energy annual | Annual | none | No |
| 16 | Pakistan | ~40 | B- | OGRA monthly gas sales by sector (power, fertiliser, industry, commercial, domestic, CNG, cement), SNGPL / SSGC monthly (PDF) | Monthly PDF | Full split, PDF only | No |
| 17 | Bangladesh | ~30 | B | Petrobangla daily and monthly production / supply reports (power, fertiliser, captive, industry, CNG, domestic) (PDF) | Daily / monthly PDF | Full split, PDF | No |
| 18 | **Taiwan** | ~24 | A- | MOEAEA Energy Statistics Monthly (natural gas by sector) and data.gov.tw open data `https://www.moeaea.gov.tw/` | Monthly, long history | Power, industry, residential / commercial | No; next candidate |
| 19 | Nigeria | ~20 | C | NUPRC monthly fiscal and production report, NNPC monthly report (PDF; domestic gas supply to power vs industry) | Monthly PDF | Power vs other, partial | No |
| 20 | South Africa | ~5 | C | No gas balance (Sasol, Transnet, Egypt-style PDFs only) | n/a | none | No |

Country-level detail for Europe (inside row 1 and as follow-ups): Germany BNetzA gas consumption (SLP households
and small business vs RLM industry, daily from 2022, bundesnetzagentur.de) and THE; France GRTgaz / Teréga on
ODRE open data (`odre.opendatasoft.com`, daily consumption incl. by activity sector, Opendatasoft API, grade A
if the dataset names hold); Italy MASE monthly "bilancio gas" (civil / industrial / thermoelectric) and Snam;
Spain Enagas daily demand (conventional vs power generation); Netherlands CBS StatLine gas balance (00372) and GTS;
Belgium Fluxys (consumption by distribution / industry / power). The repo's earlier `EU_TSO_SECTOR_DISCOVERY.py`
found only ENTSOG "adjacent system" categories for several TSOs, so Eurostat is the uniform source and national
TSOs are second-pass additions for daily resolution.

Fallbacks, to be labelled as such whenever used: JODI-Gas (monthly CSV, ~90 countries, `jodidata.org`,
production / trade / stock / total demand, **no sector split**); Ember monthly gas-fired generation (power-sector
proxy, 80+ countries); Energy Institute Statistical Review (annual totals, gas-fired power); IEA monthly gas
statistics (paid).

## Discrepancy to resolve

`WORLD_COVERAGE_MAP.py` and `asia/SOUTH_SOUTHEAST_ASIA_COVERAGE_MAP.py` mark India (PPAC), Thailand (EPPO/PTT)
and Bangladesh (Petrobangla) green for "demand by sector", but this checkout and `origin/main` contain no pull
script, workflow or workbook for them (only the coverage-map text). Either those pulls live somewhere not synced
here (private repo / another branch) or the maps are ahead of the data. Check before merging the new India script
so it does not duplicate existing work.

## Step 2: what was built

1. **`europe/EUROSTAT_GAS_BY_SECTOR.py`** -> `output/Data and Chart Outputs/eurostat_gas_by_sector.xlsx`
   (Eurostat annual sector balance + monthly gas; EU-27, 27 members, Norway, UK, candidates). Incremental per
   country, atomic `xlsx_notes.write_workbook`, Units tab. Reads the unit and balance codes from the response and
   prints which expected sector codes are missing. Workflow `.github/workflows/eurostat_gas_by_sector.yml`
   (1st/15th 06:20 UTC). Charts via `add_charts.py` registry `eurostat_gas_by_sector.xlsx`: EU-27, DE, IT, FR, NL,
   ES, PL, UK annual stacked sector charts (years) and an EU-27 monthly balance chart (mmm/yy).
2. **`asia/INDIA_PPAC_GAS_BY_SECTOR.py`** -> `india_ppac_gas_by_sector.xlsx`. Layout-tolerant PPAC reader; skips
   files whose Last-Modified / size is unchanged (recorded on a "Source files" tab); refuses to overwrite on a parse
   failure. Workflow `india_ppac_gas_by_sector.yml` (1st/15th 06:35). Registry entry `india_ppac_gas_by_sector.xlsx`.
3. Probe and tests in `discovery_archive/world/`: `GAS_SECTOR_WORLD_PROBE.py`, `TEST_EUROSTAT_JSONSTAT.py`,
   `TEST_INDIA_PPAC_PARSER.py` (synthetic-data plumbing tests; they prove the code path and chart build, not the
   live source), manual workflow `discovery_archive/workflows/world_gas_sector_probe.yml`.

### Unverified items (all of the above)
- Eurostat dataset ids, the `siec=G3000` filter, unit codes (`TJ_GCV` / `TJ`), the annual balance codes
  (`TI_EHG_E`, `FC_IND_E`, `FC_OTH_HH_E`, `FC_OTH_CP_E`, `FC_TRA_E`, `FC_NE`, ...) and which monthly items exist
  were taken from memory of the Eurostat API, not seen. Whether `nrg_cb_gasm` carries a final-use sector split is
  genuinely unknown; the annual `nrg_bal_c` is the guaranteed sector split.
- PPAC file URLs and layout are unknown; the parser assumes a header row of month labels and a sector label in the
  first or second column.
- Run the probe workflow first, then the two pull workflows once by hand, and read the logs.

### Where they belong (do not register in the Americas masters)
Neither dataset belongs in `SOUTH_AMERICA_MASTER.py` / `NORTH_AMERICA_MASTER.py`. Add a future
`europe/EUROPE_MASTER.py` -> `europe_master.xlsx` (Eurostat; DATASETS list + `SOURCES` entry "Eurostat nrg_bal_c /
nrg_cb_gasm") and an `asia/ASIA_MASTER.py` -> `asia_master.xlsx` (India PPAC, then Thailand EPPO, Taiwan MOEAEA),
both reusing the master code in `south_america/SOUTH_AMERICA_MASTER.py`, each with its `add_charts.py` REGISTRY
entry (already added) and a SOURCES entry (publisher + link). Add the new scripts to the public-repo sync only
through the normal flow (develop in Energy-public; this branch is local only).

### Recommended next (in order)
UK National Gas daily demand by category, Thailand EPPO, Taiwan MOEAEA, Turkey EPDK / EPIAS, Japan METI via e-Stat,
Italy MASE bilancio gas, France ODRE, then Bangladesh / Pakistan PDFs.
