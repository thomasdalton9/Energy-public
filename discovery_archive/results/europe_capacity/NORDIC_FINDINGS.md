# Monthly installed capacity, Nordic and Baltic countries (probed 12 Oct 2026, probes NORDIC_CAPACITY_PROBE1-7)

BUILT (europe/capacity/NORDIC_CAPACITY_MONTHLY.py, workflow europe_capacity_nordic.yml)
- Denmark: Energinet Energi Data Service CapacityPerMunicipality, monthly from Jan 2016, wind on/offshore, solar, thermal by size only (no fuel split; "under development" at Energinet). PARTIAL.
- Norway: NVE hydropower and wind databases (api.nve.no), month-end capacity rebuilt from commissioning/removal dates, hydro + pumped-storage plants + onshore wind. PARTIAL.

NOT AVAILABLE MONTHLY (raw)
- Sweden: Energimyndigheten PxWeb (pxexternal.energimyndigheten.se) wind EN0105_1/_5 and solar EN0123_1 are ANNUAL; EN_IND4-1A installed capacity by fuel annual from 1996. Vindbrukskollen (Lansstyrelserna) ArcGIS times out from Actions. Energy-Charts monthly exists for Germany only.
- Finland: Energiavirasto power plant register is behind Cloudflare (403 even with curl_cffi). Fingrid datasets 267 (solar), 268 (wind) and 424 (batteries) hold installed capacity with history but the API needs a free subscription key (HTTP 401, no FINGRID key secret in the repo). Suomen uusiutuvat (wind association) lists wind farms/turbines and solar plants by YEAR only, half-yearly. Statistics Finland PxWeb has no capacity table.
- Estonia: Elering dashboard API (v3/api-docs) has no capacity endpoint; its capacity table (31 Dec) is on a script-rendered page. Statistics Estonia PxWeb not probed beyond the tree.
- Latvia: AST market-review workbook has balance/prices/GOs, no capacity; no AST capacity page found; data.gov.lv search empty.
- Lithuania: Litgrid has an installed-capacity page and an open-data set "Gamybos pajegumai (MW)" (ENTSO-E 14.1.A mirror, annual by nature); pages are script-rendered, no download link found. osp.stat.gov.lt behind Cloudflare.
- Energy-Charts monthly installed_power: "monthly data is only available for the following countries: de" - no fallback for any of the seven.
- Ember monthly wind and solar capacity page is Cloudflare-blocked from Actions.
