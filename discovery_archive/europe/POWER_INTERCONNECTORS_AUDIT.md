# Power interconnector audit (Europe, Oct 2026)

Probe: `POWER_INTERCONNECTORS_PROBE.py` (Elexon FUELHH per cable, ENTSO-E A11 for candidate pairs, Elering per-link flows). Data: `data/gb_interconnectors_daily.csv`, `data/elering_electricity_daily_gwh.csv`. The Energinet part timed out on rate limits (not used; ENTSO-E Denmark already matches Energinet, see KNOWN_GAPS).

ENTSOE_FLOWS_DAILY.py queries ~94 zone pairs (both directions), including GB (EXTERNAL node), Turkey, Ukraine (3 zones), Moldova, Russia, Belarus. Net imports = flows in less flows out over all pairs of a country.

## GB cables: Elexon (GB metered end) vs ENTSO-E (far-end, net import into GB), TWh

| Cable(s) (zone pair) | Source | 2021 | 2022 | 2023 | 2024 | 2025 | Verdict |
|---|---|---|---|---|---|---|---|
| IFA + IFA2 + ElecLink (FR-GB) | Elexon | 13.43 | -9.90 | 12.76 | 19.45 | 21.94 | OK, ENTSO-E +0.4-0.5 TWh (2.2%) in 2024-25 = losses |
| | ENTSO-E | 13.36 | -9.47 | 12.75 | 19.88 | 22.41 | |
| Nemo Link (BE-GB) | Elexon | 6.85 | 0.65 | 2.98 | 4.16 | 1.99 | OK (<0.15 TWh; 2-6% loss) |
| | ENTSO-E | 7.01 | 0.77 | 3.10 | 4.30 | 2.11 | |
| BritNed (NL-GB) | Elexon | 4.25 | 1.43 | 2.68 | 1.59 | 0.44 | OK (<0.11 TWh) |
| | ENTSO-E | 4.34 | 1.55 | 2.79 | 1.68 | 0.54 | |
| Viking Link (DK1-GB, Dec 2023) | Elexon | - | - | 0.04 | 3.66 | 2.53 | OK, captured both sides |
| | ENTSO-E | - | - | 0.04 | 3.74 | 2.57 | |
| North Sea Link (NO2-GB, Oct 2021) | Elexon | 1.39 | 2.73 | 8.53 | 9.62 | 8.97 | OK; ENTSO-E +0.3 TWh (3-4%) = losses (NSL ~1.4 GW) |
| | ENTSO-E | 1.41 | 2.81 | 8.57 | 9.92 | 9.31 | |
| Moyle + EWIC + Greenlink (IE(SEM)-GB) | Elexon | -1.36 | 0.86 | -1.15 (163 d) | -4.95 | -6.48 | 2025: ENTSO-E 1.0 TWh short (Greenlink unreported Feb-May 2025); 2023 ENTSO-E has no data Apr-Nov |
| | ENTSO-E | -1.35 | 0.84 | -1.14 (163 d) | -4.83 | -5.48 | |

Per cable (Elexon, TWh, + = into GB): IFA 9.3/-3.0/5.9/10.1/10.3; IFA2 4.3/-3.9/3.1/4.2/5.7; ElecLink 0.0/-3.1/3.8/5.2/5.9; Nemo 6.9/0.7/3.0/4.2/2.0; BritNed 4.3/1.4/2.7/1.6/0.4; NSL 1.4/2.7/8.5/9.6/9.0; Viking 0/0/0.05/3.7/2.5; Moyle -0.9/0.3/-2.0/-2.5/-2.2; EWIC -0.5/0.6/-1.7/-2.7/-1.7; Greenlink 0/0/0/0/-2.8. EirGrid agrees with Elexon for 2025 (EWIC 1.70, Greenlink 2.81, Moyle 2.17 TWh).
GB total (Elexon NetImports) vs sum of the six ENTSO-E mirror borders: 2021 24.8/24.8, 2022 -4.2/-3.5 (IE gap), 2024 33.3/34.7, 2025 29.1/31.5 (losses 2-3%, plus none double counted: GB is not in the ENTSO-E generation set, so the continental net imports count each GB cable once).

## Other links

| Link | Check | Verdict |
|---|---|---|
| Estlink 1/2 (FI-EE) | Elering vs ENTSO-E 2021-25: 6.54/6.64, 6.70/6.80, 6.75/6.85, 3.55/3.63, 4.54/4.62 | match (1-1.5%, losses) |
| Estonia-Latvia | Elering vs ENTSO-E: identical to 0.01 TWh each year | OK |
| Estonia-Russia (Narva, Pihkva) | identical; ends 2025 (Baltic synchronisation Feb 2025) | OK |
| NordBalt (SE4-LT), LitPol (PL-LT), Lithuania-Belarus/Kaliningrad | ENTSO-E both zones; BY/RU-KGD flows stop Feb 2025 (35 days 2025), handled as ended links | OK (no second source; see Lithuania KNOWN_GAPS) |
| Skagerrak (DK1-NO2), Kontek (DK2-DE), Baltic Cable (SE4-DE), COBRA (DK1-NL), NordLink (DE-NO2), NorNed (NL-NO2), Great Belt/Oresund, Konti-Skan | queried, 365 days every year | OK, no second source here (Denmark nets match Energinet, 7.4 TWh 2025) |
| ALEGrO (DE-BE, Nov 2020) | inside DE-LU>BE / BE>DE-LU, data every day 2021-26 (DE>BE 2.6/3.3/2.5/1.8/1.7 TWh) | captured |
| Italy-France (Piossasco-Grande Ile), Italy-Switzerland/Austria/Slovenia | queried both ways | OK |
| SACOI (Sardinia-Corsica-Tuscany) | internal to Italy | n/a |
| Montenegro-Italy (MONITA) | IT-Centre-South>ME / ME> data every day | captured (ME-BA mirror issue documented) |
| Italy-Greece | IT-South<->GR | captured |
| Sicily-Malta | NOT queried before; ENTSO-E IT-Sicily>MT 0.54/0.64/0.64/0.97/1.01 TWh, MT>IT 0-0.03 | HOLE, FIXED (Italy net imports fall by that amount) |
| Slovakia-Hungary, Hungary-Romania, Bulgaria-Greece, Poland-Germany, Poland-Lithuania, Poland-Sweden (SwePol), Turkey-Bulgaria/Greece, Hungary-Slovenia (from mid-2022) | queried, data all days | OK |
| Ukraine/Moldova-Romania/Poland/Slovakia/Hungary | three Ukraine zones merged by max; RO>MD from Mar 2022 | OK; SK-UA gap Jul-Sep 2022 filled by the UA-IPS zone |
| Spain-Morocco/Andorra | no ENTSO-E zone (A11 no data); REE added to Spain's net imports | OK |
| Norway-Russia, Latvia-Belarus, Sardinia/Corsica-France, Greece-Cyprus, Italy-Albania, Montenegro-N.Macedonia | ENTSO-E has no data either way | no link / not published |
| Celtic (FR-IE, 2027), Harmony Link (PL-LT), Biscay Gulf (FR-ES, 2028), ELMED (IT-Tunisia) | not yet in service | n/a |
| Aurora Line (SE1-FI, 2025) | inside SE1>FI | captured |
| Switzerland | BFE monthly net imports (-6.39 / -14.40 / +0.02 TWh 2023-25) vs ENTSO-E CH borders (-5.06 / -12.05 / +0.38) | differ by 1.3 / 2.4 / 0.4 TWh: BFE is physical incl. Liechtenstein and its monthly totals; kept (owner choice), indicative |

## Fixes
1. `ENTSOE_FLOWS_DAILY.py`: added Malta (zone `10Y1001A1001A93C`) as an external node and `IT-Sicily: MT`; the entsoe_flows workflow was re-run, Borders now has `IT-Sicily>MT` / `MT>IT-Sicily`. Italy net imports 2021-25 fall by 0.52 / 0.64 / 0.62 / 0.97 / 1.03 TWh (2025: 48.64 -> 47.61), supply/load -0.2 to -0.3 points (ENTSO-E basis about -0.35).

## Cannot be closed
- Greenlink Feb-May 2025 (about 0.9 TWh) and IE(SEM)-GB Apr-Nov 2023 are missing in ENTSO-E. The flows workbook's Ireland column treats gaps over 14 days as zero, but the master takes Ireland's net imports from EirGrid demand less generation and GB's from Elexon, so no balance is affected.
- Northern Ireland (Moyle, about 2.2 TWh a year from GB) is in no country's balance; GB shows it as an export, correctly.
- Cable losses (2-4%) separate GB-end and far-end metering; each country's own side is used for its balance.
