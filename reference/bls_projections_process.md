# BLS Employment Projections — browser-assisted annual check

Unlike wage data (see `fetch_bls_wages.py`, which hits `api.bls.gov` directly), BLS's
**Employment Projections** program has no equivalent public API. Its data only lives on
`www.bls.gov` pages and an XLSX bulk download, both of which block plain scripted requests
(Akamai bot detection returns 403 for `curl`/`requests`, even with realistic browser headers).
A real browser gets through fine. Decided 2026-08-18: do this one via the browser tool once a
year rather than adding a Playwright dependency for a once-a-year task.

## How to check it

1. Open **https://www.bls.gov/emp/tables/occupational-projections-and-characteristics.htm**
   ("Table 1.2 Occupational projections and worker characteristics") in the browser tool.
2. The whole table renders server-side in the initial HTML (confirmed via network-request
   inspection — no separate JSON/XHR data call to intercept instead).
3. Extract the row for the SOC code you need via JS, e.g.:
   ```js
   const rows = Array.from(document.querySelectorAll('tr'));
   const match = rows.find(tr => tr.textContent.includes('17-2031'));
   Array.from(match.querySelectorAll('td, th')).map(c => c.textContent.trim());
   ```
4. Columns (per the page's header row): occupation title, SOC code, occupation type, Employment
   2024 (thousands), Employment 2034 (thousands), Employment distribution % 2024, % 2034,
   Employment change numeric 2024-34 (thousands), Employment change percent 2024-34, % self
   employed 2024, Occupational openings 2024-34 annual average (thousands), Median annual wage
   2024, typical education, work experience, on-the-job training, OOH link.
5. The page title itself states which 10-year cycle is currently live (e.g. "Occupational
   projections, 2024–2034") — check this first since it tells you immediately whether BLS has
   already rolled to the next cycle without needing to compare individual numbers.

## Release cadence

Per `reference/IPEDS and BLS Data Refresh Cadence.xlsx`: a new 10-year projection cycle
(e.g. 2025-2035) releases "Late Aug" of the year after the cycle's start year. So the 2025-2035
cycle is expected **Late August 2026** — check back after that date for the next refresh.

## Baseline (SOC 17-2031, Bioengineers and Biomedical Engineers)

Checked 2026-08-18, page still showed the **2024-2034 cycle** (2025-2035 not yet released):

| Metric | Value |
|---|---|
| Employment, 2024 | 22,200 |
| Employment, 2034 | 23,300 |
| Employment change, 2024-34 | +1,100 (+5.2%) |
| Occupational openings, 2024-34 annual avg | 1,300 |
| Median annual wage, 2024 | $106,950 |

This **matches** the pillar article's currently-published figures ("5% growth," "about 1,300
openings a year") exactly — no update needed yet. Re-check after Late Aug 2026 for the 2025-2035
cycle; if the numbers change meaningfully, that's what should replace this baseline table and
get logged in the BME Statistics doc's Update Log (ask Lydia first per the out-of-scope-write
rule before touching that file).
