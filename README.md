# Portland Budget Treemap

An interactive treemap of the City of Portland's FY 2026-27 Adopted Budget.

**Live site:** https://brycelied.github.io/PortlandBudget/

## What it shows

- **Levels:** City (bureaus grouped by service area) → service area → bureau. At the bureau level, switch between a breakdown by program or by fund. Select a program to break it down by sub-program, and select any sub-program or fund to see its funding sources and spending categories.
- **Funds:** All funds, or the General Fund (Fund 100) only.
- **Budget measure:**
  - **Total budget** ($8.55B): the legal budget under Oregon Local Budget Law, including fund-level requirements such as contingency, debt service, transfers, reserves and ending fund balance.
  - **Capital** ($1.19B): spending on Capital Improvement Plan projects.
  - **Operating** ($7.35B): everything else.

## Source

City of Portland Ad Hoc Reporting Tool export, FY 2026-27 Adopted Budget (expense requirements). The data is embedded in `index.html`.

Totals tie to the FY 2026-27 Adopted Budget: Table 1 for all funds ($8,546,060,062), Figure 10 for the General Fund ($1,186,002,037), and Table 10 for the operating/capital split.

Prosper Portland shows only its City-budgeted amount.

---

# Budget Treemap with Decision Packages (mockup)

The adopted budget treemap with the FY 2026-27 Fall Supplemental decision packages (DPs) merged in.

**Live site:** https://brycelied.github.io/PortlandBudget/treemap-dp/

## What it adds to the treemap

- **Decision packages toggle:** *Base budget* (adopted only), *All DPs* (adopted plus every request; tiles show the percentage change), and *DPs only* (just the requested change; tile size is the size of the change and striped tiles are net decreases). Works with the existing fund, budget measure and bureau breakdown controls.
- **Service impact toggle:** *All requests*, *Service impact* or *No service impact* limits which DPs are counted (sample tags in this mockup). It is turned off while *Base budget* is selected.
- **Program narrative:** opening a program shows its description from the adopted budget book above the treemap. Programs with several sub-programs list each one; selecting a sub-program shows its full description and services.
- **Commitment items:** every spending category (and resource or requirement line on DP pages) expands to its commitment items, in the side panel, the DP tables and the DP pages. Amounts follow the same level and toggles; adopted commitment items tie to the treemap on every bureau, program, fund, category and sub-program combination.
- **Decision package pages:** each request has its own page (`treemap-dp/#dp-<id>`) with its full amount across every program and sub-program, spending categories, resources and requirements by fund, and positions. On program views, each request is tagged *Portion in this program* or *Whole request in this program* and links to its page.
- **Decision packages at every level:** below the treemap, requests are listed for the level you're on and follow every toggle. Citywide they are grouped by service area, then bureau; in a service area by bureau; in a bureau by program (each group row shows requests, adopted, DP change, with DPs and change, and expands to its requests). In a program, a sub-program table is followed by each request with its description, sample impact tag, a spending-category table compared with the program's adopted budget, and positions. Requests shown only in part are tagged.

Totals tie to the adopted budget ($8,546,060,062; General Fund $1,186,002,037) and the requested changes ($530,518,741; General Fund $42,711,774). DP position costs are added to Personnel as operating spending; DP lines are counted as capital when their project is a CIP project (same rule as the adopted Capital measure).

## Rebuild

```
python3 dp/build/prep.py <DP export.xlsx> dp/build/bureau_sa.csv dp.json
python3 dp/build/prep_programs.py <DP export.xlsx> <Ad Hoc.xlsx> dp/build/bureau_sa.csv dp.json dp/build/program_descriptions.json prog.json
python3 treemap-dp/build/prep_treemap.py index.html <DP export.xlsx> <Ad Hoc.xlsx> prog.json tmdp.json
python3 dp/build/build.py treemap-dp/build/template.html tmdp.json treemap-dp/index.html
```

---

# Fall Budget Requests, Program by Program (mockup)

A mockup of a public report showing how decision packages (DPs) in the FY 2026-27 Fall Supplemental Budget would change each program's adopted budget.

**Live site:** https://brycelied.github.io/PortlandBudget/dp/

## What it shows

- **Drill-down:** citywide → service area → bureau → program offer. Each level shows adopted budget, proposed budget (adopted plus requests) and the change.
- **Programs:** each bureau lists every program offer in its adopted budget, grouped by program, with the adopted budget book's description and services.
- **Requests per program:** opening a program shows a ledger from the adopted budget through each request that changes it to the proposed budget, so programs changed by several DPs are easy to read.
- **Decision package detail:** selecting a request opens its description, service impacts, equity statement, every program it changes, resources and requirements by fund, and positions.
- **Filters:** service impact / no service impact, impact type tags, all funds / General Fund / other funds, and search.

## Mock data

Service impact flags and impact type tags are **not yet collected** in BFM. They are sample values assigned in `dp/build/prep.py` from each package's request type, amount and service impact text. Dollar amounts, FTE, package text and program descriptions are real.

## Sources

- **Requests:** BFM Fall TAO – D7 – DP Data Export (financial and personnel exports): 3,736 line items, 271 packages. The financial export leaves out personnel costs generated by positions, so the $7,417,520 in position costs from the personnel export is added to Personnel requirements, assigned to programs by each position's functional area.
- **Adopted budget:** Ad Hoc Reporting Tool, FY 2026-27 Adopted, by bureau and sub-program (program offer). Totals tie to $8,546,060,062 all funds and $1,186,002,037 General Fund.
- **Program descriptions:** FY 2026-27 Adopted Budget book on OpenBook (portland.openbook.questica.com), saved in `dp/build/program_descriptions.json`; 272 of 338 program offers matched by bureau and name.

## Rebuild

```
python3 dp/build/prep.py <DP export.xlsx> dp/build/bureau_sa.csv dp.json
python3 dp/build/prep_programs.py <DP export.xlsx> <Ad Hoc.xlsx> dp/build/bureau_sa.csv dp.json dp/build/program_descriptions.json prog.json
python3 dp/build/build.py dp/build/template.html prog.json dp/index.html
```

When bureaus start entering real tags, replace the sample tagging block in `prep.py` with the new BFM fields.
