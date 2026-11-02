# IT Separation Evidence Pack

**Synthetic, simulated worked example.** Project Meridian is a simulated divestiture of a specialty-components division to a simulated PE-backed buyer, with a synthetic enterprise value of approximately GBP 400 million. This is not a client deliverable and is not a record of any real engagement.

## Source data and scope

The core workbook loads its records from `data/apps.csv`, `data/tsa.csv`, `data/day1.csv`, `data/synergies.csv`, and `data/risks.csv`. The CSVs are left unchanged. Synthetic assumptions requested for TSA pricing, standalone costs, synergy tracking, and exit timing are blue inputs on `README_Assumptions`. Formula cells are black. Summary totals and reconciliations are formula-driven.

`PMI_100Day` is the Add-on A buyer integration plan, built from `pmi_milestones.csv`, `risks.csv`, `synergies.csv`, and `outsourcing.csv`. It retains the source rows, links identified synergies to `Cost_Synergy`, and treats outsourcing choices as alternatives that are excluded from the cost bridge and synergy total. `DD_Redflags` retains the supplied Project Alder source rows used by the executive deck.

## Workbook tabs

- **README_Assumptions** - scenario, synthetic model inputs, thresholds, color key, and calculation conventions.
- **App_Disposition** - 60 application records, formula-suggested and final disposition, editable validated override with rationale, TSA flag, input wave, complexity score, disposition counts/costs, and Day 1 critical app plan check. Rule precedence is shared with seller → TSA-then-replace; else transferable Y and Low entanglement → Transfer; else Consent → Clone; else High entanglement → Replicate; otherwise Replace. Final disposition uses a nonblank override before the suggested disposition. APP010 (criticality 2, suggested Clone) is the single synthetic Retire override; its rationale is “Synthetic example: function absorbed by buyer platform.”
- **TSA_Schedule** - 35 TSA services with monthly fee, inclusive calendar months, line amount, extension amount, buyer-side totals, a 24-month buyer/reverse run-off, and parsed app-link tokens. Buyer TSA cost excludes Buyer-to-seller lines; reverse lines remain separately visible.
- **Exit_Plan** - buyer-side TSA services only, with formula dates and milestone statuses, dependency-based float and RAG, design-date clamp flag, and source risks/mitigations. App-ID dependencies use the 30-day planned-end-to-cutover float convention; service-ID dependencies use the referenced service's cutover.
- **Day1_Checklist** - source checklist, RAG, readiness by workstream, and open No-go counts.
- **Cost_Synergy** - annual allocated/standalone cost by tower, input one-off cost by wave, identified synergies, 24-month planned and synthetic actual tracker, and two consistent-basis bridges. Bridge A is annual: allocated baseline + dis-synergy = standalone run-rate, less identified annual midpoint synergies = target run-rate. Bridge B shows Day 1, Wave 2 and Wave 3, then one-off separation cost (sum of waves) and one-time cost to separate (one-off + buyer TSA fees). TSA fees replace standalone run cost for covered services during the TSA period; they are not added to the run-rate.
- **Checks** - core reconciliations, a date-based TSA SUMPRODUCT independent of the Months, Monthly Fee and Total Planned columns, independent ties for each bridge, override rationale completeness, plus Add-on A checks for milestone dependency IDs, synergy links, risk linked IDs, and exclusion of outsourcing alternatives.
- **Dashboard** - headline metrics for the annual target run-rate and one-time separation cost, disposition and TSA run-off views, and distinct annual and one-time bridge tables/charts.
- **PMI_100Day** - governance principles, 30 milestone schedule with dependency RAG and resolution checks, linked synergy levers, five alternative outsourcing options, ten linked risks, and a formula-driven scorecard. Source CSV rows are retained below the presentation tables.
- **DD_Redflags** - supplied Alder domain ratings, red flags, and management questions used by the executive deck.

## Calculation conventions

- Day 1 is 2027-04-01; synergy month 1 is April 2027. TSA fee is cost basis plus 5%; extension months use a further 20% step-up on the monthly fee.
- Applications standalone annual cost is the source annual run-cost sum plus 8%. Other tower baselines and standalone estimates use the blue synthetic inputs.
- Synergy monthly plan uses the annual low/high midpoint divided by 12, ramped linearly from the source start month over the source ramp duration, then flat. Synthetic actuals use the supplied 85% / 92% realization assumptions through month 15, with the 85% period boundary held as a blue input; later months are blank. RAG thresholds are shown on `README_Assumptions`.
- Exit milestone dates derive from each buyer-side TSA planned end date and the blue lead/lag assumptions. The as-of date is 2027-07-09. Design dates earlier than Day 1 are set to Day 1 and flagged.
- `linked_app_ids` from `synergies.csv` and `linked_id` from `risks.csv` are retained and checked. TSA dependencies may resolve to an app ID or TSA service ID; only service-ID dependencies enter the TSA end-before-dependency-cutover check.
- SYN13 shares apps with SYN10 and SYN12 by design as process savings; this is not treated as double counting.
- PMI milestone RAG follows a directly linked TSA service's exit RAG; app and milestone dependencies are green when their IDs resolve. The CSVs do not provide completion states, so the scorecard's on-track percentage is a dependency-health proxy, not a delivery-status claim. The synergy comparison target is the sum of the supplied annual midpoint list because no separate target is present in the CSVs.

## Limitations

All figures and milestones are synthetic. Disposition and complexity rules are illustrative. One illustrative override does not change the suggested disposition or wave input. The source CSVs do not provide standalone tower inputs, TSA pricing assumptions, synergy actuals, or exit milestone dates; the synthetic assumptions supply those model drivers. No synergy is described as achieved. Workbook formula calculation and PDF values are independently checked by `build_check.py`; FINDINGS.md records a worked hand-trace of one TSA service and one application.

An override changes disposition and counts; the one-off cost estimate is an owner input and must be re-estimated separately.
