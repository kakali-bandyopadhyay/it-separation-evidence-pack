# Findings - IT separation evidence pack

Simulated worked example, synthetic data. Not a client deliverable and not past engagement work.

## Verification
- Checks tab: 13 PASS / 0 FAIL. build_check.py recomputes every total independently and matches the workbook to GBP 0.
- Hand-trace TSA05: cost basis 4,500 x (1 + 5% mark-up) = 4,725 per month x 7 months (Apr-Oct 2027) = 33,075. Appears in the run-off at 4,725 for each month.
- Trap found: the Oct-Nov run-off drop is also 33,075, but only by coincidence (TSA05, TSA13, TSA18 and TSA27 all end 31 Oct). The correct check is the monthly total against the active lines.
- Hand-trace APP001 (Core ERP): shared with seller -> TSA-then-replace; TSA01 covers ERP to Mar 2028; one-off 1,250,000 sits in Wave 2.

## Changes made after review
- Cost bridge split in two. The first version added an annual run-rate, one-off cost, about 24 months of TSA fees and annual synergies into one number, and counted the TSA months twice. Now:
  - Bridge A, annual run-rate: baseline 13,227,000 + dis-synergy 1,924,160 = standalone 15,151,160 - synergies 1,482,500 = target 13,668,660.
  - Bridge B, one-time cost: one-off 7,839,000 (Day 1 3,235,000 / Wave 2 4,151,000 / Wave 3 453,000) + TSA fees 4,953,270 = 12,792,270.
- Override column now drives final disposition, counts and TSA flag; every override needs a rationale (checked). Illustrative override: APP010 Clone -> Retire.
- Two self-referential checks replaced with independent ties to source data.
- A build pass changed the disposition rules to fit the example override; reverted. Rules come first, examples must fit them.

## Reading the result
- Even after identified synergies, the standalone run-rate sits above the allocated baseline: dis-synergy outweighs synergy, so the buyer case cannot rest on IT savings (synthetic figures).

## Limitations
- Illustrative disposition rules, synthetic costs and dates; one-off cost is an owner input per app and is not re-estimated on override.
