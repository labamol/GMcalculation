# Gross Margin Workflow — Consolidated SOP / Specification (Steps 1–3)

**Revision:** 2.0 | **Revenue refresh:** Weekly | **Headcount refresh:** Daily | **Margin:** On demand using selected inputs

# SOP — Quarterly Baseline Revenue Aggregation

**Version:** 2.0 (updated for Step 3 integration)
**Process:** Gross Margin Workflow — Step 1
**Status:** Approved working specification (subject to workbook-specific validation)
**Output:** Three aggregated worksheets in the original quarterly baseline workbook

## 1. Purpose and scope
Transform the detailed quarterly revenue baseline into three independent revenue views for downstream gross-margin processing. This step **does not calculate gross margin**.

## 2. Inputs
- Excel workbook: Quarterly Baseline Report (`.xlsx`).
- Source worksheet: `<Q> Baseline Data`, where `<Q>` is the reporting quarter (e.g., `Q3 Baseline Data`).
- Reporting quarter must be supplied or determined unambiguously from the workbook; if multiple candidate sheets exist, request confirmation.
- Revenue units: **INR Crores**. Preserve original numerical precision.

### 2.1 Exact source-column mapping
| Excel column | Source header / meaning | Role |
|---|---|---|
| B | Segment | Grouping dimension |
| C | LoB | Grouping dimension |
| D | Service Line | Grouping dimension |
| G | Q1 Rev Actual | Sum of Q1 actual revenue |
| J | Q2 Rev Actual | Sum of Q2 actual revenue |
| R | Q3 Target(Per AOP) Catch up | Sum of reporting-quarter AOP catch-up revenue target |
| U | Q3 Current Rev (MIH) - Total | Sum of reporting-quarter Money In Hand revenue |

The Q3 portion of quarter-dependent headers is illustrative and must match the selected reporting quarter. Confirm header identity as well as the physical column positions before processing. Column U may be displayed under a grouped/merged header; read the actual numeric values from U and verify its meaning. Do not interpret the adjacent grouped columns as additional measures.

## 3. Output worksheets
All outputs are written into the **same workbook** as the source. Each output is derived **independently** from the source, not from another output view.

| Worksheet | Grouping keys | Output fields |
|---|---|---|
| `Segment View` | Segment (B) | Segment; Q1 Rev Actual; Q2 Rev Actual; `<Q> Target(Per AOP) Catch up`; `<Q> Current Rev (MIH) - Total` |
| `Account View` | Segment (B) + LoB (C) | Segment; AccountOrLoB; the same four financial measures |
| `SP View` | Service Line (D) | Service Line; the same four financial measures |

`Account View` remains a Segment–LoB rollup using source `LoB` (column C), but its **output column header is `AccountOrLoB`**. Do not substitute an unprovided account-name or account-ID field.

## 4. Transformation rules
1. Read source rows below the actual header row, excluding headings, blank rows, and verified subtotal/grand-total rows to avoid double counting.
2. Validate the four numeric measures. Treat genuine blank numeric cells as zero only when confirmed to represent missing contribution; otherwise flag ambiguous blanks. Do not silently coerce text or Excel errors to zero.
3. Preserve all source records in reconciliation. Flag blank grouping keys and do not silently drop their revenue. Use an explicit `(Missing Segment)`, `(Missing LoB)`, or `(Missing Service Line)` exception bucket if necessary, and report affected row numbers.
4. Normalize grouping keys conservatively: trim incidental leading/trailing whitespace; do not merge differently spelled names, change case, or map aliases without an approved rule. Flag suspected duplicates for review.
5. For each unique grouping key or key tuple, independently SUM columns G, J, R, and U. Negative numbers remain negative. Never average revenue measures or sum across different reporting views.
6. Write the results to the three output worksheets with clear headers, INR Crores unit notation, numeric formatting, frozen headers, and filters where supported.
7. If a destination sheet already exists, refresh its generated data only; do not overwrite unrelated workbook sheets. If the destination contains non-generated content, stop and request approval rather than delete it.
8. Save a new version of the **same workbook** (preserving the original source workbook as a backup) and return its downloadable path. Do not claim that the original was overwritten unless an actual in-place save occurred.

## 5. Reconciliation and controls
For each of the four revenue measures, calculate:
- Source detail total (after excluding verified source subtotals).
- `Segment View` total.
- `Account View` total.
- `SP View` total.
- Difference of each view versus source total.

**Pass condition:** each difference is zero within a documented small floating-point tolerance (suggested absolute tolerance: 0.000001 INR Crores). If not, report the difference and mark validation FAILED; do not silently declare success.

Also check:
- Unique grouping keys per view (no duplicate output keys).
- Missing/invalid grouping values and nonnumeric/errored measures.
- Row count and source rows processed/excluded.
- Quarter-dependent headers match the reporting quarter.
- Formula cells in source are read from valid calculated values; if formula caches are missing or stale, calculate with a capable engine or stop and flag the limitation.

Provide a concise validation report in the response; an additional validation worksheet is optional only if explicitly requested. Do not modify the source data or its formulas.

## 6. Quarter handling
The source sheet is `<Q> Baseline Data`. Q1/Q2 historical actual headers remain as provided for the Q3 example; do **not** infer what prior-quarter actual columns should be for Q1, Q2, or Q4. For other quarters, validate the source layout and ask for revised mappings if the four financial fields do not exist at G/J/R/U or the meaning changes.

## 7. Open items / assumptions requiring workbook validation
- Exact header row, first data row, and whether the source contains subtotals.
- Whether U's grouped heading is represented by merged cells and which row carries the exact measure label.
- Treatment of blank numeric cells and source formulas without calculated cached values.
- Source workbook may have macros or external links: preservation must be verified; do not use a library that strips unsupported workbook features without warning.

## 8. Handoff to next stage
The three views are revenue baselines for subsequent gross-margin steps. No gross-margin numerator, cost allocation, or margin percentage is computed at this stage.


---

# SOP — Daily Headcount Processing (Step 2)

**Version:** 2.0 (cost unit clarified)
**Status:** Agreed functional logic; pending validation against an actual daily workbook
**Frequency:** Daily (Step 1 quarterly baseline refreshes weekly)
**Input:** Daily Headcount Workbook, source sheet `Data`
**Outputs in same headcount workbook:** `Key Data`, `Headcount Dashboard`, `Costs By Segment`, `Costs for MetLife`.

## 2.1 Key Data extraction

Find source fields by **header label, not Excel letter**. Match after trimming and collapsing whitespace; otherwise require exact label (case-insensitive matching permitted only when unambiguous). For the date-varying billability field, match the stable prefix `Billable / Non-billable - As on` and retain the full dated source label. If zero or multiple matches exist, stop and report. Preserve `Data` unchanged.

Extract all data records without filtering, deduplication or aggregation. Write `Key Data` with the following **21 columns in this exact order**:

| Output position | Source field label |
|---|---|
| A | Employee Number |
| B | Employee Grade |
| C | BPS/Non BPS |
| D | WON/SWON |
| E | Project Number |
| F | Cost |
| G | SP Type |
| H | Project Sub-SP |
| I | Group Customer |
| J | Customer |
| K | BFSI Account Name |
| L | Project Location wrt India |
| M | Project Type |
| N | LOB |
| O | Segment Wise |
| P | Billable / Non-billable - As on [Snapshot Date] (retain actual source heading) |
| Q | New Mapping |
| R | New Mapping Segment |
| S | NCI Ageing Days |
| T | Domain |
| U | Segment |

Preserve values, blanks, identifiers, dates, numerical precision, and record order. Confirm all 21 fields exist and `Key Data` record count equals `Data`. Report missing/ambiguous fields before writing. If an existing output contains user-maintained data, request confirmation before replacing it.

## 2.2 Headcount Dashboard

Read `Key Data` by field labels. Employee Number is expected to be **unique and nonblank**, one row per employee. Validate this; flag any violation instead of silently deduplicating. Location classification uses trimmed, case-insensitive `Project Location wrt India`: `ONSITE` or `OFFSITE`.

Produce five separate aggregation tables in one worksheet `Headcount Dashboard`:

| View | Grouping |
|---|---|
| Segment Summary | Segment Wise |
| Billability Summary | Segment Wise + dated Billable / Non-billable field |
| New Mapping Summary | Segment Wise + New Mapping |
| WON/SWON Summary | Segment Wise + WON/SWON |
| SP Type Summary | Segment Wise + SP Type |

For every group, output `Onsite HC` = count of nonblank Employee Number where location is ONSITE; `Offsite HC` = corresponding count where location is OFFSITE; `Total HC` = Onsite HC + Offsite HC. Add a grand total to each table. Retain blank grouping values as an explicit `(Blank)` category so totals reconcile; flag them for review. Flag unrecognized/blank locations and show exception counts separately; do not silently include them in Total HC. `Employee Grade` remains available in Key Data but is not a grouping key. All five views should reconcile to the count of employees with valid locations. Do not calculate FTE or allocation weighting.

## 2.3 Costs By Segment

**Read `Data` by header labels** (not positional references). Group by unique `Segment Wise` + `Customer`. For each group:

- `ONSITE COST` = sum of numeric `Cost` for location `ONSITE`.
- `OFFSITE COST` = sum of numeric `Cost` for location `OFFSITE`.
- `TOTAL COST` = ONSITE COST + OFFSITE COST.

Output columns in order: `Segment Wise`, `Customer`, `ONSITE COST`, `OFFSITE COST`, `TOTAL COST`. Sort by Segment Wise, then Customer. Insert a `SEGMENT TOTAL` row after each segment's customer detail rows, plus a `GRAND TOTAL` row at bottom. The Segment Wise column is the primary filter; Customer is secondary. Enable Excel AutoFilter on the table/header and freeze the header row. Segment subtotals and grand totals must respond to filters; do not double-count subtotals. For an Excel formula implementation, use `SUBTOTAL(109,...)` over customer-detail rows for segment subtotals and over the full detail-plus-subtotal range for grand totals: Excel SUBTOTAL ignores nested SUBTOTAL formulas. Validate this behavior after filtering. If the workbook engine cannot support correctly filter-aware subtotal rows, report the limitation rather than substituting static totals. Display numeric amounts without changing underlying precision.

## 2.4 Costs for MetLife

**Read `Data` by header labels.** First restrict to rows where `Customer` contains the substring `MetLife` (case-insensitive; e.g., MetLife US qualifies). Then group by unique `Segment Wise` + `LOB`. Compute `ONSITE COST`, `OFFSITE COST`, `TOTAL COST` with the same location rules as Step 2.3. Output columns: `Segment Wise`, `LOB`, `ONSITE COST`, `OFFSITE COST`, `TOTAL COST`. Sort by Segment Wise, then LOB; insert filter-aware `SEGMENT TOTAL` rows and a filter-aware `GRAND TOTAL`. Segment Wise is the primary filter, LOB the secondary filter. Apply the same formula, formatting, reconciliation, and no-double-counting requirements as Step 2.3.

## Cross-cutting checks and exceptions

1. Keep original `Data` untouched; do not alter Step 1 workbook or its three revenue views.
2. Treat source `Cost` values as-is; **Cost is monthly cost per employee, denominated in INR**. Step 2 aggregates monthly INR costs without conversion; Step 3 converts to INR Crores for comparison with quarterly revenue.
3. For costs, preserve blank grouping keys in explicit `(Blank)` groups and flag them. Non-numeric/blank Cost values must be reported by row and excluded from sums pending review; do not silently coerce them to zero. Classifications other than ONSITE/OFFSITE must be flagged and excluded from the two cost buckets.
4. Reconcile each dashboard table's valid-location total to the source valid-location employee count. Reconcile cost detail, segment subtotals, and grand total to eligible source Cost totals separately for each location. For MetLife, reconcile only against the filtered MetLife source subset.
5. Record daily source filename, snapshot date (from verified source metadata/header), counts of processed rows, missing values, duplicate employee IDs, invalid locations/costs, and PASS/FAIL checks. Never invent a snapshot date.
6. Preserve original workbook features where possible; warn if macros, external links, formulas, or unsupported features cannot be retained. Save as a new updated copy to protect the original. Existing generated sheets may be refreshed; do not overwrite unrelated manual content without approval.

## Open clarification

- Cost currency and periodicity are confirmed: monthly INR per employee. Confirm whether any cost adjustment beyond the specified ×3 projection is required only if later instructed.
- Confirm any additional business-specific exceptions or special handling for blank/unknown categories if the source workbook reveals them.


---

# SOP — Gross Margin Calculation (Step 3)

**Version:** 2.0
**Status:** Confirmed calculation logic; actual workbook mapping and financial reconciliation pending input files
**Input A:** Weekly Quarterly Baseline Workbook with `Segment View` and `Account View` from Step 1
**Input B:** Daily Headcount Workbook with `Costs By Segment` from Step 2
**Output:** New workbook `Margin Calculation.xlsx`

## 3.1 Create a margin worksheet for each Segment

1. Read distinct `Segment` values from the Quarterly Revenue Workbook's `Segment View`.
2. Create one worksheet per Segment, named `Margin-[Segment]`. Excel sheet names are limited to 31 characters and prohibit `: \ / ? * [ ]`. Where needed, use deterministic safe shortening and a collision-proof suffix, and document the full original Segment-to-sheet mapping. Never merge two Segments because their shortened names collide.
3. For each Segment, select rows from the Quarterly Revenue Workbook's `Account View` where `Segment` equals the current Segment. Source field `LoB` is exposed in this view as **`AccountOrLoB`**. Populate output `Account` from `AccountOrLoB`.
4. Create one detail row per distinct Segment + AccountOrLoB; source `Account View` is already aggregated at this level. Preserve revenue precision. Flag missing/duplicate keys or unmatched Segments.

## 3.2 Required 12-column output layout

`<Q>` means reporting quarter (e.g., Q3). Revenue and calculated margin cost columns are in **INR Crores**; J/K/L are percentages.

| Excel column | Output header | Data source / formula | Unit |
|---|---|---|---|
| A | `Account` | `AccountOrLoB` in `Account View` for the current Segment | Text |
| B | `Q1 Actual` | `Q1 Rev Actual` in `Account View` | INR Crores |
| C | `Q2 Actual` | `Q2 Rev Actual` in `Account View` | INR Crores |
| D | `<Q> Target(Per AOP) Catch up` | Same reporting-quarter field in `Account View` | INR Crores |
| E | `<Q> Current Rev (MIH) - Total` | Same reporting-quarter field in `Account View` | INR Crores |
| F | `Monthly manpower Cost (MP)` | Matched `TOTAL COST` from `Costs By Segment` **divided by 10,000,000** | INR Crores/month |
| G | `<Q> MP` | F × 3 | INR Crores/quarter |
| H | `<Q> NMP` | Default numeric 0 (editable for future inputs) | INR Crores/quarter |
| I | `Total Cost` | G + H | INR Crores/quarter |
| J | `GM Projection versus <Q> Target` | (D − I) / D | Percentage |
| K | `GM Projection versus MIH` | (E − I) / E | Percentage |
| L | `ISU Target` | Default numeric 0 | Percentage |

**Example formulas (data row 2):** `G2=IF(ISNUMBER(F2),F2*3,"")`, `H2=0`, `I2=IF(ISNUMBER(G2),G2+H2,"")`, `J2=IF(AND(ISNUMBER(I2),ISNUMBER(D2),D2<>0),(D2-I2)/D2,"")`, `K2=IF(AND(ISNUMBER(I2),ISNUMBER(E2),E2<>0),(E2-I2)/E2,"")`, `L2=0`. Format J, K, L as `0.00%` (numeric 0 displays `0.00%`). Do not multiply stored percentage results by 100. For zero/blank D or E, flag undefined GM and display blank rather than 0%.

## 3.3 Monthly manpower cost mapping (Column F)

- Read `Costs By Segment` in the **Daily Headcount Workbook**, using **customer-level detail rows only**; exclude `SEGMENT TOTAL` and `GRAND TOTAL`.
- Match **both**: current Segment = `Segment Wise`; output Account (sourced from `AccountOrLoB`) = `Customer`.
- Match full values after trimming incidental whitespace; case-insensitive exact matching may be used consistently. **No fuzzy, substring, or cross-Segment matching**.
- On exactly one matching detail row, read `TOTAL COST`, a **monthly INR** value, and calculate F = `TOTAL COST / 10000000`.
- If zero or multiple matching rows exist, report an exception, leave F and dependent financial outputs unresolved (blank/error marker with explicit exception), and **never substitute zero**.
- Ensure a genuine zero-cost matching row is treated as a valid zero rather than an unmatched lookup.
- `Costs for MetLife` remains a separate Step 2 analysis; it is **not** the default source for F.

## 3.4 Financial units and period alignment

- Step 1 revenue measures are quarterly **INR Crores**.
- Step 2 source `Cost` is **monthly INR per employee**, and `Costs By Segment` totals remain **monthly INR**.
- Step 3 F converts monthly INR to monthly INR Crores using **1 crore = INR 10,000,000**.
- Step 3 G multiplies F by **3 months** to project quarterly manpower cost. This assumes the current monthly snapshot is representative of the quarter; it is a projection, not a historical quarter-to-date payroll sum.
- Step 3 H is quarterly non-manpower cost, default 0 INR Crores. Step 3 I = G + H.
- Do not convert revenue a second time, do not multiply source INR costs by 3 before converting and again afterward, and do not include Step 2 segment subtotal/grand total rows in F lookups.

## 3.5 Formatting, totals and validations

- Set readable column widths, a frozen header row, and filters for the Account detail table. Display B–I as numbers (e.g., `#,##0.00` while retaining precision); J–L as percentages (`0.00%`). Clearly label revenue/cost units INR Crores.
- Do not invent a Segment total row or aggregate GM percentages without a confirmed weighted-total rule. A reconciliation report can provide source totals separately.
- For each Segment, verify B–E sums against matching `Account View` and against `Segment View` totals for all four revenue measures (within a stated numerical tolerance).
- Verify F values against eligible `Costs By Segment` customer detail rows, and G/I arithmetic. Report unmatched/ambiguous costs, unmatched Segment/Account, zero revenue denominators, and invalid numeric values.
- Record exact input filenames, reporting quarter, headcount snapshot date if verifiable, and source-to-output mapping. Weekly and daily input refreshes are independent; never silently assume matching as-of dates.
- Create a new `Margin Calculation.xlsx` workbook, leaving both input workbooks unchanged. Preserve any existing user-maintained output content by requesting approval before replacing it. Return actual saved file(s) and validation status.

## 3.6 Remaining clarification / not yet specified

- Whether Step 3 should include per-Segment totals or any cross-Segment summary worksheet (not specified).
- How to treat differences in account/customer naming beyond exact normalized matching (do not infer mappings).
- Whether special account exceptions or nonzero NMP/ISU targets will be supplied later.
