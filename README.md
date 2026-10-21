# Gross margin calculation

Deterministic Python implementation of the [Steps 1–3 specification](docs/specification.md), with a Streamlit dashboard and downloadable Excel workbooks. No AI calls, fuzzy mappings, external services, or credentials are used in calculations.

## Run

Python 3.12 is tested.

```bash
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements-lock.txt
streamlit run app.py
```

For the prepared cloud environment, use `/workspace/gm-venv/bin/streamlit run app.py` from `/workspace/GMcalculation`.

```bash
python -m pytest -q
python examples/create_demo.py /tmp/gm-demo
```

The demo creates synthetic inputs only; no employee data is included in the repository. Upload both demo files in **All steps**. Expected Banking/Account A monthly cost = 2 crores, quarterly cost = 6 crores, target GM = 40%, MIH GM = 40%.

## Workflows and outputs

- **All steps**: upload the quarterly original and daily original to generate all three workbooks.
- **Step 1: Revenue**: refresh Segment View, Account View and SP View in a new copy of the quarterly workbook.
- **Step 2: Headcount**: extract the 21 fields and generate the five headcount tables and two cost analyses in a new daily workbook copy.
- **Step 3: Margin**: upload refreshed workbooks with Segment View, Account View and Costs By Segment. Original Data/baseline detail sheets are not required. This supports independent weekly and daily refreshes.

Every run exposes tables in the UI and downloadable Excel and JSON validation reports. Downloaded names are `Quarterly Baseline Updated.xlsx`, `Headcount Updated.xlsx`, and `Margin Calculation.xlsx`. Inputs stay untouched. Margin output has 12 columns per segment, editable zero NMP/ISU defaults, and a Run Information sheet containing filenames, dates, units, and deterministic segment-to-sheet mapping. There are no invented segment GM totals.

## Input controls

Select the reporting quarter explicitly. Select the last baseline header row (for grouped headings this is the bottom of the header band) and the Data header row. Revenue mapping requires B/C/D/G/J/R/U **and** the specified labels; other quarters retain Q1/Q2 actual labels and require the same validated financial meanings. The app does not guess alternative quarter mappings.

List verified subtotal/heading rows to exclude from baseline processing. Completely empty rows are ignored; other rows remain subject to validation. Suspected total labels stop processing rather than double counting. Explicitly confirm whether genuine blank revenue cells mean zero; text and Excel errors are never treated as zero. Source formulas require fresh Excel/LibreOffice recalculation, saved caches, and explicit user confirmation; freshness cannot be established from the workbook alone.

Headcount fields are matched by whitespace-normalized, unambiguous case-insensitive headers; grouping/margin keys are only trimmed and remain case-sensitive. Employee identifiers, source number formats, dates and record order are retained in Key Data. Snapshot dates are recorded only if unambiguous in the dated billability header (ISO, `21-Oct-2026`, or `21 Oct 2026` for Step 2; ISO for imported Step 3). Unverified dates remain unset, never invented.

Missing keys remain explicit exception buckets. Duplicate/blank employees, unknown locations and invalid costs appear in validation reports. Invalid costs/locations are excluded from eligible cost totals; **an affected Segment/Customer is unresolved for margin**, including partially valid groups. App-generated headcount files retain this exception metadata for independent Step 3. Third-party Costs By Segment workbooks require upstream validation: the app cannot infer excluded employee rows from a standalone aggregate. Duplicate imported cost matches remain unresolved. A genuinely matched zero cost is valid.

## Precision and reconciliation

Aggregation and calculations use `Decimal` from the source numeric values. Sorting, matching and sheet-name generation are deterministic; there is no randomness. Calculations retain precision internally; Excel numeric serialization uses its supported floating-point representation and display formats do not round stored business values. ZIP timestamps/workbook metadata mean exported files are not promised byte-for-byte identical.

Revenue reconciliation tolerance is 0.000001 INR Crores. Cost arithmetic checks use 0.000001 INR; headcount checks require exact integer equality. Each revenue view derives independently from source records. Monthly INR cost is divided by 10,000,000 once, then multiplied by three once. Missing costs leave F/G/I/J/K blank. Zero denominators leave the corresponding GM blank and generate an exception. Validation reports distinguish issue-free PASS from FAIL/review required; numeric reconciliation can pass while data exceptions remain.

Cost subtotals use `SUBTOTAL(109,…)`, with nested subtotals excluded from grand totals. They are recalculated by Excel/LibreOffice when opened. openpyxl cannot calculate formulas, so cost-total formula caches may be blank until opened in a spreadsheet engine. Margin formulas include independently calculated cached values for immediate display. The UI uses the same calculation results rather than reading formula caches. The app clearly reports that interactive filtering needs spreadsheet-engine validation; static totals are never substituted for filter-aware subtotal formulas.

## Workbook preservation and limits

Only plain `.xlsx` workbooks are accepted. Macro content, external links, drawings/charts/comments with VML, pivots, slicers, embedded objects and Excel extensions are rejected because safe preservation cannot be guaranteed. Source cells/formulas and unrelated sheets are retained semantically, but Excel package bytes and formula caches are rewritten by openpyxl. Recalculate updated workbooks before feeding cached formula cells to other software. Keep original uploads as backups. Test against representative real files before operational use.

Generated sheets have an application marker and a content digest. Unrecognized destination sheets or changes to previously generated cell content stop replacement. To preserve user changes, use the original source or rename the modified sheet in a separate copy. Formatting-only changes are not protected by the content digest. The app returns new files and never writes over an existing filesystem output.

This is an in-memory tool intended for trusted internal workbook uploads, not a hardened public multi-tenant upload service. Employee data remains in the app session memory and the downloaded files. Host it only in an appropriately controlled environment. Changing widgets does not recalculate existing results; click **Process selected inputs** again.

## Validation

`python -m pytest -q` covers aggregation, reconciliation, costs and unit conversions, zero/unmatched costs, missing keys, duplicates, invalid input, formula-cache rejection, refresh protection, sheet-name collisions, independent Step 3, and Streamlit rendering. Synthetic tests do not replace acceptance checks against actual quarterly/daily workbooks. See [validation evidence](docs/validation.md).
