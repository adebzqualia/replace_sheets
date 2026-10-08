# POPS Update

Refresh selected country worksheets in an existing final POPS workbook, then let Excel recalculate the existing intermediary tables and synthesis formulas.

**Author:** Adnane El Bouhali  
**Team:** Value Chain team — BNP Paribas Personal Finance

This is an independent tool for corrections received after the four-stage POPS process has produced its final workbook. It does not modify the four existing projects.

## When to use it

Use this tool when a country sends corrected values or formulas in the **same template layout**. Select one or more countries and, for each, the sheets to refresh.

```text
Existing final workbook + Corrected country workbooks + Update configuration
                                  ↓
                  Refresh selected country cell contents
                                  ↓
              Excel recalculates intermediary and synthesis formulas
                                  ↓
                     Separate verified output workbook
```

The destination worksheets keep their names, identities, order, formatting, tables and drawings. They are not deleted and recreated. Only their cell contents are refreshed at the corresponding coordinates, preserving the existing downstream references.

If KPI rows, period columns, FTE blocks or group definitions have changed, regenerate the affected stages through the original POPS process. This updater does not discover or rebuild those mappings.

## Requirements and installation

- Windows, Python 3.10 or later, and installed desktop Microsoft Excel for an update.
- The final `.xlsx` workbook containing country sheets, intermediary tables and synthesis sheets.
- Corrected country `.xlsx` workbooks with compatible layouts.
- Write access to the chosen output directory. Close any existing output before replacing it.

From PowerShell in this project directory:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py --check-excel
```

The explicit interpreter avoids needing to activate the environment. `--check-excel` checks desktop Excel availability; it does not validate reporting workbooks.

## Configure the update

Edit [`config/update.yaml`](config/update.yaml). The supplied paths and Germany/Belgium selections are examples to replace with the intended corrections.

**Relative workbook paths start at the YAML file's directory**, not the terminal directory. With the supplied file under `config/`, `../data/` refers to the project's `data/` directory.

Replace the relevant fields while retaining the supplied country prefixes and sheet mapping:

```yaml
final_workbook: ../data/Consolidated_with_PPT.xlsx
output_workbook: ../output/Consolidated_with_PPT_updated.xlsx
overwrite_output: false
check_text_layout: true
updates:
  - country: Germany
    file: ../data/updates/POPS2027-Germany.xlsx
    sheets: [ID Card, FTE]
  - country: Belgium
    file: ../data/updates/POPS2027-Belgium.xlsx
    sheets: [ID Card, FTE]
```

This fragment illustrates an update plan; the full file also requires `version`, `country_prefixes` and `sheet_mapping`.

| Setting | Meaning |
| --- | --- |
| `final_workbook` | Existing final workbook to use as the base. |
| `output_workbook` | Separate result path; it must differ from every input. |
| `overwrite_output` | `false` prevents replacement of an existing output; set `true` deliberately when replacement is intended. |
| `check_text_layout` | `true` checks the positions and content of constant text labels before updating. Keep it enabled for normal use. |
| `sheet_password` | Password supplied for protected sheets; configure it for the applicable template. |
| `country_prefixes` | Country selection name to existing consolidated prefix. |
| `sheet_mapping` | Country source-sheet name to consolidated sheet suffix. |
| `updates[].country` | Exact key from `country_prefixes`. The filename is not used to infer the country. |
| `updates[].file` | Corrected workbook for that country. |
| `updates[].sheets` | Source-sheet names to refresh. Omitting this field selects all entries in `sheet_mapping`, currently 13. |

For example, `country: Germany` and `sheets: [ID Card]` select source `ID Card` and destination `Germany_ID Card`. Both must exist. `Findomestic` and `Italy` use prefix `Italy`; `PFF` and `France` use prefix `PFF`. Do not select two aliases for the same destination country in one run.

The loader rejects unknown fields, duplicate YAML keys, duplicate destinations, repeated source files, invalid worksheet names and colliding input/output paths. Selecting a subset is intentional: other country worksheets remain as they were.

## Check and run

```powershell
# Validate the configuration structure without checking workbook files
.\.venv\Scripts\python.exe main.py --validate-config

# Inspect the files and workbook layout without starting Excel
.\.venv\Scripts\python.exe main.py --dry-run

# Perform the update and Excel recalculation
.\.venv\Scripts\python.exe main.py
```

The default configuration is `config/update.yaml`. Use `--config` to select another file, for example:

```powershell
.\.venv\Scripts\python.exe main.py --config 'C:/POPS/config/update.yaml' --dry-run
.\.venv\Scripts\python.exe main.py --config 'C:/POPS/config/update.yaml'
```

A dry run checks files and their XML structure; it cannot demonstrate successful native Excel transfer or calculation. Review the selected countries and sheets before running the update.

On a successful run, the configured Excel output is accompanied by `<output-stem>.update_report.json`, containing the selected country prefixes, output path and warnings.

## How the update works

1. Validate the plan and compare each selected source sheet with its existing destination.
2. Open the corrected source read-only. Excel adjusts mapped worksheet and table names in memory to match the destination names; the source is closed without saving.
3. Refresh values and formulas at the same coordinates inside the existing destination worksheets. Existing destination formatting and objects remain in place; changed source charts, drawings or formatting are not imported.
4. Preserve the existing intermediary and synthesis formulas and ask Excel to rebuild dependencies and recalculate. Python does not calculate the business results.
5. Save and verify a candidate workbook, then publish it to the separate configured output. The original final workbook and corrected sources remain unchanged.

Missing selected sheets, incompatible merged ranges, mismatched table areas or headers, unsupported array/spill formulas, unsupported or external imported formula references, and conflicting defined names stop the update. These restrictions prevent unsupported transfers from being silently accepted.

The default text-layout comparison helps detect moved or renamed KPI labels and period headings. It does not prove every cell has the same business meaning, particularly for numeric headings. Setting `check_text_layout: false` disables that comparison; it does not remap moved data or make a changed template compatible.

## Review the result and its limits

Open the output in Excel and check the corrected country values, representative intermediary totals and their linked synthesis cells. Recalculation updates formulas that already exist; it does not validate the business accuracy of the submission.

- `FILLING_RATE`, `VALIDATION` and `PPT_VALIDATION` remain snapshots from their earlier generation. Their values, statuses and recorded source values are not refreshed by this updater.
- A previously missing KPI or period may have left an intermediary cell without a source formula. Updating the country data does not create that missing link. Previously unmapped or unresolved synthesis targets also need the original generator to be rerun.
- Existing formulas still use their original coordinates, country groups, rounding and presentation rules. This is a correction workflow, not a new reporting-year configuration workflow.
- Excel recalculation is not a refresh of external workbooks or data connections. Imported formulas that require unsupported external references are rejected.
- The output remains an Excel workbook. No PowerPoint file is generated.

## Tests

Run the existing portable tests without Excel:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests
```

To include native Excel integration tests on a suitable Windows machine:

```powershell
$env:POPS_RUN_EXCEL_TESTS = '1'
.\.venv\Scripts\python.exe -m unittest discover -s tests
Remove-Item Env:POPS_RUN_EXCEL_TESTS
```

Native tests require desktop Excel and are opt-in. Portable checks do not certify a production workbook or replace an actual Excel run and output review.

## Excel operations

The implementation uses Excel's [ClearContents](https://learn.microsoft.com/en-us/office/vba/api/excel.range.clearcontents) to remove cell contents while retaining formatting, [Formula2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula2) to read and write formula content, and [CalculateFullRebuild](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatefullrebuild) to recalculate and rebuild formula dependencies. Formula2 support does not imply that this updater accepts array or spill formulas.
