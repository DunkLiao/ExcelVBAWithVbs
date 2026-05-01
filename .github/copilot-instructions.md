# Copilot instructions for this repository

## Validation commands

This repository does not have a separate build, lint, or automated test suite. Use the script-driven Excel flow as the working validation path:

```powershell
cscript //nologo scripts\bootstrap-report-workbook.vbs
cscript //nologo scripts\run-sales-report.vbs sample-data\sales-input.csv
cscript //nologo scripts\export-vba-modules.vbs
```

The closest equivalent to running a single test is refreshing the sample report against the checked-in CSV input:

```powershell
cscript //nologo scripts\run-sales-report.vbs sample-data\sales-input.csv
```

## High-level architecture

The repository is organized around a split source/artifact model for Excel automation:

- `src\vba\` contains the editable VBA source exported from Excel. Treat these files as the primary source of truth for workbook logic.
- `workbooks\ReportAutomationTemplate.xlsm` is the generated macro-enabled workbook artifact that hosts the VBA code at runtime.
- `scripts\*.vbs` are orchestration entry points that use Excel COM automation to create the workbook, run macros, and export VBA modules back to the repository.
- `sample-data\sales-input.csv` is the reference input used by the example reporting flow.

The current end-to-end path is:

1. `scripts\bootstrap-report-workbook.vbs` creates a new `.xlsm` workbook, normalizes the workbook to `InputData` and `Report` sheets, and imports `src\vba\modules\ReportAutomation.bas`.
2. `scripts\run-sales-report.vbs` opens the workbook and calls `ReportAutomation.LoadCsvAndGenerateReport`.
3. `ReportAutomation.LoadCsvAndGenerateReport` copies CSV data into `InputData`, aggregates totals from the `Team` and `Amount` columns, and writes the summary into `Report`.
4. `scripts\export-vba-modules.vbs` exports standard/class/form modules from the workbook back into `src\vba\`.

## Key conventions

- Keep new workbook logic in exported source under `src\vba\` and treat the `.xlsm` file in `workbooks\` as a generated runtime artifact, not the primary editing surface.
- New VBS entry points should resolve `repoRoot` from `WScript.ScriptFullName` and default to `workbooks\ReportAutomationTemplate.xlsm` unless a path is explicitly passed in.
- `ReportAutomation.bas` assumes a two-sheet workflow: raw imported data in `InputData`, generated output in `Report`. Preserve that pattern unless you are intentionally introducing a new flow.
- The reference CSV schema is header-driven and currently expects `Team` and `Amount`. If you change the schema, update both the VBA aggregation logic and the sample/usage documentation together.
- Runtime failures are surfaced explicitly with `Err.Raise` in both VBA and VBS. Follow that pattern instead of silently swallowing Excel, file, or schema errors.
- Use `Option Explicit` in both VBA and VBS files and keep helper logic inside the same script/module unless there is a clear shared abstraction to extract.
- `scripts\export-vba-modules.vbs` intentionally skips workbook/sheet document modules. If you add event-driven code to `ThisWorkbook` or worksheet modules, update the export workflow so those modules are preserved in source control.
