# Copilot Instructions

## Overview

This repository contains VBScript (`.vbs`) files that automate Microsoft Office applications (Excel, Word, Outlook) from the Windows command line using Windows Script Host — no VBA editor required.

## Running Scripts

```cmd
cscript Excel\<filename>.vbs
```

Scripts must be run via `cscript` (not `wscript`) to see console output. There are no build, lint, or test tools.

## Architecture

- Scripts are organized by Office application (`Excel/`, `Word/`, `Outlook/`)
- Each `.vbs` file is standalone and self-contained — no shared libraries or imports
- Output files (`.xlsx`, etc.) are saved to the user's Desktop via `WScript.Shell.SpecialFolders("Desktop")`
- Excel/Office runs headlessly (`objExcel.Visible = False`, `DisplayAlerts = False`)

## Key Conventions

**File structure pattern (follow this in all scripts):**
1. File header comment block with description and run instructions
2. `Option Explicit` (always required)
3. Constants section (`' ── 常數設定 ─────`)
4. Sample data arrays
5. Main logic
6. Object cleanup (`Set obj = Nothing` for every COM object)

**VBScript-specific patterns:**
- Excel/Office numeric constants (e.g. `xlClusteredColumn = 51`) must be declared as local `Const` — VBScript cannot reference the Excel type library enum names directly
- COM objects are created with `CreateObject("Excel.Application")` and always released at the end
- Use `With ... End With` blocks for setting multiple properties on the same object

**Comments and naming:**
- Comments and string literals are in Traditional Chinese (繁體中文)
- Section separators use the pattern: `' ── 區段名稱 ────────────────────────────────────────────────`
- Variable names use camelCase with `obj` prefix for COM objects (e.g. `objExcel`, `objWorkbook`, `objSheet`)
- Array names use `arr` prefix (e.g. `arrMonths`, `arrSales`)
