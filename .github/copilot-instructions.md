# Copilot Instructions — SlidePrep

## Overview

This repository contains a single PowerShell script (`SlidePrep.ps1`) that batch-processes PowerPoint (PPTX) files via COM Automation. It targets two audiences: **workshop creators** (content authors) and **workshop presenters** (delivery/CSAs). There is no build system, test suite, or CI pipeline.

## Running the Script

```powershell
# Requires Windows + Microsoft PowerPoint installed + PowerShell 5.1+
.\SlidePrep.ps1 -<ModeName> -SourceFolder <path> [-DestinationFolder <path>]
```

Lint with PSScriptAnalyzer:
```powershell
Invoke-ScriptAnalyzer -Path .\SlidePrep.ps1
```

## Architecture

The script is a single-file, parameter-set-driven tool with 8 modes selected by switch parameters:

| Mode | Switch | Audience | Purpose |
|------|--------|----------|---------|
| 1 | `-CleanPPTX` | Creators | Strip notes, comments, metadata; optionally remove hidden slides |
| 2 | `-MarkFinal` | Creators | Set "Final" document property |
| 3 | `-SetLanguage` | Creators | Set proofing language on all text shapes (default: English US) |
| 4 | `-ConvertToPDF` | Presenters | Export PPTX → PDF |
| 5 | `-DiscoverVariables` | Presenters | List all `<<Variable>>` placeholders found in decks |
| 6 | `-SetVariables` | Presenters | Search-and-replace `<<Variable>>` placeholders via hashtable |
| 7 | `-AddLogo` | Presenters | Insert customer logo on title slides |
| 8 | `-RemoveFinal` | Presenters | Remove "Final" flag to enable editing |

**Execution flow:** `param()` → `$script:ActiveParameterSet` captured → `Start-Main` → `Test-Environment` (checks for running POWERPNT processes) → `switch` dispatches to the appropriate mode function → `Save-LogRecords` writes a timestamped CSV.

## Key Conventions

- **COM lifecycle:** Every mode function creates its own `PowerPoint.Application` COM object. Cleanup (`Invoke-ComCleanup`) always runs in a `finally` block. Individual presentations are also released via `Remove-ComObject`. Never skip COM cleanup—orphaned `POWERPNT.EXE` processes are a real problem.
- **Retry pattern:** Transient COM errors are expected. Use `Invoke-WithRetry` (3 attempts) when calling worker functions that open presentations. Mode 4 (PDF export) has its own inline retry loop.
- **Logging:** All significant actions go through `Write-LogAndHost`, which writes to console AND appends to an in-memory `$script:LogRecords` list. At script end, `Save-LogRecords` exports to a semicolon-delimited CSV in the source/destination folder.
- **Strict mode:** `Set-StrictMode -Version Latest` and `$ErrorActionPreference = 'Stop'` are set at script scope. All variables are strongly typed with `[type]` annotations.
- **Parameter sets:** Each mode is a named `ParameterSetName` (e.g., `Mode1CleanPPTX`). `-SourceFolder` is mandatory for all modes. `-DestinationFolder` is required only for Modes 1 and 4.
- **Version scheme:** `Major.YYYYMMDD.Revision` (e.g., `1.20260211.2`). Update `$script:ScriptVersion` on every change.
- **Variable placeholders** use `<<VariableName>>` syntax (double angle brackets). The regex pattern is `<<.+?>>` (non-greedy).
- **Final flag awareness:** Several modes check `$presentation.Final` and either skip the file or auto-remove the flag before editing. Keep this behavior consistent when adding new modes.
- **No `Get-WmiObject`:** Use `Get-CimInstance` instead (PSScriptAnalyzer compliance).
