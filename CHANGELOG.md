# Changelog

All notable changes to SlidePrep are documented in this file.

## [1.20260401.2] — 2026-04-01

### Changed
- **Mode 8 (ConvertToPDF):** The Purview label ID, display name, and justification message are now configurable via `-PurviewLabelId`, `-PurviewLabelName`, and `-PurviewJustification` parameters. Defaults remain unchanged (label "Public", justification "Customer Workshop delivery").
- **Mode 8 (ConvertToPDF):** On ARM64 systems, a warning is now displayed that Purview labeling may be significantly slower due to x86 emulation.

## [1.20260401.1] — 2026-04-01

### Added
- **Mode 8 (ConvertToPDF):** After exporting PDFs, the script now sets the Microsoft Purview Information Protection label **"Public"** (`87867195-f2b8-4ac2-b0b6-6bb73cb33afc`) on every exported PDF using `Set-FileLabel` with justification "Customer Workshop delivery".
- Requires the `PurviewInformationProtection` PowerShell module (≥ 3.2.57.0). If the module is not installed, labeling is skipped and the user is directed to the [Microsoft Purview Information Protection client download page](https://www.microsoft.com/en-us/download/details.aspx?id=53018).
- On **ARM64** systems (where the Purview module is unsupported), the labeling step automatically delegates to the x86 PowerShell host (`C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe`).

## [1.20260216.1] — 2026-02-16

### Fixed
- **Mode 6/7 (SetVariables / AddLogo):** Fixed hang when processing certain PPTX files. Root causes: missing `DisplayAlerts` suppression allowed PowerPoint to show blocking modal dialogs; default `WithWindow = True` created visible windows that could stall on UI prompts; accessing `TextFrame` on OLE/embedded shapes without a `HasTextFrame` guard could trigger a blocking OLE server call.

### Added
- **Mode 6/7:** Variable replacement now also covers table cells, matching the discovery scope of Mode 5 (DiscoverVariables).

### Changed
- All modes now set `DisplayAlerts = ppAlertsNone` and open presentations with `WithWindow = 0` to prevent COM automation hangs across the board.

## [1.20260213.2] — 2026-02-13

### Changed
- **Mode numbering:** Swapped Mode 4 and Mode 8 for a more logical workflow order. RemoveFinal is now Mode 4 (first step for presenters), ConvertToPDF is now Mode 8 (last step).

## [1.20260213.1] — 2026-02-13

### Added
- **Mode 5 (DiscoverVariables):** New `-VariablePrefix` and `-VariableSuffix` parameters allow configurable placeholder delimiters (default: `<<` and `>>`). Use `-VariablePrefix '{{' -VariableSuffix '}}'` to scan for `{{Variable}}` patterns instead.

## [1.20260211.2] — 2026-02-11

### Added
- **Mode 5 (DiscoverVariables):** Scan all PPTX files in a folder and list every unique `<<Variable>>` placeholder found. Outputs a sorted, deduplicated list and a ready-to-use hashtable hint for Mode 6.

### Changed
- Renamed project prefix from CCX to MSFT-CSU.

## [1.20260211.1] — 2026-02-11

### Changed
- Modernized codebase: strict typing on all variables, improved COM cleanup with dedicated `Invoke-ComCleanup` and `Remove-ComObject` helper functions.
- Enhanced comment-based help with full synopsis, description, parameter docs, and usage examples for all 8 modes.
- PSScriptAnalyzer-compliant verb-noun function naming throughout.
- Consistent error handling with retry logic (`Invoke-WithRetry`, up to 3 attempts).
- Replaced `Get-WmiObject` with `Get-CimInstance` for forward compatibility.
- Eliminated global-scope variables where possible; moved to `$script:` scope.

## [1.20260202.1] — 2026-02-02

### Fixed
- Assembly not loaded error when Office Interop assemblies are missing from the GAC.

## [1.20231124.1] — 2023-11-24

### Added
- Initial release with Modes 1–4 and 6–8: CleanPPTX, MarkFinal, SetLanguage, ConvertToPDF, SetVariables, AddLogo, RemoveFinal.
