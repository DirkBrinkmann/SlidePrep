# Changelog

All notable changes to SlidePrep are documented in this file.

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
