# SlidePrep

Batch-process PowerPoint (PPTX) files for workshop preparation and delivery. Sanitize decks, replace placeholder variables, insert logos, convert to PDF, and more — all from the command line.

## Prerequisites

- **Windows** with **Microsoft PowerPoint** installed
- **PowerShell 5.1** or later
- *(Optional)* **[Microsoft Purview Information Protection client](https://www.microsoft.com/en-us/download/details.aspx?id=53018)** — required for automatic PDF labeling in Mode 8

## Before You Begin

Download all workshop PPTX files to a local folder on your machine (e.g., `C:\Decks`). The script processes files in-place or copies them to a destination folder depending on the mode — it does not work with files on SharePoint, OneDrive, or network shares directly.

## Quick Start

```powershell
# Clone or download the script, then run with the desired mode:
.\SlidePrep.ps1 -<Mode> -SourceFolder C:\MyDecks [-DestinationFolder C:\Output]
```

## Modes

### For Workshop Creators

| Mode | Switch | Description |
|------|--------|-------------|
| 1 | `-CleanPPTX` | Copy decks to a destination folder and strip comments, notes, and metadata. Optionally remove hidden slides with `-RemoveHiddenSlidesFromCleanedPPT`. |
| 2 | `-MarkFinal` | Set the "Final" document property on every deck to prevent accidental edits. |
| 3 | `-SetLanguage` | Set the proofing language on all text shapes, table cells, and notes (default: English US). |

### For Workshop Presenters

| Mode | Switch | Description |
|------|--------|-------------|
| 4 | `-RemoveFinal` | Remove the "Final" flag so decks can be edited and customized. |
| 5 | `-DiscoverVariables` | Scan all decks and list every unique `<<Variable>>` placeholder found. Supports `-VariablePrefix` / `-VariableSuffix` for custom delimiters. Run this before Mode 6. |
| 6 | `-SetVariables` | Search-and-replace placeholder variables across all slides using a hashtable. |
| 7 | `-AddLogo` | Insert a customer logo image (JPG/PNG) onto every title slide. |
| 8 | `-ConvertToPDF` | Export every deck to PDF for customer handout. Automatically sets the Purview Information Protection label to **Public** on each PDF (requires the Purview client; see [below](#purview-information-protection-labeling)). |

## Usage Examples

### Clean decks for handout

```powershell
.\SlidePrep.ps1 -CleanPPTX -SourceFolder C:\Decks -DestinationFolder C:\Decks\Clean
```

With hidden slide removal:

```powershell
.\SlidePrep.ps1 -CleanPPTX -RemoveHiddenSlidesFromCleanedPPT -SourceFolder C:\Decks -DestinationFolder C:\Decks\Clean
```

### Remove / restore Final flag

```powershell
.\SlidePrep.ps1 -RemoveFinal -SourceFolder C:\Decks
.\SlidePrep.ps1 -MarkFinal -SourceFolder C:\Decks
```

### Set proofing language

```powershell
# Default (English US)
.\SlidePrep.ps1 -SetLanguage -SourceFolder C:\Decks

# German
.\SlidePrep.ps1 -SetLanguage -SourceFolder C:\Decks -MSOLanguageID msoLanguageIDGerman
```

### Convert to PDF

```powershell
.\SlidePrep.ps1 -ConvertToPDF -SourceFolder C:\Decks -DestinationFolder C:\Decks\PDF
```

After exporting, the script automatically sets the Microsoft Purview Information Protection label **"Public"** on every PDF. If the Purview module is not installed the labeling step is skipped and PDFs are still exported normally. See [Purview Information Protection Labeling](#purview-information-protection-labeling) for details.

### Discover and replace variables

```powershell
# Step 1 — Find all placeholders (default << >> delimiters)
.\SlidePrep.ps1 -DiscoverVariables -SourceFolder C:\Decks

# Step 1 (alt) — Find placeholders with custom delimiters
.\SlidePrep.ps1 -DiscoverVariables -SourceFolder C:\Decks -VariablePrefix '{{' -VariableSuffix '}}'

# Step 2 — Replace them
$vars = @{ '<<Presenter>>' = 'Jane Doe'; '<<Company>>' = 'Contoso Ltd.'; '<<Date>>' = '2026-03-15' }
.\SlidePrep.ps1 -SetVariables -SourceFolder C:\Decks -SlideVariables $vars
```

### Add a customer logo

```powershell
.\SlidePrep.ps1 -AddLogo -SourceFolder C:\Decks
# A file dialog will prompt you to select a JPG or PNG logo.
# Position and scaling should be verified manually afterwards.
```

## Typical Workflows

### Presenter preparing for a customer engagement

1. **Remove Final** — unlock the decks
2. **Discover Variables** — see which placeholders need values
3. **Set Variables** — fill in customer-specific values
4. **Add Logo** — brand the title slides
5. **Convert to PDF** — create handout PDFs

### Creator finalizing a new release

1. **Set Language** — normalize proofing language
2. Perform manual spell-check
3. **Clean PPTX** — strip notes and metadata
4. **Mark Final** — lock the decks

## Purview Information Protection Labeling

When running **Mode 8 (ConvertToPDF)**, the script automatically sets the Microsoft Purview Information Protection label **"Public"** on every exported PDF using `Set-FileLabel` with justification *"Customer Workshop delivery"*.

### Requirements

- **PowerShell module:** `PurviewInformationProtection` version **3.2.57.0** or later.
- **Install the module** by downloading the [Microsoft Purview Information Protection client](https://www.microsoft.com/en-us/download/details.aspx?id=53018).

### Behavior

| Scenario | What happens |
|----------|-------------|
| Module installed | Each PDF is labeled **Public** after export. |
| Module not installed | A warning is displayed with the download link. PDF export still succeeds; only labeling is skipped. |
| ARM64 architecture | The labeling step is automatically delegated to the x86 PowerShell host (`C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe`) because the Purview module does not support ARM64 natively. |

## Logging

Every run produces a timestamped CSV log file (semicolon-delimited) in the source or destination folder with details of each action taken.

## Troubleshooting

This script drives PowerPoint via COM automation, which can occasionally produce transient errors. Built-in retry logic (up to 3 attempts) handles most cases. If errors persist:

1. Close all PowerPoint instances before running the script.
2. Start a fresh PowerShell session.
3. As a last resort, reboot and retry.

## License

This sample script is provided AS IS without warranty of any kind. See the script header for the full disclaimer.
