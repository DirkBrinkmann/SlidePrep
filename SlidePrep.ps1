#Requires -Version 5.1

<#
.SYNOPSIS
    Automates common PowerPoint (PPTX) batch operations for workshop creators and presenters.

.DESCRIPTION
    SlidePrep processes all PPTX files in a folder, supporting eight operation
    modes designed for two audiences:

    WORKSHOP CREATORS (content authors and v-Teams):
      Mode 1 - CleanPPTX        : Sanitize decks by stripping comments, notes, and metadata.
                                   Optionally removes hidden slides. Outputs to a separate folder.
      Mode 2 - MarkFinal         : Set the "Final" document property on every deck.
      Mode 3 - SetLanguage       : Set the proofing language on every text shape, table cell,
                                   and notes shape (default: English US) so spell-check works.

    WORKSHOP PRESENTERS (CSAs / delivery):
      Mode 4 - RemoveFinal       : Remove the "Final" flag so decks can be edited/customized.
      Mode 5 - DiscoverVariables : Scan all decks and list every unique <<Variable>> placeholder
                                   found. Run this BEFORE Mode 6 to know which variables to fill.
      Mode 6 - SetVariables      : Search-and-replace placeholder variables (e.g. <<Presenter>>)
                                   across all slides using a hashtable of key/value pairs.
      Mode 7 - AddLogo           : Insert a customer logo image onto every title slide.
                                   Position and scaling should be verified manually afterwards.
      Mode 8 - ConvertToPDF      : Export every deck to PDF for customer handout.

    The script uses PowerPoint COM Automation and must run on a Windows machine with
    Microsoft PowerPoint installed. All COM objects are released in a finally block to
    minimize orphaned POWERPNT.EXE processes.

    TYPICAL WORKFLOWS
    -----------------
    Presenter preparing for a customer engagement:
      1. Run Mode 4 (RemoveFinal)         - unlock the decks
      2. Run Mode 5 (DiscoverVariables)   - see which placeholders need values
      3. Run Mode 6 (SetVariables)        - fill in customer-specific values
      4. Run Mode 7 (AddLogo)             - brand the title slides
      5. Run Mode 8 (ConvertToPDF)        - create handout PDFs

    Creator finalizing a new release:
      1. Run Mode 3 (SetLanguage)    - normalize proofing language
      2. Perform manual spell-check
      3. Run Mode 1 (CleanPPTX)      - strip notes and metadata
      4. Run Mode 2 (MarkFinal)      - lock the decks

    TROUBLESHOOTING
    ---------------
    This script drives PowerPoint via COM, which can occasionally produce transient errors
    ("Can't find Object X"). The built-in retry logic (up to 3 attempts) handles most cases.
    If errors persist:
      1. Close all PowerPoint instances before running the script.
      2. Start a fresh PowerShell session.
      3. As a last resort, reboot and retry.

.PARAMETER CleanPPTX
    Activates Mode 1. Copies PPTX files from SourceFolder to DestinationFolder, then
    removes comments, notes, and document metadata. Requires -SourceFolder and
    -DestinationFolder.

.PARAMETER RemoveHiddenSlidesFromCleanedPPT
    Optional switch for Mode 1. When specified, hidden slides are deleted from the
    cleaned copies.

.PARAMETER MarkFinal
    Activates Mode 2. Sets the "Final" document property on every PPTX in SourceFolder.

.PARAMETER SetLanguage
    Activates Mode 3. Sets the proofing language on every text shape in every PPTX.

.PARAMETER MSOLanguageID
    The MSO language identifier to apply in Mode 3. Defaults to 'msoLanguageIDEnglishUS'.
    Must be a valid member of [Microsoft.Office.Core.MsoLanguageID].

.PARAMETER ConvertToPDF
    Activates Mode 8. Exports every PPTX in SourceFolder to PDF in DestinationFolder.
    Requires -SourceFolder and -DestinationFolder.

.PARAMETER RemoveFinal
    Activates Mode 4. Removes the "Final" document property from every PPTX in
    SourceFolder so the files can be edited.

.PARAMETER DiscoverVariables
    Activates Mode 5. Scans all PPTX files in SourceFolder and outputs a sorted, distinct
    list of every placeholder variable found (text matching the pattern <<VariableName>>).
    Run this mode before Mode 6 (SetVariables) to identify which variables need values.

.PARAMETER VariablePrefix
    The opening delimiter used to identify placeholder variables in Mode 5.
    Defaults to '<<'. Change this if your templates use a different prefix (e.g. '{{', '%%').

.PARAMETER VariableSuffix
    The closing delimiter used to identify placeholder variables in Mode 5.
    Defaults to '>>'. Change this if your templates use a different suffix (e.g. '}}', '%%').

.PARAMETER SetVariables
    Activates Mode 6. Performs search-and-replace of placeholder variables in all slides.
    Requires -SlideVariables.

.PARAMETER SlideVariables
    A hashtable mapping placeholder strings to replacement values.
    Example: @{ '<<Presenter>>' = 'Jane Doe'; '<<Date>>' = '2026-03-15' }

.PARAMETER AddLogo
    Activates Mode 7. Prompts for a logo file (JPG/PNG) and inserts it on slide 1 of
    each deck. Size and position should be adjusted manually afterwards.

.PARAMETER SourceFolder
    Path to the folder containing the PPTX files to process. Required for all modes.

.PARAMETER DestinationFolder
    Path to the output folder. Required for Mode 1 (CleanPPTX) and Mode 8 (ConvertToPDF).

.INPUTS
    None. This script does not accept pipeline input.

.OUTPUTS
    Modified PPTX files (in-place or in DestinationFolder) and/or PDF exports.
    Mode 5 (DiscoverVariables) outputs a sorted, distinct list of <<Variable>> placeholders
    to the console.
    A timestamped CSV log of all actions is saved in the source folder.

.EXAMPLE
    .\SlidePrep.ps1 -CleanPPTX -SourceFolder C:\Decks -DestinationFolder C:\Decks\Clean

    Mode 1: Copies all PPTX files from C:\Decks to C:\Decks\Clean and strips notes,
    comments, and metadata.

.EXAMPLE
    .\SlidePrep.ps1 -CleanPPTX -RemoveHiddenSlidesFromCleanedPPT -SourceFolder C:\Decks -DestinationFolder C:\Decks\Clean

    Mode 1 with hidden-slide removal: Same as above, but also deletes any hidden slides.

.EXAMPLE
    .\SlidePrep.ps1 -MarkFinal -SourceFolder C:\Decks

    Mode 2: Marks every PPTX in C:\Decks as "Final".

.EXAMPLE
    .\SlidePrep.ps1 -SetLanguage -SourceFolder C:\Decks

    Mode 3: Sets every text shape in every PPTX to English (US) proofing language.

.EXAMPLE
    .\SlidePrep.ps1 -SetLanguage -SourceFolder C:\Decks -MSOLanguageID msoLanguageIDGerman

    Mode 3: Sets proofing language to German instead of the default.

.EXAMPLE
    .\SlidePrep.ps1 -RemoveFinal -SourceFolder C:\Decks

    Mode 4: Removes the "Final" flag from all PPTX files so they can be edited.

.EXAMPLE
    .\SlidePrep.ps1 -DiscoverVariables -SourceFolder C:\Decks

    Mode 5: Scans all PPTX files and lists every unique <<Variable>> placeholder found.
    Example output:
      Variables found across 12 PPTX files:
        <<Company>>
        <<Date>>
        <<Presenter>>

.EXAMPLE
    .\SlidePrep.ps1 -DiscoverVariables -SourceFolder C:\Decks -VariablePrefix '{{' -VariableSuffix '}}'

    Mode 5 with custom delimiters: Scans for {{Variable}} placeholders instead of <<Variable>>.

.EXAMPLE
    $vars = @{ '<<Presenter>>' = 'Jane Doe'; '<<Company>>' = 'Contoso Ltd.' }
    .\SlidePrep.ps1 -SetVariables -SourceFolder C:\Decks -SlideVariables $vars

    Mode 6: Replaces all occurrences of <<Presenter>> and <<Company>> across every slide.

.EXAMPLE
    .\SlidePrep.ps1 -SetVariables -SourceFolder C:\Decks -SlideVariables @{'<<Date>>'='2026-03-15'}

    Mode 6: Inline hashtable variant — replaces <<Date>> in all decks.

.EXAMPLE
    .\SlidePrep.ps1 -AddLogo -SourceFolder C:\Decks

    Mode 7: Prompts for a logo file and inserts it on each title slide. Review positioning
    manually afterwards.

.EXAMPLE
    .\SlidePrep.ps1 -ConvertToPDF -SourceFolder C:\Decks -DestinationFolder C:\Decks\PDF

    Mode 8: Exports all PPTX files as PDF to C:\Decks\PDF.

.NOTES
    File Name : SlidePrep.ps1
    Author    : Dirk Brinkmann (dirk.brinkmann@microsoft.com)
    Requires  : Windows with Microsoft PowerPoint installed, PowerShell 5.1+

    Version History:
      1.20260213.2  2026-02-13  Swapped Mode 4/8: RemoveFinal is now Mode 4,
                                 ConvertToPDF is now Mode 8 for logical workflow order.
      1.20260213.1  2026-02-13  Mode 5: Added configurable -VariablePrefix and
                                 -VariableSuffix parameters (default << and >>).
      1.20260211.2  2026-02-11  Added Mode 5 (DiscoverVariables) to scan decks for
                                <<Variable>> placeholders. Renamed CCX -> MSFT-CSU.
      1.20260211.1  2026-02-11  Modernized: strict typing, improved COM cleanup, enhanced
                                help, PSScriptAnalyzer-compliant verb-noun naming, consistent
                                error handling with retry logic, replaced Get-WmiObject with
                                Get-CimInstance, eliminated global-scope variables where possible.
      1.20260202.1  2026-02-02  BUGFIX: Assembly not loaded
      1.20231124.1  2023-11-24  NEW: First release

    DISCLAIMER:
    This sample script is not supported under any Microsoft standard support program or
    service. It is provided AS IS without warranty of any kind. Microsoft further disclaims
    all implied warranties including, without limitation, any implied warranties of
    merchantability or of fitness for a particular purpose. The entire risk arising out of
    the use or performance of this sample script remains with you.

.LINK
    https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation
#>

# ============================================================================
# PARAMETERS
# ============================================================================
[CmdletBinding(DefaultParameterSetName = 'Mode1CleanPPTX')]
param (
    # --- Mode switches ---
    [Parameter(Mandatory, ParameterSetName = 'Mode1CleanPPTX',
        HelpMessage = 'Clean PPTX files: strip notes, comments, and metadata.')]
    [switch]$CleanPPTX,

    [Parameter(ParameterSetName = 'Mode1CleanPPTX',
        HelpMessage = 'Also remove hidden slides during the clean operation.')]
    [switch]$RemoveHiddenSlidesFromCleanedPPT,

    [Parameter(Mandatory, ParameterSetName = 'Mode2MarkFinal',
        HelpMessage = 'Set the Final document property on all PPTX files.')]
    [switch]$MarkFinal,

    [Parameter(Mandatory, ParameterSetName = 'Mode3SetLanguage',
        HelpMessage = 'Set the proofing language on all text shapes.')]
    [switch]$SetLanguage,

    [Parameter(ParameterSetName = 'Mode3SetLanguage',
        HelpMessage = 'MSO language ID (e.g. msoLanguageIDEnglishUS, msoLanguageIDGerman).')]
    [string]$MSOLanguageID = 'msoLanguageIDEnglishUS',

    [Parameter(Mandatory, ParameterSetName = 'Mode4RemoveFinal',
        HelpMessage = 'Remove the Final document property from all PPTX files.')]
    [switch]$RemoveFinal,

    [Parameter(Mandatory, ParameterSetName = 'Mode5DiscoverVariables',
        HelpMessage = 'Scan all PPTX files and list every <<Variable>> placeholder found.')]
    [switch]$DiscoverVariables,

    [Parameter(ParameterSetName = 'Mode5DiscoverVariables',
        HelpMessage = 'Opening delimiter for placeholders. Defaults to <<.')]
    [string]$VariablePrefix = '<<',

    [Parameter(ParameterSetName = 'Mode5DiscoverVariables',
        HelpMessage = 'Closing delimiter for placeholders. Defaults to >>.')]
    [string]$VariableSuffix = '>>',

    [Parameter(Mandatory, ParameterSetName = 'Mode6SetVariables',
        HelpMessage = 'Replace placeholder variables in all slides.')]
    [switch]$SetVariables,

    [Parameter(Mandatory, ParameterSetName = 'Mode6SetVariables',
        HelpMessage = "Hashtable of placeholders to replace. Example: @{'<<Var>>'='Value'}")]
    [hashtable]$SlideVariables,

    [Parameter(Mandatory, ParameterSetName = 'Mode7AddLogo',
        HelpMessage = 'Insert a customer logo on every title slide.')]
    [switch]$AddLogo,

    [Parameter(Mandatory, ParameterSetName = 'Mode8ConvertToPDF',
        HelpMessage = 'Convert all PPTX files to PDF.')]
    [switch]$ConvertToPDF,

    # --- Shared path parameters ---
    [Parameter(Mandatory, ParameterSetName = 'Mode1CleanPPTX',
        HelpMessage = 'Folder containing the PPTX files to process.')]
    [Parameter(Mandatory, ParameterSetName = 'Mode2MarkFinal')]
    [Parameter(Mandatory, ParameterSetName = 'Mode3SetLanguage')]
    [Parameter(Mandatory, ParameterSetName = 'Mode4RemoveFinal')]
    [Parameter(Mandatory, ParameterSetName = 'Mode5DiscoverVariables')]
    [Parameter(Mandatory, ParameterSetName = 'Mode6SetVariables')]
    [Parameter(Mandatory, ParameterSetName = 'Mode7AddLogo')]
    [Parameter(Mandatory, ParameterSetName = 'Mode8ConvertToPDF')]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$SourceFolder,

    [Parameter(Mandatory, ParameterSetName = 'Mode1CleanPPTX',
        HelpMessage = 'Destination folder for cleaned or exported files.')]
    [Parameter(Mandatory, ParameterSetName = 'Mode8ConvertToPDF')]
    [string]$DestinationFolder
)

# ============================================================================
# SCRIPT-SCOPED VARIABLES
# ============================================================================
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

[string]$script:ScriptVersion   = '1.20260213.2'
[string]$script:ScriptName      = $MyInvocation.MyCommand.Name
[datetime]$script:ScriptStart   = Get-Date
[string]$script:LogFileName     = '{0}-{1:yyyy-MM-dd-HH-mm}.csv' -f
    ([IO.FileInfo]$MyInvocation.MyCommand.Definition).BaseName, (Get-Date)
[string]$script:SessionGUID     = [Guid]::NewGuid().ToString()
[string]$script:PPTXFilter      = '*.pptx'
[int]$script:StepCount          = 1
[int]$script:ConsoleWidth       = 120
[string]$script:Delimiter       = '#' * $script:ConsoleWidth
[System.Collections.Generic.List[PSCustomObject]]$script:LogRecords = @()

# Load Office Interop assemblies (required for language ID enum)
try {
    Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint -ErrorAction SilentlyContinue
    Add-Type -AssemblyName Office -ErrorAction SilentlyContinue
}
catch {
    # Assemblies may already be loaded or available via COM; non-fatal
    Write-Verbose "Office Interop assemblies not found in GAC; COM automation will still work."
}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Write-LogAndHost {
    <#
    .SYNOPSIS
        Writes a message to the console and appends it to the in-memory log.
    .DESCRIPTION
        Combines Write-Host output with structured log record creation in a
        single call to reduce code duplication throughout the script.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('Information', 'Warning', 'Error')]
        [string]$Status = 'Information',

        [System.ConsoleColor]$ForegroundColor = 'White',

        [switch]$NoConsole
    )

    if (-not $NoConsole) {
        Write-Host $Message -ForegroundColor $ForegroundColor
    }

    $osVersion = try { (Get-CimInstance Win32_OperatingSystem).Version } catch { 'Unknown' }

    $record = [PSCustomObject]@{
        DateTimeUTC      = (Get-Date).ToUniversalTime()
        LogRecordVersion = 2
        SessionGUID      = $script:SessionGUID
        ScriptVersion    = $script:ScriptVersion
        OSVersion        = $osVersion
        User             = [Security.Principal.WindowsIdentity]::GetCurrent().Name
        PSVersion        = $PSVersionTable.PSVersion.ToString()
        PSCulture        = (Get-Culture).Name
        StepNumber       = $script:StepCount
        Status           = $Status
        Message          = $Message
    }
    $script:LogRecords.Add($record)
}

function Save-LogRecords {
    <#
    .SYNOPSIS
        Exports the accumulated log records to a semicolon-delimited CSV file.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Path
    )

    $fullPath = Join-Path -Path $Path -ChildPath $script:LogFileName
    $script:LogRecords | Export-Csv -Path $fullPath -Delimiter ';' -NoTypeInformation -Force -ErrorAction SilentlyContinue
    Write-Host ("Log saved: {0}" -f $fullPath) -ForegroundColor Yellow
}

function Set-ConsoleWidth {
    <#
    .SYNOPSIS
        Attempts to widen the console buffer and window to the specified width.
    #>
    [CmdletBinding()]
    param (
        [int]$Width = 120
    )

    try {
        $rawUI   = $Host.UI.RawUI
        $bufSize = $rawUI.BufferSize
        $winSize = $rawUI.WindowSize

        if ($bufSize.Width -lt $Width) {
            $bufSize.Width = $Width
            $rawUI.BufferSize = $bufSize
        }
        if ($winSize.Width -lt $Width) {
            $winSize.Width = $Width
            $rawUI.WindowSize = $winSize
        }
    }
    catch {
        Write-Verbose "Could not resize console window: $_"
    }
}

function Test-FolderReady {
    <#
    .SYNOPSIS
        Validates that a source folder exists or prompts to create/recreate a destination folder.
    .OUTPUTS
        [bool] $true if the folder is ready for use; $false otherwise.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory)]
        [string]$Folder,

        [switch]$IsSource
    )

    if (Test-Path -LiteralPath $Folder) {
        if ($IsSource) { return $true }

        # Destination already exists — ask the user
        Write-Host ("Folder [{0}] already exists." -f $Folder)
        $answer = Read-Host "`nDelete and recreate this folder? (y/n)"
        if ($answer -eq 'y') {
            Remove-Item -LiteralPath $Folder -Recurse -Force
            New-Item -Path $Folder -ItemType Directory | Out-Null
            Write-LogAndHost -Message ("Recreated folder: {0}" -f $Folder) -ForegroundColor Green
            return $true
        }
        return $false
    }
    else {
        if ($IsSource) {
            Write-LogAndHost -Message ("Source folder not found: {0}" -f $Folder) -Status Error -ForegroundColor Red
            return $false
        }
        # Destination doesn't exist — offer to create
        $answer = Read-Host ("Destination folder [{0}] does not exist. Create it? (y/n)" -f $Folder)
        if ($answer -eq 'y') {
            New-Item -Path $Folder -ItemType Directory | Out-Null
            Write-LogAndHost -Message ("Created folder: {0}" -f $Folder) -ForegroundColor Green
            return $true
        }
        return $false
    }
}

function Get-FileNameDialog {
    <#
    .SYNOPSIS
        Displays a Windows file-open dialog and returns the selected file path(s).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory)]
        [string]$WindowTitle,

        [Parameter(Mandatory)]
        [string]$InitialDirectory,

        [string]$Filter = 'All files (*.*)|*.*',

        [switch]$AllowMultiSelect
    )

    Add-Type -AssemblyName System.Windows.Forms
    $dialog = [System.Windows.Forms.OpenFileDialog]@{
        Title            = $WindowTitle
        InitialDirectory = $InitialDirectory
        Filter           = $Filter
        Multiselect      = [bool]$AllowMultiSelect
        ShowHelp         = $true   # Prevents hang in certain console hosts
    }
    $null = $dialog.ShowDialog()

    if ($AllowMultiSelect) { return $dialog.FileNames }
    else { return $dialog.FileName }
}

function Get-PictureInfo {
    <#
    .SYNOPSIS
        Returns a formatted string with image dimensions and resolution.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory)]
        [string]$PicturePath
    )

    Add-Type -AssemblyName System.Drawing
    $bitmap = [System.Drawing.Bitmap]::new($PicturePath)
    try {
        return ('Image size: {0} x {1} px | Resolution: {2} x {3} DPI' -f
            $bitmap.Width, $bitmap.Height,
            $bitmap.HorizontalResolution, $bitmap.VerticalResolution)
    }
    finally {
        $bitmap.Dispose()
    }
}

function Request-CustomerLogo {
    <#
    .SYNOPSIS
        Prompts the user to select a customer logo file (JPG or PNG).
    .OUTPUTS
        [string] Full path to the selected logo, or $null if cancelled.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param ()

    $answer = Read-Host 'Insert a customer logo on each title slide? (y/n)'
    if ($answer -ne 'y') { return $null }

    $logoPath = Get-FileNameDialog -WindowTitle 'Select a customer logo (JPG or PNG)' `
        -InitialDirectory 'C:\' -Filter 'Image files (*.jpg;*.png)|*.jpg;*.png'

    if ($logoPath -and (Test-Path -LiteralPath $logoPath)) {
        Write-LogAndHost -Message ("Selected logo: {0}" -f $logoPath) -ForegroundColor Green
        return $logoPath
    }

    Write-LogAndHost -Message ("Logo file not found: {0}" -f $logoPath) -Status Warning -ForegroundColor Red
    return $null
}

# ============================================================================
# COM LIFECYCLE HELPERS
# ============================================================================

function Remove-ComObject {
    <#
    .SYNOPSIS
        Safely releases a COM object and suppresses errors.
    .DESCRIPTION
        Calls ReleaseComObject to decrement the COM reference count and
        allows the .NET garbage collector to finalize the release. This
        helps prevent orphaned POWERPNT.EXE processes.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $ComObject
    )

    try {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
    catch {
        Write-Verbose "COM release warning: $_"
    }
}

function Invoke-ComCleanup {
    <#
    .SYNOPSIS
        Quits a PowerPoint COM application, releases all references, and forces GC.
    .DESCRIPTION
        Should be called in a finally{} block to guarantee COM cleanup regardless
        of whether the operation succeeded or threw an exception.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $Application
    )

    try { $Application.Quit() } catch { Write-Verbose "Application.Quit warning: $_" }
    Remove-ComObject -ComObject $Application
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ============================================================================
# ENVIRONMENT CHECK
# ============================================================================

function Test-Environment {
    <#
    .SYNOPSIS
        Checks for running PowerPoint processes and offers to terminate them.
    .DESCRIPTION
        COM automation can conflict with existing PowerPoint instances. This
        function detects running POWERPNT processes and gives the user the
        option to force-close them before proceeding.
    #>
    [CmdletBinding()]
    param ()

    Write-LogAndHost -Message "`nPre-flight checks:" -ForegroundColor Yellow

    $pptProcesses = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue)
    if ($pptProcesses.Count -gt 0) {
        $answer = Read-Host ('Found {0} running PowerPoint process(es). Force-close them to continue? (y/n)' -f $pptProcesses.Count)
        if ($answer -eq 'y') {
            $pptProcesses | Stop-Process -Force
            Write-LogAndHost -Message 'PowerPoint processes terminated.' -ForegroundColor Yellow
        }
        else {
            Write-LogAndHost -Message 'User chose not to close PowerPoint. Exiting.' -Status Warning -ForegroundColor Yellow
            exit 0
        }
    }
    else {
        Write-LogAndHost -Message "`t-> No running PowerPoint instances found." -ForegroundColor Green
    }
}

# ============================================================================
# CORE MODE FUNCTIONS
# ============================================================================

function Clear-Presentation {
    <#
    .SYNOPSIS
        Mode 1 worker: Strips metadata, comments, notes text, and optionally hidden slides.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Folder,

        [bool]$RemoveHiddenSlides = $false
    )

    $ppRDIAll = 99   # ppRemoveDocInfoType.ppRDIAll
    $application = New-Object -ComObject PowerPoint.Application

    try {
        if ($RemoveHiddenSlides) {
            Write-LogAndHost -Message 'Hidden slides will be removed during cleanup.' -ForegroundColor Cyan
        }

        foreach ($file in Get-ChildItem -Path $Folder -Filter $script:PPTXFilter) {
            $presentation = $application.Presentations.Open($file.FullName)
            try {
                # Remove final flag if present
                if ($presentation.Final) {
                    $presentation.Final = $false
                    Write-LogAndHost -Message ("`t-> Removed Final flag: {0}" -f $file.Name) -ForegroundColor Magenta
                }

                # Strip all document metadata
                $presentation.RemoveDocumentInformation($ppRDIAll) | Out-Null

                # Process each slide
                foreach ($slide in $presentation.Slides) {
                    if ($RemoveHiddenSlides -and $slide.SlideShowTransition.Hidden) {
                        $slideIndex = $slide.SlideIndex
                        $slide.Delete()
                        Write-LogAndHost -Message ("`t-> Deleted hidden slide #{0}" -f $slideIndex) -NoConsole
                        continue
                    }

                    # Clear notes text
                    if ($slide.HasNotesPage) {
                        foreach ($shape in $slide.NotesPage.Shapes) {
                            if ($shape.PlaceholderFormat.Type -eq 2 -and
                                $shape.HasTextFrame -eq -1 -and
                                $shape.TextFrame.HasText -eq -1) {
                                $shape.TextFrame.TextRange.Text = ''
                            }
                        }
                    }
                }

                $presentation.Save()
                Write-LogAndHost -Message ("`t-> Cleaned: {0}" -f $file.Name) -ForegroundColor Green
            }
            finally {
                $presentation.Close()
                Remove-ComObject -ComObject $presentation
            }
        }
    }
    finally {
        Invoke-ComCleanup -Application $application
    }
}

function Set-PresentationFinal {
    <#
    .SYNOPSIS
        Mode 2: Sets the "Final" document property on every PPTX in the folder.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    Write-LogAndHost -Message ("Step [{0}] - Marking PPTX files as Final in: {1}" -f $script:StepCount, $FolderPath) -ForegroundColor Yellow
    $script:StepCount++

    $application = New-Object -ComObject PowerPoint.Application
    try {
        foreach ($file in Get-ChildItem -Path $FolderPath -Filter $script:PPTXFilter) {
            Write-LogAndHost -Message ("Processing: {0}" -f $file.Name) -ForegroundColor Green
            $presentation = $application.Presentations.Open($file.FullName)
            try {
                if ($presentation.Final) {
                    Write-LogAndHost -Message ("`t-> Already marked as Final.") -ForegroundColor Green
                }
                else {
                    $presentation.Final = $true
                    $presentation.Save() | Out-Null
                    Write-LogAndHost -Message ("`t-> Set to Final.") -ForegroundColor Green
                }
            }
            finally {
                $presentation.Close() | Out-Null
                Remove-ComObject -ComObject $presentation
            }
        }
    }
    finally {
        Invoke-ComCleanup -Application $application
    }
}

function Remove-PresentationFinal {
    <#
    .SYNOPSIS
        Mode 4: Removes the "Final" document property from every PPTX in the folder.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    Write-LogAndHost -Message ("Step [{0}] - Removing Final flag from PPTX files in: {1}" -f $script:StepCount, $FolderPath) -ForegroundColor Yellow
    $script:StepCount++

    $application = New-Object -ComObject PowerPoint.Application
    try {
        foreach ($file in Get-ChildItem -Path $FolderPath -Filter $script:PPTXFilter) {
            Write-LogAndHost -Message ("Processing: {0}" -f $file.Name) -ForegroundColor Green
            $presentation = $application.Presentations.Open($file.FullName)
            try {
                if ($presentation.Final) {
                    $presentation.Final = $false
                    $presentation.Save() | Out-Null
                    Write-LogAndHost -Message ("`t-> Final flag removed.") -ForegroundColor Green
                }
                else {
                    Write-LogAndHost -Message ("`t-> Not marked as Final; no change needed.") -ForegroundColor Green
                }
            }
            finally {
                $presentation.Close() | Out-Null
                Remove-ComObject -ComObject $presentation
            }
        }
    }
    finally {
        Invoke-ComCleanup -Application $application
    }
}

function Set-PresentationLanguage {
    <#
    .SYNOPSIS
        Mode 3: Sets the proofing language on every text shape, table cell, and notes shape.
    .DESCRIPTION
        Iterates through all slides and their shapes. For shapes with text frames the
        LanguageID property is set to the specified MSO value. Table cells and notes
        shapes are also updated. Presentations marked as Final are skipped with a warning.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$FolderPath,

        [Parameter(Mandatory)]
        [string]$LanguageID
    )

    Write-LogAndHost -Message ("Step [{0}] - Setting language to [{1}] in: {2}" -f $script:StepCount, $LanguageID, $FolderPath) -ForegroundColor Yellow
    $script:StepCount++

    # Resolve the language enum value
    $msoLangValue = $null
    try {
        $msoLangValue = [Microsoft.Office.Core.MsoLanguageID]::$LanguageID
    }
    catch {
        Write-LogAndHost -Message ("Invalid MSOLanguageID: '{0}'. Check available values at: https://learn.microsoft.com/en-us/office/vba/api/office.msolanguageid" -f $LanguageID) -Status Error -ForegroundColor Red
        return
    }

    $application = New-Object -ComObject PowerPoint.Application
    try {
        foreach ($file in Get-ChildItem -Path $FolderPath -Filter $script:PPTXFilter) {
            Write-LogAndHost -Message ("Processing: {0}" -f $file.Name) -ForegroundColor Green
            $presentation = $application.Presentations.Open($file.FullName)
            try {
                if ($presentation.Final) {
                    Write-LogAndHost -Message "`t-> Skipped (marked as Final). Remove the Final flag first." -Status Warning -ForegroundColor Yellow
                    continue
                }

                $replaceCount = 0
                $slideCount   = 0

                foreach ($slide in $presentation.Slides) {
                    $slideCount++

                    # Process regular shapes
                    foreach ($shape in $slide.Shapes) {
                        # Text frame language
                        try {
                            if ($shape.TextFrame.TextRange.LanguageID -ne $msoLangValue) {
                                $shape.TextFrame.TextRange.LanguageID = $msoLangValue
                                $replaceCount++
                            }
                        }
                        catch { <# Shape may not have a text frame #> }

                        # Table cells
                        if ($shape.HasTable) {
                            for ($r = 1; $r -le $shape.Table.Rows.Count; $r++) {
                                foreach ($cell in $shape.Table.Rows($r).Cells) {
                                    try {
                                        $cell.Shape.TextFrame.TextRange.LanguageID = $msoLangValue
                                        $replaceCount++
                                    }
                                    catch { <# Cell may lack text #> }
                                }
                            }
                        }
                    }

                    # Process notes shapes
                    foreach ($notesShape in $slide.NotesPage.Shapes) {
                        try {
                            if ($notesShape.TextFrame.TextRange.LanguageID -ne $msoLangValue) {
                                $notesShape.TextFrame.TextRange.LanguageID = $msoLangValue
                                $replaceCount++
                            }
                        }
                        catch { <# Notes shape may not have text #> }
                    }
                }

                Write-LogAndHost -Message ("`t-> Updated {0} shape(s) across {1} slide(s)." -f $replaceCount, $slideCount) -ForegroundColor Cyan
                $presentation.Save() | Out-Null
            }
            finally {
                $presentation.Close() | Out-Null
                Remove-ComObject -ComObject $presentation
            }
        }
    }
    finally {
        Invoke-ComCleanup -Application $application
    }
}

function Export-PresentationToPdf {
    <#
    .SYNOPSIS
        Mode 8 worker: Exports a single presentation to PDF.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $Presentation,

        [Parameter(Mandatory)]
        [string]$PptxFileName,

        [Parameter(Mandatory)]
        [string]$OutputFolder
    )

    # PDF export constants
    $ppFixedFormatTypePDF           = 2
    $ppFixedFormatIntentPrint       = 2
    $ppPrintHandoutHorizontalFirst  = 2
    $ppPrintOutputSlides            = 1
    $ppPrintAll                     = 1
    $ppShowAll                      = 1
    $msoTrue                        = -1
    $msoFalse                       = 0

    if ($Presentation.Final) {
        Write-LogAndHost -Message ("`t-> Skipped (marked as Final): {0}" -f $PptxFileName) -Status Warning -ForegroundColor Magenta
        return
    }

    $outputFile = Join-Path -Path $OutputFolder -ChildPath ([IO.Path]::GetFileName($PptxFileName))
    $outputFile = [IO.Path]::ChangeExtension($outputFile, '.pdf')

    $printOptions = $Presentation.PrintOptions
    $range = $printOptions.Ranges.Add(1, $Presentation.Slides.Count)
    $printOptions.RangeType = $ppShowAll

    $Presentation.ExportAsFixedFormat(
        $outputFile, $ppFixedFormatTypePDF, $ppFixedFormatIntentPrint,
        $msoTrue, $ppPrintHandoutHorizontalFirst, $ppPrintOutputSlides,
        $msoFalse, $range, $ppPrintAll, 'Slideshow Name',
        $false, $false, $false, $false, $false
    )

    Write-LogAndHost -Message ("`t-> Exported PDF: {0}" -f $outputFile) -ForegroundColor Green
}

function Convert-PresentationsToPdf {
    <#
    .SYNOPSIS
        Mode 8: Iterates all PPTX files and exports each to PDF with retry logic.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$SourcePath,

        [Parameter(Mandatory)]
        [string]$DestinationPath
    )

    Write-LogAndHost -Message ("Step [{0}] - Converting PPTX to PDF: {1} -> {2}" -f $script:StepCount, $SourcePath, $DestinationPath) -ForegroundColor Yellow
    $script:StepCount++

    # Ensure destination exists
    if (-not (Test-Path -LiteralPath $DestinationPath)) {
        New-Item -Path $DestinationPath -ItemType Directory | Out-Null
    }

    $application = New-Object -ComObject PowerPoint.Application
    try {
        foreach ($file in Get-ChildItem -Path $SourcePath -Filter $script:PPTXFilter) {
            Write-LogAndHost -Message ("`t-> Processing: {0}" -f $file.Name) -ForegroundColor Green
            $maxRetries = 3
            for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
                try {
                    $presentation = $application.Presentations.Open($file.FullName)
                    try {
                        Export-PresentationToPdf -Presentation $presentation -PptxFileName $file.FullName -OutputFolder $DestinationPath
                    }
                    finally {
                        $presentation.Close() | Out-Null
                        Remove-ComObject -ComObject $presentation
                    }
                    break  # Success — exit retry loop
                }
                catch {
                    Write-LogAndHost -Message ("`t-> Attempt {0}/{1} failed: {2}" -f $attempt, $maxRetries, $_.Exception.Message) -Status Error -ForegroundColor Red
                    if ($attempt -eq $maxRetries) {
                        Write-LogAndHost -Message ("`t-> Giving up on: {0}" -f $file.Name) -Status Error -ForegroundColor Red
                    }
                }
            }
        }
    }
    finally {
        Invoke-ComCleanup -Application $application
    }
}

function Find-PresentationVariables {
    <#
    .SYNOPSIS
        Mode 5: Scans all PPTX files and returns a sorted, distinct list of <<Variable>> placeholders.
    .DESCRIPTION
        Opens each PPTX file read-only, iterates every shape on every slide (including tables),
        and extracts all text fragments matching the pattern <<...>>. The results are
        deduplicated, sorted alphabetically, and displayed to the console. This helps
        presenters identify which variables they need to supply values for before running
        Mode 6 (SetVariables).
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$FolderPath,

        [string]$Prefix = '<<',

        [string]$Suffix = '>>'
    )

    Write-LogAndHost -Message ("Step [{0}] - Discovering {1}Variables{2} in PPTX files in: {3}" -f $script:StepCount, $Prefix, $Suffix, $FolderPath) -ForegroundColor Yellow
    $script:StepCount++

    # Regex pattern: match anything between prefix and suffix (non-greedy)
    $escapedPrefix   = [regex]::Escape($Prefix)
    $escapedSuffix   = [regex]::Escape($Suffix)
    $variablePattern = '{0}.+?{1}' -f $escapedPrefix, $escapedSuffix

    $allVariables  = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $fileCount     = 0

    $application = New-Object -ComObject PowerPoint.Application
    try {
        foreach ($file in Get-ChildItem -Path $FolderPath -Filter $script:PPTXFilter) {
            $fileCount++
            Write-LogAndHost -Message ("Processing: {0}" -f $file.Name) -ForegroundColor Green
            $presentation = $application.Presentations.Open($file.FullName, <# ReadOnly #> -1, <# Untitled #> 0, <# WithWindow #> 0)
            try {
                foreach ($slide in $presentation.Slides) {
                    foreach ($shape in $slide.Shapes) {
                        # Check regular text frames
                        try {
                            $text = $shape.TextFrame.TextRange.Text
                            if ($text) {
                                foreach ($match in [regex]::Matches($text, $variablePattern)) {
                                    [void]$allVariables.Add($match.Value)
                                }
                            }
                        }
                        catch { <# Shape without text frame — skip #> }

                        # Check table cells
                        try {
                            if ($shape.HasTable) {
                                for ($r = 1; $r -le $shape.Table.Rows.Count; $r++) {
                                    foreach ($cell in $shape.Table.Rows($r).Cells) {
                                        try {
                                            $cellText = $cell.Shape.TextFrame.TextRange.Text
                                            if ($cellText) {
                                                foreach ($match in [regex]::Matches($cellText, $variablePattern)) {
                                                    [void]$allVariables.Add($match.Value)
                                                }
                                            }
                                        }
                                        catch { <# Cell without text — skip #> }
                                    }
                                }
                            }
                        }
                        catch { <# No table — skip #> }
                    }
                }
            }
            finally {
                $presentation.Close() | Out-Null
                Remove-ComObject -ComObject $presentation
            }
        }
    }
    finally {
        Invoke-ComCleanup -Application $application
    }

    # --- Output results ---
    Write-Host ''
    if ($allVariables.Count -eq 0) {
        Write-LogAndHost -Message ("No {0}Variables{1} found in {2} PPTX file(s)." -f $Prefix, $Suffix, $fileCount) -ForegroundColor Yellow
    }
    else {
        $sorted = $allVariables | Sort-Object
        Write-LogAndHost -Message ("{0} unique variable(s) found across {1} PPTX file(s):" -f $allVariables.Count, $fileCount) -ForegroundColor Cyan
        Write-Host ''
        foreach ($var in $sorted) {
            Write-Host ("  {0}" -f $var) -ForegroundColor White
        }
        Write-Host ''
        Write-LogAndHost -Message 'TIP: To replace these variables, create a hashtable and run Mode 6 (SetVariables):' -ForegroundColor Yellow
        # Build a ready-to-use hashtable hint
        $hint = '$vars = @{ ' + (($sorted | ForEach-Object { "'{0}' = ''" -f $_ }) -join '; ') + ' }'
        Write-Host ("  {0}" -f $hint) -ForegroundColor Gray
        Write-Host ("  .\SlidePrep.ps1 -SetVariables -SourceFolder '{0}' -SlideVariables `$vars" -f $FolderPath) -ForegroundColor Gray
    }
}

function Set-PresentationVariablesAndLogo {
    <#
    .SYNOPSIS
        Mode 6/7 worker: Replaces placeholder variables and/or inserts a logo on slide 1.
    .DESCRIPTION
        Opens a single PPTX, optionally removes the Final flag, iterates all shapes to
        perform text replacement from the CustomerVariables hashtable, and optionally
        inserts a customer logo image on the first slide.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$PptxPath,

        [hashtable]$CustomerVariables,

        [switch]$InsertCustomerLogo,

        [string]$CustomerLogoPath
    )

    $application = New-Object -ComObject PowerPoint.Application
    try {
        $presentation = $application.Presentations.Open($PptxPath)
        try {
            $replaceCounter     = 0
            $needsSave          = $false

            Write-LogAndHost -Message ("`t-> Processing: {0}" -f (Split-Path $PptxPath -Leaf)) -ForegroundColor Green

            # Remove final flag if present
            if ($presentation.Final) {
                $presentation.Final = $false
                Write-LogAndHost -Message "`t   Removed Final flag to allow editing." -ForegroundColor Cyan
            }

            foreach ($slide in $presentation.Slides) {
                # --- Variable replacement in shapes ---
                if ($CustomerVariables -and $slide.Shapes.Count -gt 0) {
                    foreach ($shape in $slide.Shapes) {
                        foreach ($key in $CustomerVariables.Keys) {
                            try {
                                if ($shape.TextFrame.TextRange.Text -match [regex]::Escape($key)) {
                                    $shape.TextFrame.TextRange.Text = $shape.TextFrame.TextRange.Text -replace
                                        [regex]::Escape($key), $CustomerVariables[$key]
                                    $needsSave = $true
                                    $replaceCounter++
                                }
                            }
                            catch { <# Shape without text frame #> }
                        }
                    }
                }

                # --- Logo insertion on slide 1 ---
                if ($InsertCustomerLogo -and $slide.SlideIndex -eq 1 -and $CustomerLogoPath) {
                    $logoLeft   = 180
                    $logoTop    = 40
                    $maxWidth   = 220
                    $maxHeight  = 80

                    Write-LogAndHost -Message ("`t   Inserting logo at Left:{0} Top:{1}" -f $logoLeft, $logoTop) -ForegroundColor Magenta
                    $slide.Shapes.AddPicture(
                        $CustomerLogoPath, <# LinkToFile #> 0, <# SaveWithDocument #> -1,
                        [int]$logoLeft, [int]$logoTop, [int]$maxWidth, [int]$maxHeight
                    ) | Out-Null
                    $needsSave = $true
                }
            }

            if ($needsSave) { $presentation.Save() }
            Write-LogAndHost -Message ("`t   Replaced {0} variable occurrence(s)." -f $replaceCounter) -ForegroundColor Magenta
        }
        finally {
            $presentation.Close() | Out-Null
            Remove-ComObject -ComObject $presentation
        }
    }
    finally {
        Invoke-ComCleanup -Application $application
    }
}

function Invoke-WithRetry {
    <#
    .SYNOPSIS
        Executes a script block with automatic retry on failure (up to 3 attempts).
    .DESCRIPTION
        COM automation can produce transient errors. This wrapper retries the given
        operation up to MaxRetries times before giving up.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [scriptblock]$Action,

        [int]$MaxRetries = 3,

        [string]$OperationName = 'operation'
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            & $Action
            return
        }
        catch {
            Write-LogAndHost -Message ("Attempt {0}/{1} for {2} failed: {3}" -f $attempt, $MaxRetries, $OperationName, $_.Exception.Message) -Status Error -ForegroundColor Red
            if ($attempt -eq $MaxRetries) {
                Write-LogAndHost -Message ("Giving up on: {0}" -f $OperationName) -Status Error -ForegroundColor Red
            }
        }
    }
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

function Start-Main {
    <#
    .SYNOPSIS
        Entry point that dispatches to the correct mode based on the active parameter set.
    #>
    [CmdletBinding()]
    param ()

    Write-LogAndHost -Message 'Script started.' -NoConsole
    Set-ConsoleWidth -Width $script:ConsoleWidth

    Write-Host $script:Delimiter -ForegroundColor Yellow
    Write-Host ("SlidePrep v{0}" -f $script:ScriptVersion) -ForegroundColor Yellow
    Write-Host $script:Delimiter -ForegroundColor Yellow

    Test-Environment

    $logPath = $SourceFolder   # default log location

    switch ($script:ActiveParameterSet) {

        'Mode1CleanPPTX' {
            Write-Host $script:Delimiter -ForegroundColor Yellow
            Write-LogAndHost -Message 'Mode: Clean PPTX files' -ForegroundColor Yellow

            if ((Test-FolderReady -Folder $SourceFolder -IsSource) -and
                (Test-FolderReady -Folder $DestinationFolder)) {

                # Copy originals to destination, then clean
                Copy-Item -Path (Join-Path $SourceFolder $script:PPTXFilter) -Destination $DestinationFolder -Force
                Clear-Presentation -Folder $DestinationFolder -RemoveHiddenSlides ([bool]$RemoveHiddenSlidesFromCleanedPPT)
                $logPath = $DestinationFolder

                Write-Host ''
                Write-Host ('IMPORTANT: Only hand over files from [{0}]!' -f $DestinationFolder) -ForegroundColor Yellow -BackgroundColor Red
            }
        }

        'Mode2MarkFinal' {
            Write-Host $script:Delimiter -ForegroundColor Yellow
            Write-LogAndHost -Message 'Mode: Mark all PPTX as Final' -ForegroundColor Yellow

            if (Test-FolderReady -Folder $SourceFolder -IsSource) {
                Set-PresentationFinal -FolderPath $SourceFolder
            }
        }

        'Mode3SetLanguage' {
            Write-Host $script:Delimiter -ForegroundColor Yellow
            Write-LogAndHost -Message ("Mode: Set language to [{0}]" -f $MSOLanguageID) -ForegroundColor Yellow

            if (Test-FolderReady -Folder $SourceFolder -IsSource) {
                Set-PresentationLanguage -FolderPath $SourceFolder -LanguageID $MSOLanguageID
            }
        }

        'Mode4RemoveFinal' {
            Write-Host $script:Delimiter -ForegroundColor Yellow
            Write-LogAndHost -Message 'Mode: Remove Final flag from all PPTX files' -ForegroundColor Yellow

            if (Test-FolderReady -Folder $SourceFolder -IsSource) {
                Remove-PresentationFinal -FolderPath $SourceFolder
            }
        }

        'Mode5DiscoverVariables' {
            Write-Host $script:Delimiter -ForegroundColor Yellow
            Write-LogAndHost -Message ("Mode: Discover {0}Variables{1} in all PPTX files" -f $VariablePrefix, $VariableSuffix) -ForegroundColor Yellow

            if (Test-FolderReady -Folder $SourceFolder -IsSource) {
                Find-PresentationVariables -FolderPath $SourceFolder -Prefix $VariablePrefix -Suffix $VariableSuffix
            }
        }

        'Mode6SetVariables' {
            Write-Host $script:Delimiter -ForegroundColor Yellow
            Write-LogAndHost -Message 'Mode: Replace slide variables' -ForegroundColor Yellow

            if (Test-FolderReady -Folder $SourceFolder -IsSource) {
                foreach ($file in Get-ChildItem -Path $SourceFolder -Filter $script:PPTXFilter) {
                    Invoke-WithRetry -OperationName $file.Name -Action {
                        Set-PresentationVariablesAndLogo -PptxPath $file.FullName -CustomerVariables $SlideVariables
                    }
                }
            }
        }

        'Mode7AddLogo' {
            Write-Host $script:Delimiter -ForegroundColor Yellow
            Write-LogAndHost -Message 'Mode: Add customer logo to title slides' -ForegroundColor Yellow

            if (Test-FolderReady -Folder $SourceFolder -IsSource) {
                $logoFile = Request-CustomerLogo
                if ($logoFile) {
                    Write-LogAndHost -Message ("`t{0}" -f (Get-PictureInfo -PicturePath $logoFile)) -ForegroundColor Magenta

                    foreach ($file in Get-ChildItem -Path $SourceFolder -Filter $script:PPTXFilter) {
                        Invoke-WithRetry -OperationName $file.Name -Action {
                            Set-PresentationVariablesAndLogo -PptxPath $file.FullName `
                                -InsertCustomerLogo -CustomerLogoPath $logoFile `
                                -CustomerVariables $SlideVariables
                        }
                    }
                }
            }
        }

        'Mode8ConvertToPDF' {
            Write-Host $script:Delimiter -ForegroundColor Yellow
            Write-LogAndHost -Message 'Mode: Convert PPTX to PDF' -ForegroundColor Yellow

            if ((Test-FolderReady -Folder $SourceFolder -IsSource) -and
                (Test-FolderReady -Folder $DestinationFolder)) {
                Convert-PresentationsToPdf -SourcePath $SourceFolder -DestinationPath $DestinationFolder
                $logPath = $DestinationFolder
            }
        }

        default {
            Write-LogAndHost -Message 'No mode selected. Run Get-Help .\SlidePrep.ps1 -Full for usage instructions.' -Status Warning -ForegroundColor Magenta
        }
    }

    # --- Wrap up ---
    Write-Host ''
    Write-Host $script:Delimiter -ForegroundColor Yellow
    $runtime = (New-TimeSpan -Start $script:ScriptStart -End (Get-Date)).TotalSeconds
    Write-LogAndHost -Message ("Script finished. Runtime: {0:N1} seconds." -f $runtime) -ForegroundColor Yellow
    Save-LogRecords -Path $logPath
    Write-Host $script:Delimiter -ForegroundColor Yellow
}

# Capture the script-level parameter set name before calling Start-Main,
# because $PSCmdlet inside a function refers to that function's own CmdletBinding.
$script:ActiveParameterSet = $PSCmdlet.ParameterSetName

# Launch
Start-Main
