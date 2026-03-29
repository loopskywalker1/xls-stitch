#Requires -Version 5.1
<#
.SYNOPSIS
    Consolidates identically-named sheets from multiple monthly Excel workbooks
    into single per-sheet output files with a "Month Year" source column.

.DESCRIPTION
    Iterates over Excel files in a folder, extracts sheets by name, unions
    columns across all months, normalizes rows, and exports one consolidated
    .xlsx file per sheet name.

    Handles column mismatches across months — when a column exists in one
    month but not another, the missing values are filled with blanks.

.PARAMETER FolderPath
    Mandatory. Path to folder containing monthly Excel workbooks.

.PARAMETER OutputPath
    Optional. Directory for output files. Default: <FolderPath>\output

.PARAMETER SheetNames
    Optional. Specific sheet names to process. Default: auto-discover all.

.PARAMETER FilePattern
    Optional. Glob for input files. Default: *.xlsx

.EXAMPLE
    .\Stitch-ExcelSheets.ps1 -FolderPath "C:\Data\Monthly Reports"

.EXAMPLE
    .\Stitch-ExcelSheets.ps1 -FolderPath "C:\Data\Monthly Reports" -SheetNames "Sheet1","Sheet2"

.EXAMPLE
    .\Stitch-ExcelSheets.ps1 -FolderPath "C:\Data\Monthly Reports" -OutputPath "D:\Consolidated" -FilePattern "*.xlsx"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Path to the folder containing monthly Excel workbooks.")]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$FolderPath,

    [Parameter(Mandatory = $false, HelpMessage = "Directory to save consolidated output workbooks. Default: <FolderPath>\output")]
    [string]$OutputPath,

    [Parameter(Mandatory = $false, HelpMessage = "Array of specific sheet names to process. Omit to auto-discover all.")]
    [string[]]$SheetNames = @(),

    [Parameter(Mandatory = $false, HelpMessage = "Glob pattern for matching input files.")]
    [string]$FilePattern = "*.xlsx"
)

# ============================================================================
# Phase 0: Prerequisites & Setup
# ============================================================================

# Resolve OutputPath default
if (-not $OutputPath) {
    $OutputPath = Join-Path $FolderPath "output"
}

# Check and install ImportExcel module
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "The ImportExcel module is not installed." -ForegroundColor Yellow
    $response = Read-Host "Install it now? (Y/N)"
    if ($response -eq 'Y' -or $response -eq 'y') {
        Write-Host "Installing ImportExcel module..." -ForegroundColor Cyan
        try {
            Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
            Write-Host "ImportExcel module installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install ImportExcel module: $_"
            exit 1
        }
    }
    else {
        Write-Error "ImportExcel module is required. Install it with: Install-Module ImportExcel -Scope CurrentUser"
        exit 1
    }
}
Import-Module ImportExcel -ErrorAction Stop

# Create output directory
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    Write-Host "Created output directory: $OutputPath" -ForegroundColor Cyan
}

# Gather input files
$excelFiles = Get-ChildItem -Path $FolderPath -Filter $FilePattern -File | Sort-Object Name

if ($excelFiles.Count -eq 0) {
    Write-Warning "No files matched pattern '$FilePattern' in '$FolderPath'."
    exit 0
}

Write-Host "`nFound $($excelFiles.Count) file(s) matching '$FilePattern' in:" -ForegroundColor Cyan
Write-Host "  $FolderPath`n" -ForegroundColor Cyan

# ============================================================================
# Phase 1: Sheet Discovery
# ============================================================================

Write-Host "Discovering sheets across all workbooks..." -ForegroundColor Cyan

$fileSheetMap = @{}
$allDiscoveredSheets = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
$skippedFiles = [System.Collections.Generic.List[string]]::new()

foreach ($file in $excelFiles) {
    try {
        $sheetInfoList = Get-ExcelSheetInfo -Path $file.FullName -ErrorAction Stop
        $names = $sheetInfoList | ForEach-Object { $_.Name }
        $fileSheetMap[$file.FullName] = $names
        foreach ($name in $names) {
            $allDiscoveredSheets.Add($name) | Out-Null
        }
    }
    catch {
        Write-Warning "Skipping '$($file.Name)' during discovery: $_"
        $skippedFiles.Add($file.Name)
    }
}

# Determine target sheets
if ($SheetNames.Count -gt 0) {
    $targetSheets = $SheetNames
    Write-Host "Processing specified sheets: $($targetSheets -join ', ')" -ForegroundColor Cyan
    foreach ($requested in $SheetNames) {
        if (-not $allDiscoveredSheets.Contains($requested)) {
            Write-Warning "Requested sheet '$requested' was not found in any workbook."
        }
    }
}
else {
    $targetSheets = $allDiscoveredSheets | Sort-Object
    Write-Host "Auto-discovered $($targetSheets.Count) unique sheet(s):" -ForegroundColor Cyan
    foreach ($s in $targetSheets) {
        Write-Host "  - $s" -ForegroundColor White
    }
}

Write-Host ""

# ============================================================================
# Phase 2 & 3: Per-Sheet Import, Normalize, Export
# ============================================================================

$summary = [System.Collections.Generic.List[PSCustomObject]]::new()
$usedOutputNames = @{}
$exportFailures = 0

foreach ($sheetName in $targetSheets) {
    Write-Host "Processing sheet: '$sheetName'" -ForegroundColor Yellow
    Write-Host ("-" * 60)

    $allRows = [System.Collections.Generic.List[PSCustomObject]]::new()
    $allColumnNamesSet = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $allColumnNames = [System.Collections.Generic.List[string]]::new()
    $filesProcessedForSheet = 0

    foreach ($file in $excelFiles) {
        # Check if this file has the target sheet
        $sheetsInFile = $fileSheetMap[$file.FullName]
        if ($null -eq $sheetsInFile -or $sheetName -notin $sheetsInFile) {
            Write-Verbose "  '$($file.Name)' does not contain sheet '$sheetName'. Skipping."
            continue
        }

        try {
            $sheetData = Import-Excel -Path $file.FullName -WorksheetName $sheetName -ErrorAction Stop

            # Handle empty sheet
            if ($null -eq $sheetData -or @($sheetData).Count -eq 0) {
                Write-Verbose "  '$($file.Name)' sheet '$sheetName' is empty. Skipping."
                continue
            }

            # Ensure sheetData is always an array (single-row sheets return a single object)
            $sheetData = @($sheetData)

            # Derive "Month Year" from filename
            $monthLabel = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

            # Add "Month Year" to each row, collect column names from ALL rows (not just first)
            foreach ($row in $sheetData) {
                foreach ($col in $row.PSObject.Properties.Name) {
                    if ($allColumnNamesSet.Add($col)) {
                        $allColumnNames.Add($col)
                    }
                }
                $row | Add-Member -NotePropertyName "Month Year" -NotePropertyValue $monthLabel -Force
                $allRows.Add($row)
            }

            $filesProcessedForSheet++
            Write-Host "  Imported $($sheetData.Count) rows from '$($file.Name)'" -ForegroundColor Green
        }
        catch {
            Write-Warning "  Error reading sheet '$sheetName' from '$($file.Name)': $_"
            continue
        }
    }

    # Skip if no data collected
    if ($allRows.Count -eq 0) {
        Write-Warning "  No data found for sheet '$sheetName' across any file. Skipping."
        Write-Host ""
        continue
    }

    # --- Column Union & Normalization ---

    # Build final column order: "Month Year" first, then discovery order
    $finalColumns = [System.Collections.Generic.List[string]]::new()
    $finalColumns.Add("Month Year")
    foreach ($col in $allColumnNames) {
        if ($col -ne "Month Year") {
            $finalColumns.Add($col)
        }
    }

    # Normalize every row to have all columns
    Write-Host "  Normalizing $($allRows.Count) rows across $($finalColumns.Count) columns..." -ForegroundColor Cyan
    $normalizedRows = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($row in $allRows) {
        $hash = [ordered]@{}
        foreach ($col in $finalColumns) {
            $prop = $row.PSObject.Properties[$col]
            if ($null -ne $prop) {
                $hash[$col] = $prop.Value
            }
            else {
                $hash[$col] = ""
            }
        }
        $normalizedRows.Add([PSCustomObject]$hash)
    }

    # --- Export ---

    # Sanitize sheet name for filesystem (strip invalid chars and control characters)
    $safeSheetName = $sheetName -replace '[\\/:*?"<>|\x00-\x1F]', '_'
    $safeSheetName = $safeSheetName.Trim('. ')

    # Handle collision: if two sheet names sanitize to the same filename, append suffix
    $baseOutputName = "$safeSheetName All Months"
    if ($usedOutputNames.ContainsKey($baseOutputName.ToLower())) {
        $counter = $usedOutputNames[$baseOutputName.ToLower()] + 1
        $usedOutputNames[$baseOutputName.ToLower()] = $counter
        $outputFileName = "$baseOutputName ($counter).xlsx"
        Write-Warning "  Output name collision detected for '$sheetName'. Saving as '$outputFileName'."
    }
    else {
        $usedOutputNames[$baseOutputName.ToLower()] = 1
        $outputFileName = "$baseOutputName.xlsx"
    }
    $outputFile = Join-Path $OutputPath $outputFileName

    try {
        $normalizedRows | Export-Excel -Path $outputFile `
                                       -WorksheetName $sheetName `
                                       -AutoSize `
                                       -AutoFilter `
                                       -FreezeTopRow `
                                       -ClearSheet `
                                       -ErrorAction Stop

        Write-Host "  Exported $($normalizedRows.Count) rows -> '$outputFile'" -ForegroundColor Green
    }
    catch {
        Write-Error "  Failed to export sheet '$sheetName': $_"
        $exportFailures++
    }

    $summary.Add([PSCustomObject]@{
        Sheet        = $sheetName
        FilesMatched = $filesProcessedForSheet
        TotalRows    = $normalizedRows.Count
        TotalColumns = $finalColumns.Count
        OutputFile   = (Split-Path $outputFile -Leaf)
    })

    # Cleanup for next iteration
    $allRows.Clear()
    $normalizedRows.Clear()
    $allColumnNames.Clear()
    $allColumnNamesSet.Clear()

    Write-Host ""
}

# ============================================================================
# Phase 4: Summary
# ============================================================================

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  SUMMARY" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "Input folder  : $FolderPath"
Write-Host "File pattern  : $FilePattern"
Write-Host "Files found   : $($excelFiles.Count)"
if ($skippedFiles.Count -gt 0) {
    Write-Host "Files skipped : $($skippedFiles.Count) ($($skippedFiles -join ', '))" -ForegroundColor Yellow
}
Write-Host "Output folder : $OutputPath"
Write-Host ""

if ($summary.Count -gt 0) {
    $summary | Format-Table -AutoSize
}
else {
    Write-Warning "No sheets were processed."
}

if ($exportFailures -gt 0) {
    Write-Host "Done with $exportFailures export failure(s)." -ForegroundColor Yellow
    exit 1
}

Write-Host "Done." -ForegroundColor Green
