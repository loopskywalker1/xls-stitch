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

.PARAMETER ExcludeSheetNames
    Optional. Array of sheet names to exclude from processing. Applied after
    auto-discovery or -SheetNames filtering. Supports name normalization
    (dash-whitespace variants are matched unless -SkipNameNormalization is set).

.PARAMETER MaxHeaderRow
    Optional. Maximum row number to search for valid headers when a sheet has
    no headers on row 1 or has duplicate headers. Default: 5.

.PARAMETER SkipNameNormalization
    Optional switch. By default, sheet names are normalized so that variants
    differing only in whitespace around dashes are merged (e.g. "Report - OT",
    "Report -OT", and "Report-OT" are treated as the same sheet). Use this
    switch to disable normalization and treat each variant as a separate sheet.

.EXAMPLE
    .\Stitch-ExcelSheets.ps1 -FolderPath "C:\Data\Monthly Reports"

.EXAMPLE
    .\Stitch-ExcelSheets.ps1 -FolderPath "C:\Data\Monthly Reports" -SheetNames "Sheet1","Sheet2"

.EXAMPLE
    .\Stitch-ExcelSheets.ps1 -FolderPath "C:\Data\Monthly Reports" -ExcludeSheetNames "Sheet3","Sheet4"

.EXAMPLE
    .\Stitch-ExcelSheets.ps1 -FolderPath "C:\Data\Monthly Reports" -OutputPath "D:\Consolidated" -MaxHeaderRow 10

.EXAMPLE
    .\Stitch-ExcelSheets.ps1 -FolderPath "C:\Data\Monthly Reports" -SkipNameNormalization
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
    [string]$FilePattern = "*.xlsx",

    [Parameter(Mandatory = $false, HelpMessage = "Sheet names to exclude from processing.")]
    [string[]]$ExcludeSheetNames = @(),

    [Parameter(Mandatory = $false, HelpMessage = "Maximum row to search for valid headers (for sheets with missing or duplicate headers).")]
    [ValidateRange(1, 100)]
    [int]$MaxHeaderRow = 5,

    [Parameter(Mandatory = $false, HelpMessage = "Disable sheet name normalization (whitespace around dashes).")]
    [switch]$SkipNameNormalization
)

# ============================================================================
# Helper: Normalize sheet name (collapse whitespace around dashes)
# ============================================================================

function Get-NormalizedSheetName {
    param([string]$Name)
    # Collapse any whitespace around dashes/hyphens to " - "
    $normalized = $Name -replace '\s*-\s*', ' - '
    # Collapse multiple consecutive spaces to single space
    $normalized = $normalized -replace '\s+', ' '
    return $normalized.Trim()
}

# ============================================================================
# Helper: Robust Excel import with fallback for common errors
# ============================================================================

function Import-ExcelWithFallback {
    param(
        [string]$Path,
        [string]$WorksheetName,
        [int]$MaxStartRow = 5,
        [switch]$ForceDedupHeaders
    )

    $isDuplicate = $false
    $isNoHeaders = $false
    $errorMsg = ""

    if (-not $ForceDedupHeaders) {
        # Attempt 1: Normal import
        try {
            $data = Import-Excel -Path $Path -WorksheetName $WorksheetName -ErrorAction Stop
            return @{ Data = $data; Strategy = "normal" }
        }
        catch {
            $caughtError = $_
            $errorMsg = $_.Exception.Message
        }

        $isDuplicate = $errorMsg -match "Duplicate column headers"
        $isNoHeaders = $errorMsg -match "No column headers found"

        if (-not $isDuplicate -and -not $isNoHeaders) {
            throw $caughtError
        }
    }
    else {
        # Forced dedup: skip normal import, go straight to raw-import with dedup
        $isDuplicate = $true
    }

    # Attempt 2: Try increasing StartRow values (headers may not be on row 1)
    $fileName = [System.IO.Path]::GetFileName($Path)
    if ($isNoHeaders) {
        Write-Host "    [RETRY] '$fileName' tab '$WorksheetName': No headers on row 1. Trying alternate start rows..." -ForegroundColor DarkYellow
        foreach ($row in 2..$MaxStartRow) {
            try {
                $data = Import-Excel -Path $Path -WorksheetName $WorksheetName -StartRow $row -ErrorAction Stop
                if ($null -ne $data -and @($data).Count -gt 0) {
                    Write-Host "    [RESOLVED] Found valid headers at row $row." -ForegroundColor DarkCyan
                    return @{ Data = $data; Strategy = "startrow-$row" }
                }
            }
            catch {
                $retryMsg = $_.Exception.Message
                if ($retryMsg -match "Duplicate column headers") {
                    Write-Host "    [RETRY] Row $row has duplicate headers. Switching to raw-import with deduplication..." -ForegroundColor DarkYellow
                    break
                }
                continue
            }
        }
    }

    if ($isDuplicate -and -not $isNoHeaders) {
        Write-Host "    [RETRY] '$fileName' tab '$WorksheetName': Duplicate column headers detected. Using raw-import with deduplication..." -ForegroundColor DarkYellow
    }

    # Attempt 3: Import raw (no headers) and manually build objects with deduplicated headers
    # Handles both duplicate headers and sheets where no valid header row was found
    # For NoHeaders path, skip row 1 (already proven invalid); for Duplicate path, start at row 1
    $startFrom = if ($isNoHeaders -and -not $isDuplicate) { 2 } else { 1 }
    $lastAttemptError = $null

    foreach ($row in $startFrom..$MaxStartRow) {
        try {
            $raw = Import-Excel -Path $Path -WorksheetName $WorksheetName -NoHeader -StartRow $row -ErrorAction Stop
            if ($null -eq $raw -or @($raw).Count -lt 2) { continue }
            $raw = @($raw)

            # First row of raw data = candidate header row
            $headerRow = $raw[0]
            $propNames = @($headerRow.PSObject.Properties.Name)

            # Pass 1: Collect all raw header values to detect existing suffixed names
            $rawHeaderValues = [System.Collections.Generic.HashSet[string]]::new(
                [System.StringComparer]::OrdinalIgnoreCase
            )
            $nonBlankCount = 0
            foreach ($propName in $propNames) {
                $val = $headerRow.PSObject.Properties[$propName].Value
                if ($null -ne $val) { $val = $val.ToString().Trim() } else { $val = "" }
                if ($val -ne "") {
                    $nonBlankCount++
                    $rawHeaderValues.Add($val) | Out-Null
                }
            }

            # Require at least 2 non-blank header values to consider this a valid header row
            if ($nonBlankCount -lt 2) { continue }

            # Pass 2: Build deduplicated header list, avoiding collisions with existing names
            $headers = [System.Collections.Generic.List[string]]::new()
            $headerCounts = @{}

            foreach ($propName in $propNames) {
                $val = $headerRow.PSObject.Properties[$propName].Value
                if ($null -ne $val) { $val = $val.ToString().Trim() } else { $val = "" }
                if ($val -eq "") { $val = "Column" }

                if ($headerCounts.ContainsKey($val)) {
                    $headerCounts[$val]++
                    $candidate = "${val}_$($headerCounts[$val])"
                    # Ensure generated suffix doesn't collide with an existing column name
                    while ($rawHeaderValues.Contains($candidate)) {
                        $headerCounts[$val]++
                        $candidate = "${val}_$($headerCounts[$val])"
                    }
                    $val = $candidate
                }
                else {
                    $headerCounts[$val] = 1
                }
                $headers.Add($val)
            }

            # Build data rows with proper deduplicated headers
            $dataRows = [System.Collections.Generic.List[PSCustomObject]]::new()
            for ($i = 1; $i -lt $raw.Count; $i++) {
                $hash = [ordered]@{}
                $rowProps = @($raw[$i].PSObject.Properties.Name)
                for ($j = 0; $j -lt $headers.Count -and $j -lt $rowProps.Count; $j++) {
                    $hash[$headers[$j]] = $raw[$i].PSObject.Properties[$rowProps[$j]].Value
                }
                $dataRows.Add([PSCustomObject]$hash)
            }

            if ($dataRows.Count -gt 0) {
                # Report which columns were deduplicated
                $dupNames = @($headerCounts.GetEnumerator() | Where-Object { $_.Value -gt 1 } | ForEach-Object { "'$($_.Key)' (x$($_.Value))" })
                if ($dupNames.Count -gt 0) {
                    Write-Host "    [RESOLVED] Headers found at row $row. Deduplicated columns: $($dupNames -join ', ')" -ForegroundColor DarkCyan
                }
                else {
                    Write-Host "    [RESOLVED] Headers found at row $row via raw-import." -ForegroundColor DarkCyan
                }
                return @{ Data = $dataRows; Strategy = "noheader-dedup-startrow-$row" }
            }
        }
        catch {
            Write-Host "    [RETRY] Raw-import at row $row failed: $($_.Exception.Message)" -ForegroundColor DarkYellow
            $lastAttemptError = $_
            continue
        }
    }

    $detail = if ($lastAttemptError) { $lastAttemptError } else { $errorMsg }
    throw "All import strategies exhausted (tried normal, StartRow 2-$MaxStartRow, raw-import with dedup): $detail"
}

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

# $fileSheetMap: filePath -> array of raw sheet names in that file
$fileSheetMap = @{}
# $allRawSheets: set of every raw sheet name encountered
$allRawSheets = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
$skippedFiles = [System.Collections.Generic.List[string]]::new()

foreach ($file in $excelFiles) {
    try {
        $sheetInfoList = Get-ExcelSheetInfo -Path $file.FullName -ErrorAction Stop
        $names = $sheetInfoList | ForEach-Object { $_.Name }
        $fileSheetMap[$file.FullName] = $names
        foreach ($name in $names) {
            $allRawSheets.Add($name) | Out-Null
        }
    }
    catch {
        $reason = $_.Exception.Message
        if ($reason -match "being used by another process") {
            Write-Warning "[SKIP] '$($file.Name)': File is locked (open in another application). Close it and re-run to include this file."
        }
        elseif ($reason -match "corrupt" -or $reason -match "zip" -or $reason -match "invalid") {
            Write-Warning "[SKIP] '$($file.Name)': File appears to be corrupt or not a valid Excel file."
        }
        else {
            Write-Warning "[SKIP] '$($file.Name)': Could not read file during discovery. Reason: $reason"
        }
        $skippedFiles.Add($file.Name)
    }
}

# Build normalized-name mapping: canonical name -> list of raw variants
# When normalization is off, each raw name maps to itself
$normalizedMap = @{}
foreach ($rawName in $allRawSheets) {
    if ($SkipNameNormalization) {
        $canonical = $rawName
    }
    else {
        $canonical = Get-NormalizedSheetName $rawName
    }
    if (-not $normalizedMap.ContainsKey($canonical)) {
        $normalizedMap[$canonical] = @{
            CanonicalName = $canonical
            RawVariants   = [System.Collections.Generic.List[string]]::new()
        }
    }
    $normalizedMap[$canonical].RawVariants.Add($rawName)
}

# Log normalization merges
if (-not $SkipNameNormalization) {
    $mergedCount = 0
    foreach ($entry in $normalizedMap.Values) {
        if ($entry.RawVariants.Count -gt 1) {
            $mergedCount++
            Write-Host "  Merging variants into '$($entry.CanonicalName)':" -ForegroundColor Yellow
            foreach ($v in $entry.RawVariants) {
                Write-Host "    - '$v'" -ForegroundColor White
            }
        }
    }
    if ($mergedCount -gt 0) {
        Write-Host "  ($mergedCount sheet name group(s) merged by normalization)" -ForegroundColor Yellow
        Write-Host ""
    }
}

# Build the set of canonical sheet names for target selection
$allDiscoveredSheets = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
foreach ($entry in $normalizedMap.Values) {
    $allDiscoveredSheets.Add($entry.CanonicalName) | Out-Null
}

# Determine target sheets
if ($SheetNames.Count -gt 0) {
    # Normalize user-specified names too (unless skipped)
    if ($SkipNameNormalization) {
        $targetSheets = $SheetNames
    }
    else {
        $targetSheets = $SheetNames | ForEach-Object { Get-NormalizedSheetName $_ } | Select-Object -Unique
    }
    Write-Host "Processing specified sheets: $($targetSheets -join ', ')" -ForegroundColor Cyan
    foreach ($requested in $targetSheets) {
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

# Apply exclusions
if ($ExcludeSheetNames.Count -gt 0) {
    if ($SkipNameNormalization) {
        $normalizedExclusions = $ExcludeSheetNames
    }
    else {
        $normalizedExclusions = $ExcludeSheetNames | ForEach-Object { Get-NormalizedSheetName $_ } | Select-Object -Unique
    }
    $beforeCount = @($targetSheets).Count
    $targetSheets = @($targetSheets | Where-Object { $_ -notin $normalizedExclusions })
    $excludedCount = $beforeCount - @($targetSheets).Count
    if ($excludedCount -gt 0) {
        Write-Host "[EXCLUDE] Excluded $excludedCount sheet(s) from processing: $($normalizedExclusions -join ', ')" -ForegroundColor Yellow
    }
    else {
        Write-Host "[EXCLUDE] None of the specified exclusions matched discovered sheets: $($normalizedExclusions -join ', ')" -ForegroundColor DarkYellow
    }
}

if (@($targetSheets).Count -eq 0) {
    Write-Warning "No sheets to process after applying filters/exclusions."
    exit 0
}

Write-Host ""

# ============================================================================
# Phase 1.5: Pre-scan for sheets with duplicate headers
# ============================================================================

Write-Host "Pre-scanning for duplicate column headers..." -ForegroundColor Cyan

$sheetsWithDuplicates = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

foreach ($sheetName in $targetSheets) {
    if ($sheetsWithDuplicates.Contains($sheetName)) { continue }

    $rawVariants = @()
    if ($normalizedMap.ContainsKey($sheetName)) {
        $rawVariants = $normalizedMap[$sheetName].RawVariants
    }

    # Scan ALL files for this sheet — if ANY file has duplicate headers, flag the sheet
    foreach ($file in $excelFiles) {
        $sheetsInFile = $fileSheetMap[$file.FullName]
        if ($null -eq $sheetsInFile) { continue }

        $testVariant = $null
        foreach ($variant in $rawVariants) {
            if ($variant -in $sheetsInFile) {
                $testVariant = $variant
                break
            }
        }
        if ($null -eq $testVariant) { continue }

        # Try a normal import — only reading header row is enough to detect duplicates
        try {
            $null = Import-Excel -Path $file.FullName -WorksheetName $testVariant -EndRow 1 -ErrorAction Stop
        }
        catch {
            if ($_.Exception.Message -match "Duplicate column headers") {
                $sheetsWithDuplicates.Add($sheetName) | Out-Null
                Write-Host "  [DUPLICATE HEADERS] Sheet '$sheetName': Duplicate column names found in '$($file.Name)'. All files will use positional dedup (1st = original, 2nd = _2, etc.) for consistent cross-month matching." -ForegroundColor Yellow
                break
            }
            # Non-duplicate errors (e.g., no headers) — continue checking other files
        }
    }
}

if ($sheetsWithDuplicates.Count -eq 0) {
    Write-Host "  No duplicate headers detected." -ForegroundColor DarkGray
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
    $filesSkippedForSheet = 0
    $filesErroredForSheet = 0
    $fallbacksUsedForSheet = 0
    $emptyTabsForSheet = 0

    # Look up all raw variant names for this canonical sheet name
    $rawVariants = @()
    if ($normalizedMap.ContainsKey($sheetName)) {
        $rawVariants = $normalizedMap[$sheetName].RawVariants
    }

    foreach ($file in $excelFiles) {
        # Skip files that failed during discovery (not in $fileSheetMap)
        $sheetsInFile = $fileSheetMap[$file.FullName]
        if ($null -eq $sheetsInFile) {
            Write-Host "  [SKIP] '$($file.Name)': Was excluded during discovery (see earlier warnings)." -ForegroundColor DarkGray
            $filesSkippedForSheet++
            continue
        }

        # Find ALL matching variant tabs in this file (not just the first)
        $matchedVariants = [System.Collections.Generic.List[string]]::new()
        foreach ($variant in $rawVariants) {
            if ($variant -in $sheetsInFile) {
                $matchedVariants.Add($variant)
            }
        }
        if ($matchedVariants.Count -eq 0) {
            Write-Host "  [SKIP] '$($file.Name)': Does not contain any tab for sheet '$sheetName'." -ForegroundColor DarkGray
            $filesSkippedForSheet++
            continue
        }

        if ($matchedVariants.Count -gt 1) {
            Write-Host "  [MULTI-TAB] '$($file.Name)' contains $($matchedVariants.Count) variant tabs that all map to '$sheetName': $($matchedVariants -join ', '). Importing all." -ForegroundColor Yellow
        }

        # Import each matched variant tab from this file
        $anyVariantSucceeded = $false
        foreach ($matchedVariant in $matchedVariants) {
            try {
                $forceDedup = $sheetsWithDuplicates.Contains($sheetName)
                $importResult = Import-ExcelWithFallback -Path $file.FullName -WorksheetName $matchedVariant -MaxStartRow $MaxHeaderRow -ForceDedupHeaders:$forceDedup
                $sheetData = $importResult.Data
                $strategy = $importResult.Strategy

                # Handle empty sheet
                if ($null -eq $sheetData -or @($sheetData).Count -eq 0) {
                    Write-Host "  [EMPTY] '$($file.Name)' tab '$matchedVariant': Sheet exists but contains no data rows." -ForegroundColor DarkYellow
                    $emptyTabsForSheet++
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

                $anyVariantSucceeded = $true
                if ($strategy -ne "normal") {
                    $fallbacksUsedForSheet++
                    Write-Host "  [OK] Imported $($sheetData.Count) rows from '$($file.Name)' [tab: '$matchedVariant'] (fallback: $strategy)" -ForegroundColor Green
                }
                else {
                    Write-Host "  [OK] Imported $($sheetData.Count) rows from '$($file.Name)' [tab: '$matchedVariant']" -ForegroundColor Green
                }
            }
            catch {
                $filesErroredForSheet++
                $reason = $_.Exception.Message
                if ($reason -match "being used by another process") {
                    Write-Warning "  [ERROR] '$($file.Name)' tab '$matchedVariant': File is locked by another application."
                }
                elseif ($reason -match "All import strategies exhausted") {
                    Write-Warning "  [ERROR] '$($file.Name)' tab '$matchedVariant': Could not find valid headers after trying multiple strategies. This sheet may have an unusual layout."
                }
                else {
                    Write-Warning "  [ERROR] '$($file.Name)' tab '$matchedVariant': Import failed. Reason: $reason"
                }
                continue
            }
        }

        if ($anyVariantSucceeded) {
            $filesProcessedForSheet++
        }
    }

    # Skip if no data collected
    if ($allRows.Count -eq 0) {
        Write-Warning "  [NO DATA] Sheet '$sheetName': No data rows were collected from any file. This sheet will not produce an output file."
        if ($filesErroredForSheet -gt 0) {
            Write-Warning "    $filesErroredForSheet file(s) had import errors (see [ERROR] messages above)."
        }
        if ($emptyTabsForSheet -gt 0) {
            Write-Warning "    $emptyTabsForSheet tab(s) existed but contained no data rows."
        }
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
    if ($usedOutputNames.ContainsKey($baseOutputName)) {
        $counter = $usedOutputNames[$baseOutputName] + 1
        $usedOutputNames[$baseOutputName] = $counter
        $outputFileName = "$baseOutputName ($counter).xlsx"
        Write-Warning "  [COLLISION] Output filename '$baseOutputName.xlsx' already used by another sheet. Saving this sheet as '$outputFileName' instead."
    }
    else {
        $usedOutputNames[$baseOutputName] = 1
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

        Write-Host "  [EXPORTED] $($normalizedRows.Count) rows, $($finalColumns.Count) columns -> '$outputFile'" -ForegroundColor Green
    }
    catch {
        Write-Error "  [EXPORT FAILED] Sheet '$sheetName': Could not write output file. Reason: $_"
        $exportFailures++
    }

    # Per-sheet stats
    $issueNotes = [System.Collections.Generic.List[string]]::new()
    if ($filesSkippedForSheet -gt 0) { $issueNotes.Add("$filesSkippedForSheet skipped") }
    if ($filesErroredForSheet -gt 0) { $issueNotes.Add("$filesErroredForSheet errored") }
    if ($emptyTabsForSheet -gt 0)    { $issueNotes.Add("$emptyTabsForSheet empty") }
    if ($fallbacksUsedForSheet -gt 0) { $issueNotes.Add("$fallbacksUsedForSheet used fallback") }
    $issueStr = if ($issueNotes.Count -gt 0) { $issueNotes -join ", " } else { "none" }
    Write-Host "  Sheet stats: $filesProcessedForSheet file(s) imported, issues: $issueStr" -ForegroundColor Cyan

    $summary.Add([PSCustomObject]@{
        Sheet        = $sheetName
        FilesOK      = $filesProcessedForSheet
        Errors       = $filesErroredForSheet
        Skipped      = $filesSkippedForSheet
        Fallbacks    = $fallbacksUsedForSheet
        TotalRows    = $normalizedRows.Count
        Columns      = $finalColumns.Count
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

Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "  SUMMARY" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Input folder  : $FolderPath"
Write-Host "File pattern  : $FilePattern"
Write-Host "Files found   : $($excelFiles.Count)"
if ($skippedFiles.Count -gt 0) {
    Write-Host "Files skipped : $($skippedFiles.Count)" -ForegroundColor Yellow
    foreach ($sf in $skippedFiles) {
        Write-Host "    - $sf" -ForegroundColor Yellow
    }
}
Write-Host "Output folder : $OutputPath"
Write-Host ""

if ($summary.Count -gt 0) {
    $summary | Format-Table -AutoSize

    # Totals
    $totalRows = ($summary | Measure-Object -Property TotalRows -Sum).Sum
    $totalErrors = ($summary | Measure-Object -Property Errors -Sum).Sum
    $totalFallbacks = ($summary | Measure-Object -Property Fallbacks -Sum).Sum
    Write-Host "Total rows exported : $totalRows"
    Write-Host "Total sheets output : $($summary.Count)"

    # Issues recap
    if ($totalErrors -gt 0 -or $totalFallbacks -gt 0 -or $skippedFiles.Count -gt 0 -or $exportFailures -gt 0) {
        Write-Host ""
        Write-Host "  ISSUES" -ForegroundColor Yellow
        Write-Host ("  " + "-" * 56) -ForegroundColor Yellow
        if ($skippedFiles.Count -gt 0) {
            Write-Host "  [SKIP] $($skippedFiles.Count) file(s) could not be read during discovery (locked or corrupt). See [SKIP] messages above." -ForegroundColor Yellow
        }
        if ($totalErrors -gt 0) {
            Write-Host "  [ERROR] $totalErrors total import error(s) across all sheets. Some tabs could not be read. See [ERROR] messages above." -ForegroundColor Yellow
        }
        if ($totalFallbacks -gt 0) {
            Write-Host "  [FALLBACK] $totalFallbacks tab(s) required fallback import (duplicate headers or missing header row). Data was recovered but verify column names in output." -ForegroundColor Yellow
        }
        if ($exportFailures -gt 0) {
            Write-Host "  [EXPORT] $exportFailures sheet(s) failed to write output file. See [EXPORT FAILED] messages above." -ForegroundColor Red
        }
        Write-Host ""
    }
}
else {
    Write-Warning "No sheets were processed."
}

if ($exportFailures -gt 0) {
    Write-Host "Done with errors." -ForegroundColor Yellow
    exit 1
}

Write-Host "Done." -ForegroundColor Green
