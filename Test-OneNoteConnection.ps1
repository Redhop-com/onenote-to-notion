# OneNote Connection Test Script
# Tests connectivity to OneNote and scans notebook structure without exporting

#Requires -PSEdition Desktop

# Set up logging
$logPath = Join-Path $PSScriptRoot "test_connection_log.txt"
$script:logContent = @()

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] $Message"
    $script:logContent += $logLine
}

# Display banner
Clear-Host
Write-Host @"
==========================================================
        OneNote Connection & Structure Test Tool

  This script tests OneNote connectivity and scans
  your notebook structure WITHOUT performing exports
==========================================================
"@ -ForegroundColor Cyan

Write-Host "`nLog file will be created at: $logPath" -ForegroundColor Gray
Write-Host "`nInitializing OneNote connection test..." -ForegroundColor Yellow

Write-Log "=== OneNote Connection Test Started ==="

# Initialize OneNote COM object
try {
    $oneNote = New-Object -ComObject OneNote.Application
    Write-Host "[OK] Successfully connected to OneNote" -ForegroundColor Green
}
catch {
    Write-Host "[ERROR] Failed to connect to OneNote" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`nPlease ensure:" -ForegroundColor Yellow
    Write-Host "  - OneNote Desktop is installed (not Windows Store version)" -ForegroundColor White
    Write-Host "  - OneNote is not running in the background" -ForegroundColor White
    Write-Host "`nPress any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# Function to sanitize filenames (same as export script)
function Get-SafeFileName {
    param([string]$FileName)

    # Remove or replace invalid characters
    $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
    $regex = "[{0}]" -f [regex]::Escape($invalidChars)
    $SafeName = $FileName -replace $regex, "_"

    # Additional replacements for safety
    $SafeName = $SafeName -replace ':', '_'
    $SafeName = $SafeName -replace '\\', '_'
    $SafeName = $SafeName -replace '/', '_'
    $SafeName = $SafeName -replace '\*', '_'
    $SafeName = $SafeName -replace '\?', '_'
    $SafeName = $SafeName -replace '"', '_'
    $SafeName = $SafeName -replace '<', '_'
    $SafeName = $SafeName -replace '>', '_'
    $SafeName = $SafeName -replace '\|', '_'

    # Trim dots and spaces from ends
    $SafeName = $SafeName.Trim('. ')

    # Limit length to prevent path issues
    if ($SafeName.Length -gt 100) {
        $SafeName = $SafeName.Substring(0, 100)
    }

    return $SafeName
}

# Function to display page hierarchy
function Show-PageHierarchy {
    param(
        [object]$Page,
        [int]$Level = 0,
        [ref]$SubpageCount
    )

    # Get page name and pageLevel
    $pageName = $Page.name
    if ([string]::IsNullOrWhiteSpace($pageName)) {
        $pageName = "Untitled"
    }

    # Get pageLevel property from OneNote XML
    $pageLevel = $Page.pageLevel
    if ([string]::IsNullOrWhiteSpace($pageLevel)) {
        $pageLevel = "N/A"
    }

    # Display page with indentation
    $indent = "  " * ($Level + 2)
    if ($Level -eq 0) {
        Write-Host "${indent}[Page] $pageName (pageLevel=$pageLevel)" -ForegroundColor Cyan
        Write-Log "${indent}[Page] $pageName (pageLevel=$pageLevel)"
    } else {
        Write-Host "${indent}└─ [Subpage L$Level] $pageName (pageLevel=$pageLevel)" -ForegroundColor DarkCyan
        Write-Log "${indent}└─ [Subpage L$Level] $pageName (pageLevel=$pageLevel)"
    }

    # Process subpages recursively
    if ($Page.Page) {
        $subCount = ($Page.Page | Measure-Object).Count
        $SubpageCount.Value += $subCount

        foreach ($subpage in $Page.Page) {
            Show-PageHierarchy -Page $subpage -Level ($Level + 1) -SubpageCount $SubpageCount
        }
    }
}

# Get hierarchy
Write-Host "`nScanning OneNote hierarchy..." -ForegroundColor Cyan
try {
    $hierarchy = ""
    $oneNote.GetHierarchy("", 4, [ref]$hierarchy)
    $xml = [xml]$hierarchy
    Write-Host "[OK] Successfully retrieved notebook structure" -ForegroundColor Green
}
catch {
    Write-Host "[ERROR] Failed to retrieve notebook structure" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($oneNote) | Out-Null
    Write-Host "`nPress any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# Statistics
$totalNotebooks = 0
$totalSectionGroups = 0
$totalSections = 0
$totalPages = 0
$totalSubpages = 0
$encryptedSections = 0
$issues = @()

Write-Host "`n" + ("=" * 60) -ForegroundColor Gray
Write-Host "NOTEBOOK STRUCTURE ANALYSIS" -ForegroundColor Cyan
Write-Host ("=" * 60) -ForegroundColor Gray

# Process each notebook
foreach ($notebook in $xml.Notebooks.Notebook) {
    $notebookName = $notebook.name
    $totalNotebooks++

    Write-Host "`nNOTEBOOK: $notebookName" -ForegroundColor Magenta
    Write-Log ""
    Write-Log "NOTEBOOK: $notebookName"

    $notebookPageCount = 0

    # Process section groups
    foreach ($sectionGroup in $notebook.SectionGroup) {
        $totalSectionGroups++
        Write-Host "  Section Group: $($sectionGroup.name)" -ForegroundColor Yellow
        Write-Log "  Section Group: $($sectionGroup.name)"

        foreach ($section in $sectionGroup.Section) {
            $totalSections++
            $sectionName = "$($sectionGroup.name)_$($section.name)"
            Write-Log "    Section: $($section.name)"

            # Try to get pages
            try {
                $pagesXml = ""
                $oneNote.GetHierarchy($section.ID, 4, [ref]$pagesXml)
                $pagesDoc = [xml]$pagesXml

                $pageCount = 0
                $subpageCount = 0

                Write-Host "    Section: $($section.name)" -ForegroundColor White

                foreach ($page in $pagesDoc.Section.Page) {
                    $pageCount++
                    $totalPages++
                    $notebookPageCount++

                    # Use helper function to show hierarchy
                    $subCount = 0
                    Show-PageHierarchy -Page $page -Level 0 -SubpageCount ([ref]$subCount)

                    $subpageCount += $subCount
                    $totalSubpages += $subCount
                    $totalPages += $subCount
                    $notebookPageCount += $subCount
                }

                Write-Host "      Total: $pageCount pages, $subpageCount subpages" -ForegroundColor Gray

                if ($pageCount -eq 0) {
                    $issues += "Empty section: $notebookName - $sectionName"
                }
            }
            catch {
                $encryptedSections++
                Write-Host "    [LOCKED] Section: $($section.name) - ENCRYPTED or INACCESSIBLE" -ForegroundColor Red
                $issues += "Cannot access section: $notebookName - $sectionName (may be encrypted)"
            }
        }
    }

    # Process direct sections
    foreach ($section in $notebook.Section) {
        $totalSections++
        Write-Log "  Section: $($section.name)"

        # Try to get pages
        try {
            $pagesXml = ""
            $oneNote.GetHierarchy($section.ID, 4, [ref]$pagesXml)
            $pagesDoc = [xml]$pagesXml

            $pageCount = 0
            $subpageCount = 0

            Write-Host "  Section: $($section.name)" -ForegroundColor White

            foreach ($page in $pagesDoc.Section.Page) {
                $pageCount++
                $totalPages++
                $notebookPageCount++

                # Use helper function to show hierarchy
                $subCount = 0
                Show-PageHierarchy -Page $page -Level 0 -SubpageCount ([ref]$subCount)

                $subpageCount += $subCount
                $totalSubpages += $subCount
                $totalPages += $subCount
                $notebookPageCount += $subCount
            }

            Write-Host "    Total: $pageCount pages, $subpageCount subpages" -ForegroundColor Gray

            if ($pageCount -eq 0) {
                $issues += "Empty section: $notebookName - $($section.name)"
            }
        }
        catch {
            $encryptedSections++
            Write-Host "  [LOCKED] Section: $($section.name) - ENCRYPTED or INACCESSIBLE" -ForegroundColor Red
            $issues += "Cannot access section: $notebookName - $($section.name) (may be encrypted)"
        }
    }

    Write-Host "  Notebook total: $notebookPageCount pages" -ForegroundColor Cyan
}

# Clean up
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($oneNote) | Out-Null
Remove-Variable oneNote

# Display summary
Write-Host "`n" + ("=" * 60) -ForegroundColor Green
Write-Host "TEST SUMMARY" -ForegroundColor Green
Write-Host ("=" * 60) -ForegroundColor Green

Write-Host "`nStructure Statistics:" -ForegroundColor Cyan
Write-Host "  Total Notebooks: $totalNotebooks" -ForegroundColor White
Write-Host "  Section Groups: $totalSectionGroups" -ForegroundColor White
Write-Host "  Total Sections: $totalSections" -ForegroundColor White
Write-Host "  Total Pages: $totalPages" -ForegroundColor White
Write-Host "  Total Subpages: $totalSubpages" -ForegroundColor White

if ($encryptedSections -gt 0) {
    Write-Host "  Encrypted/Inaccessible Sections: $encryptedSections" -ForegroundColor Red
}

# Estimates
Write-Host "`nExport Estimates:" -ForegroundColor Cyan
$estimatedTimeMinutes = [math]::Ceiling($totalPages * 2 / 60)
$estimatedSizeMB = $totalPages * 0.5
Write-Host "  Estimated export time: ~$estimatedTimeMinutes minutes" -ForegroundColor White
Write-Host "  Estimated disk space: ~$([math]::Round($estimatedSizeMB, 1)) MB" -ForegroundColor White
Write-Host "  Total files to create: $totalPages PDFs" -ForegroundColor White

# Issues
if ($issues.Count -gt 0) {
    Write-Host "`nPotential Issues Found:" -ForegroundColor Yellow
    foreach ($issue in $issues) {
        Write-Host "  [WARNING] $issue" -ForegroundColor Yellow
    }
}

# Recommendations
Write-Host "`nRecommendations:" -ForegroundColor Cyan
if ($totalPages -gt 500) {
    Write-Host "  [WARNING] Large notebook collection - export may take significant time" -ForegroundColor Yellow
}
if ($encryptedSections -gt 0) {
    Write-Host "  [WARNING] Encrypted sections cannot be exported automatically" -ForegroundColor Yellow
}
if ($totalPages -eq 0) {
    Write-Host "  [WARNING] No pages found - nothing to export" -ForegroundColor Red
} else {
    Write-Host "  [OK] Ready to export using Export-OneNoteToPDF.ps1" -ForegroundColor Green
}

# Write log to file
Write-Log ""
Write-Log "=== TEST SUMMARY ==="
Write-Log "Total Notebooks: $totalNotebooks"
Write-Log "Section Groups: $totalSectionGroups"
Write-Log "Total Sections: $totalSections"
Write-Log "Total Pages: $totalPages"
Write-Log "Total Subpages: $totalSubpages"
if ($encryptedSections -gt 0) {
    Write-Log "Encrypted/Inaccessible Sections: $encryptedSections"
}
Write-Log "=== Test completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

try {
    $script:logContent | Out-File -FilePath $logPath -Encoding UTF8
    Write-Host "`nLog file saved to: $logPath" -ForegroundColor Green
}
catch {
    Write-Host "`nWarning: Could not save log file: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "`nPress any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
