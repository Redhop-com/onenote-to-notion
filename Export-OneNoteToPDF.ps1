# OneNote to PDF Export Script
# Exports all OneNote notebooks to individual PDF files (one per page)
# Maintains folder structure and uses hierarchical naming

#Requires -PSEdition Desktop

param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "$env:USERPROFILE\Documents\OneNoteExports",

    [Parameter(Mandatory=$false)]
    [bool]$IncludeSubpages = $true,

    [Parameter(Mandatory=$false)]
    [bool]$ShowProgress = $true
)

# Suppress COM error dialogs globally
$ErrorActionPreference = "Stop"

# Create output directory if it doesn't exist
if (!(Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    Write-Host "Created output directory: $OutputPath" -ForegroundColor Green
}

# Start transcript to log all console output
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = Join-Path $OutputPath "export_log_$timestamp.txt"
Start-Transcript -Path $logPath -Force | Out-Null
Write-Host "Logging to: $logPath" -ForegroundColor Gray
Write-Host "Export started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
Write-Host ""

# Function to clean caret notation from OneNote names
# OneNote sometimes stores control characters as caret notation (^J, ^M, ^I, etc.)
# ^J = Line Feed, ^M = Carriage Return, ^I = Tab
# The OneNote UI displays these as spaces or commas, so we replicate that behavior
function Get-CleanedName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ""
    }

    # Replace caret notation with spaces
    # ^J = Line Feed → space
    # ^M = Carriage Return → space
    # ^I = Tab → space
    $cleaned = $Name -replace '\^J', ', '  # Line feeds often represent list separators, use comma
    $cleaned = $cleaned -replace '\^M', ' '  # Carriage returns → space
    $cleaned = $cleaned -replace '\^I', ' '  # Tabs → space

    # Also handle actual control characters if they exist
    $cleaned = $cleaned -replace '[\r\n\t\f\v]', ' '

    # Replace multiple spaces with single space
    $cleaned = $cleaned -replace '\s+', ' '

    # Trim leading/trailing whitespace
    $cleaned = $cleaned.Trim()

    return $cleaned
}

# Function to sanitize filenames
function Get-SafeFileName {
    param([string]$FileName)

    # Remove or replace invalid characters
    $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
    $regex = "[{0}]" -f [regex]::Escape($invalidChars)
    $SafeName = $FileName -replace $regex, "_"

    # Trim dots and spaces from ends
    $SafeName = $SafeName.Trim('. ')

    # Limit length to prevent path issues
    if ($SafeName.Length -gt 100) {
        $SafeName = $SafeName.Substring(0, 100)
    }
    
    return $SafeName
}

# Function to export a page to PDF
function Export-PageToPDF {
    param(
        [object]$Page,
        [object]$OneNoteApp,
        [string]$NotebookName,
        [string]$SectionName,
        [string]$SectionGroupName = "",
        [string]$SectionUUID,
        [string]$SectionGroupUUID = "",
        [string]$ParentPageName = "",
        [string]$ParentPageUUID = "",
        [string]$OutputFolder,
        [int]$PageLevel = 0,
        [hashtable]$ExistingPageUUIDs = @{}
    )

    try {
        # Get page name
        $pageName = ""
        $pageXml = ""
        $pageId = $Page.ID

        # Validate page ID exists
        if ([string]::IsNullOrWhiteSpace($pageId)) {
            throw "Page ID is empty or invalid"
        }

        # Get the full page content to ensure we have the correct ID
        # Use try-catch to handle GetPageContent errors gracefully
        try {
            $OneNoteApp.GetPageContent($pageId, [ref]$pageXml)
        }
        catch {
            # If we can't get page content, skip this page silently
            throw "Cannot retrieve page content: $($_.Exception.Message)"
        }

        # Validate we got XML back
        if ([string]::IsNullOrWhiteSpace($pageXml)) {
            throw "Page content is empty"
        }

        $xml = [xml]$pageXml

        # Validate the XML structure
        if ($null -eq $xml.Page) {
            throw "Invalid page XML structure"
        }

        # Use the ID from the retrieved XML (this is the current/actual ID)
        $actualPageId = $xml.Page.ID

        # Extract title - OneNote stores it as an attribute called 'name'
        $pageName = $xml.Page.name
        if ([string]::IsNullOrWhiteSpace($pageName)) {
            # Fallback: try to get from Title element if it exists
            if ($xml.Page.Title) {
                if ($xml.Page.Title -is [string]) {
                    $pageName = $xml.Page.Title
                }
                else {
                    $pageName = $xml.Page.Title.InnerText
                }
            }
        }

        # Extract datetime properties from OneNote page
        # OneNote stores datetime information in various attributes
        $createdTime = $xml.Page.dateTime
        $lastModifiedTime = $xml.Page.lastModifiedTime

        # Validate actual page ID
        if ([string]::IsNullOrWhiteSpace($actualPageId)) {
            throw "Actual page ID is empty"
        }

        if ([string]::IsNullOrWhiteSpace($pageName)) {
            $pageName = "Untitled"
        }

        # Check if page has content (skip blank pages that cause 0x80042006 errors)
        $hasContent = $false
        if ($xml.Page.Outline) {
            $hasContent = $true
        }
        elseif ($xml.Page.Image) {
            $hasContent = $true
        }
        elseif ($xml.Page.InkDrawing) {
            $hasContent = $true
        }
        elseif ($xml.Page.MediaFile) {
            $hasContent = $true
        }

        if (-not $hasContent) {
            throw "Page is blank - skipping"
        }

        # Build the filename based on page hierarchy only
        # (notebook and section are already in the folder path)
        $safePageName = Get-SafeFileName $pageName

        if (![string]::IsNullOrWhiteSpace($ParentPageName)) {
            # This is a subpage - include parent hierarchy
            $safeParentName = Get-SafeFileName $ParentPageName
            $fileName = "${safeParentName}_${safePageName}.pdf"
        } else {
            # This is a top-level page - just the page name
            $fileName = "${safePageName}.pdf"
        }

        # Create full path and handle duplicates
        $fullPath = Join-Path $OutputFolder $fileName

        # If file already exists, append a number to make it unique
        if (Test-Path $fullPath) {
            $counter = 2
            $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
            $extension = [System.IO.Path]::GetExtension($fileName)

            do {
                $fileName = "${baseFileName}_${counter}${extension}"
                $fullPath = Join-Path $OutputFolder $fileName
                $counter++
            } while (Test-Path $fullPath)
        }

        # Export to PDF (Format: 3 = PDF) using the actual page ID from XML
        # Wrap in try-catch to prevent OneNote popup dialogs
        $isBlankPage = $false
        try {
            $OneNoteApp.Publish($actualPageId, $fullPath, 3, "")
        }
        catch {
            # Catch publish errors specifically to prevent popups
            $publishError = $_.Exception.Message
            if ($publishError -match "0x8004201A") {
                throw "Page no longer accessible (0x8004201A)"
            }
            elseif ($publishError -match "0x80042010") {
                throw "Section is encrypted or read-only (0x80042010)"
            }
            elseif ($publishError -match "0x80042006") {
                # Page is blank - mark it for special handling
                # We'll create a placeholder PDF so hierarchy is preserved
                $isBlankPage = $true
                Write-Host "    Note: Page is blank - creating placeholder for hierarchy preservation" -ForegroundColor Yellow
            }
            else {
                throw "Publish failed: $publishError"
            }
        }

        # If blank page, create a placeholder PDF
        if ($isBlankPage) {
            try {
                # Create a simple placeholder PDF using ReportLab or similar
                # For now, create a minimal text file that will serve as placeholder
                $placeholderContent = @"
Page: $pageName
Status: This page was blank in OneNote
UUID: $actualPageId
Notebook: $NotebookName
Section: $SectionName
Level: $PageLevel
"@
                # Create a minimal HTML that can be converted to PDF if needed
                $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>$pageName</title>
    <meta charset="UTF-8">
</head>
<body>
    <h1>$pageName</h1>
    <p><em>(This page was blank in OneNote)</em></p>
</body>
</html>
"@
                # Save as HTML temporarily - OneNote can publish HTML to PDF
                $tempHtmlPath = [System.IO.Path]::ChangeExtension($fullPath, ".html")
                Set-Content -Path $tempHtmlPath -Value $htmlContent -Encoding UTF8

                # Convert HTML to PDF using Windows built-in or keep as marker
                # For now, just create an empty PDF marker file
                # This ensures the file exists for reference but indicates it was blank
                $emptyPdfMarker = "%PDF-1.4`n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj 2 0 obj<</Type/Pages/Count 0/Kids[]>>endobj xref`n0 3`n0000000000 65535 f`n0000000009 00000 n`n0000000058 00000 n`ntrailer<</Size 3/Root 1 0 R>>startxref`n110`n%%EOF"
                Set-Content -Path $fullPath -Value $emptyPdfMarker -Encoding ASCII -NoNewline

                # Clean up temp HTML
                if (Test-Path $tempHtmlPath) {
                    Remove-Item $tempHtmlPath -Force -ErrorAction SilentlyContinue
                }
            }
            catch {
                # If placeholder creation fails, continue anyway
                Write-Host "    Warning: Could not create placeholder PDF: $($_.Exception.Message)" -ForegroundColor DarkYellow
            }
        }

        # Also export as Markdown for Notion import
        $mdFileName = [System.IO.Path]::ChangeExtension($fileName, ".md")
        $mdFullPath = Join-Path $OutputFolder $mdFileName

        try {
            # For blank pages, create minimal markdown with metadata
            if ($isBlankPage) {
                $mdContent = "# $pageName`n`n"
                $mdContent += "_[This page was blank in OneNote]_`n`n"
                $mdContent += "**Created:** $createdTime`n"
                $mdContent += "**Modified:** $lastModifiedTime`n"
                $mdContent += "**Section:** $SectionName`n"
                if (![string]::IsNullOrWhiteSpace($SectionGroupName)) {
                    $mdContent += "**Section Group:** $SectionGroupName`n"
                }
                if (![string]::IsNullOrWhiteSpace($ParentPageName)) {
                    $mdContent += "**Parent Page:** $ParentPageName`n"
                }

                # Save the minimal markdown
                Set-Content -Path $mdFullPath -Value $mdContent -Encoding UTF8 -Force
            }
            else {
                # Extract text content from OneNote page XML (normal processing)
                $mdContent = "# $pageName`n`n"

            # Add metadata
            $mdContent += "**Created:** $createdTime`n"
            $mdContent += "**Modified:** $lastModifiedTime`n"
            $mdContent += "**Section:** $SectionName`n"
            if (![string]::IsNullOrWhiteSpace($SectionGroupName)) {
                $mdContent += "**Section Group:** $SectionGroupName`n"
            }
            if (![string]::IsNullOrWhiteSpace($ParentPageName)) {
                $mdContent += "**Parent Page:** $ParentPageName`n"
            }
            $mdContent += "`n---`n`n"

            # Extract text from outline elements
            # OneNote XML structure: Page > Outline > OEChildren > OE
            # Text content is in OE.InnerText (includes text from T elements, OCR, tables, etc.)
            $contentFound = $false
            if ($xml.Page.Outline) {
                foreach ($outline in $xml.Page.Outline) {
                    if ($outline.OEChildren -and $outline.OEChildren.OE) {
                        foreach ($oe in $outline.OEChildren.OE) {
                            # Get text from OE.InnerText (most reliable method)
                            $text = $oe.InnerText

                            if (![string]::IsNullOrWhiteSpace($text)) {
                                # Clean up the text
                                $text = $text.Trim()

                                # Remove excessive whitespace and normalize line breaks
                                $text = $text -replace '\r\n', "`n"
                                $text = $text -replace '\n{3,}', "`n`n"

                                $mdContent += "$text`n`n"
                                $contentFound = $true
                            }
                        }
                    }
                }
            }

            # If no content was extracted, add a note
            if (-not $contentFound) {
                $mdContent += "_[No text content found in this page]_`n"
            }

            # Save markdown file
            $mdContent | Out-File -FilePath $mdFullPath -Encoding UTF8
            }
            # End of else block for normal (non-blank) pages
        }
        catch {
            # Markdown export is optional, don't fail if it doesn't work
            Write-Host "    Warning: Could not export markdown: $($_.Exception.Message)" -ForegroundColor Yellow
        }

        # Log progress
        if ($ShowProgress) {
            $indent = "  " * $PageLevel
            Write-Host "${indent}Exported: $pageName" -ForegroundColor Green  -NoNewline
            Write-Host "${indent}  -> $fileName (+ .md)" -ForegroundColor Gray
        }

        # Reuse existing UUID or generate new one (use PageID as key since it's stable)
        if ($ExistingPageUUIDs.ContainsKey($actualPageId)) {
            $pageUUID = $ExistingPageUUIDs[$actualPageId]
        } else {
            $pageUUID = [System.Guid]::NewGuid().ToString()
        }

        # Return metadata for JSON index
        # Note: Section and SectionGroup names are stored in the Sections array
        # Pages only need UUIDs for reliable parent linking
        return @{
            Success = $true
            PageName = $pageName
            PageUUID = $pageUUID  # UUID for sync tracking (prevent duplicates)
            SectionUUID = $SectionUUID  # UUID for reliable linking to section
            SectionGroupUUID = $SectionGroupUUID  # UUID for reliable linking to section group (if exists)
            NotebookName = $NotebookName  # Original display name from OneNote
            ParentPageUUID = $ParentPageUUID  # UUID of immediate parent page (for reliable linking)
            CreatedTime = $createdTime
            LastModifiedTime = $lastModifiedTime
            FilePath = $fullPath
            FileName = $fileName
            MarkdownFilePath = $mdFullPath
            MarkdownFileName = $mdFileName
            PageID = $actualPageId
            PageLevel = $PageLevel
            IsBlankPage = $isBlankPage  # Flag indicating this was a blank page with placeholder content
        }
    }
    catch {
        $errorMsg = $_.Exception.Message

        # Check for specific OneNote error codes
        # Note: Blank pages (0x80042006) are now handled inline and won't reach here
        if ($errorMsg -match "0x8004201A") {
            Write-Host "  Skipped: '$pageName' (page no longer accessible)" -ForegroundColor Yellow
        }
        elseif ($errorMsg -match "0x80042010") {
            Write-Host "  Skipped: '$pageName' (encrypted or read-only)" -ForegroundColor Yellow
        }
        else {
            Write-Host "  Error exporting '$pageName': $errorMsg" -ForegroundColor Red
        }

        return @{
            Success = $false
            PageName = $pageName
            Section = $SectionName
            SectionGroup = $SectionGroupName
            NotebookName = $NotebookName
            Error = $errorMsg
        }
    }
}

# Initialize OneNote COM object
Write-Host "`nInitializing OneNote application..." -ForegroundColor Cyan
try {
    $oneNote = New-Object -ComObject OneNote.Application

    # Disable OneNote popup dialogs using Windows API
    # This prevents "You can't export an empty page" dialogs from blocking the script
    Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        public class Win32 {
            [DllImport("user32.dll")]
            public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

            [DllImport("user32.dll")]
            public static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

            [DllImport("user32.dll")]
            public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

            public const uint WM_CLOSE = 0x0010;
        }
"@

    Write-Host "OneNote initialized (popup dialogs will be suppressed)" -ForegroundColor Green

    # Start a background job to automatically close OneNote dialog boxes
    $dialogCloserJob = Start-Job -ScriptBlock {
        Add-Type @"
            using System;
            using System.Runtime.InteropServices;
            public class DialogCloser {
                [DllImport("user32.dll", SetLastError = true)]
                public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

                [DllImport("user32.dll", CharSet = CharSet.Auto)]
                public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

                public const uint WM_CLOSE = 0x0010;
            }
"@

        while ($true) {
            Start-Sleep -Milliseconds 100

            # Find and close common OneNote error dialog boxes
            $dialogTitles = @(
                "Microsoft OneNote",
                "OneNote"
            )

            foreach ($title in $dialogTitles) {
                $hwnd = [DialogCloser]::FindWindow("#32770", $title)
                if ($hwnd -ne [IntPtr]::Zero) {
                    [DialogCloser]::SendMessage($hwnd, [DialogCloser]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
                }
            }
        }
    }

    Write-Host "Dialog suppression active (background job started)" -ForegroundColor Gray
}
catch {
    Write-Host "Failed to initialize OneNote. Make sure OneNote Desktop is installed." -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Important: Sync all notebooks before export to prevent ID mismatches
Write-Host "Syncing OneNote notebooks (this may take a moment)..." -ForegroundColor Cyan
try {
    # Force sync of all notebooks to ensure pages are up to date
    $oneNote.SyncHierarchy("")
    # Wait longer for sync to complete (especially important for cloud-synced notebooks)
    Start-Sleep -Seconds 5
    Write-Host "  Sync completed" -ForegroundColor Green
}
catch {
    Write-Host "Warning: Could not sync notebooks - some pages may fail to export" -ForegroundColor Yellow
}

# Get hierarchy with retry logic
Write-Host "Fetching OneNote hierarchy..." -ForegroundColor Cyan
$hierarchy = ""
$maxRetries = 3
$retryCount = 0
$success = $false

while (-not $success -and $retryCount -lt $maxRetries) {
    try {
        $oneNote.GetHierarchy("", 4, [ref]$hierarchy)
        $success = $true
    }
    catch {
        $retryCount++
        if ($retryCount -lt $maxRetries) {
            Write-Host "  Retry $retryCount/$maxRetries - waiting 2 seconds..." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            # Try syncing again
            try {
                $oneNote.SyncHierarchy("")
                Start-Sleep -Milliseconds 1000
            }
            catch {}
        }
        else {
            Write-Host "Failed to fetch OneNote hierarchy after $maxRetries attempts." -ForegroundColor Red
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "`nTroubleshooting steps:" -ForegroundColor Yellow
            Write-Host "  1. Close and reopen OneNote Desktop" -ForegroundColor White
            Write-Host "  2. Ensure all notebooks are fully synced" -ForegroundColor White
            Write-Host "  3. Try closing other programs using OneNote" -ForegroundColor White
            Write-Host "  4. Restart your computer if the issue persists" -ForegroundColor White
            exit 1
        }
    }
}

$xml = [xml]$hierarchy

# Display available notebooks and let user select which to export
Write-Host "`n" + ("=" * 60) -ForegroundColor Cyan
Write-Host "AVAILABLE NOTEBOOKS" -ForegroundColor Cyan
Write-Host ("=" * 60) -ForegroundColor Cyan

$notebookList = @()
$notebookIndex = 1

foreach ($notebook in $xml.Notebooks.Notebook) {
    # Get the notebook path and display name
    $notebookPath = ""
    if ($notebook.path) {
        $notebookPath = $notebook.path
    }

    $displayName = ""
    if ($notebook.nickname) {
        $displayName = $notebook.nickname
    }

    $notebookList += @{
        Index = $notebookIndex
        Name = $notebook.name
        DisplayName = $displayName
        Path = $notebookPath
        Object = $notebook
        Selected = $true  # Default: all selected
    }

    # Display notebook with display name and path
    Write-Host "[$notebookIndex] $($notebook.name)" -ForegroundColor White
    if (![string]::IsNullOrWhiteSpace($displayName)) {
        Write-Host "     Display Name: $displayName" -ForegroundColor Cyan
    }
    if (![string]::IsNullOrWhiteSpace($notebookPath)) {
        Write-Host "     Path: $notebookPath" -ForegroundColor Gray
    }

    $notebookIndex++
}

if ($notebookList.Count -eq 0) {
    Write-Host "`nNo notebooks found!" -ForegroundColor Red
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($oneNote) | Out-Null
    Write-Host "`nPress any key to exit..." -ForegroundColor Gray
    $null = Read-Host -Prompt "Press Enter to continue"
    exit 1
}

Write-Host "`n" + ("=" * 60) -ForegroundColor Gray
Write-Host "Selection Options:" -ForegroundColor Yellow
Write-Host "  - Press ENTER to export ALL notebooks" -ForegroundColor White
Write-Host "  - Enter numbers separated by commas (e.g., 1,3,5) to select specific notebooks" -ForegroundColor White
Write-Host "  - Enter 'none' or '0' to cancel" -ForegroundColor White

$selection = Read-Host "`nYour selection"

# Process selection
$selectedNotebooks = @()

if ([string]::IsNullOrWhiteSpace($selection)) {
    # Export all notebooks (default)
    $selectedNotebooks = $notebookList
    Write-Host "`nExporting ALL notebooks..." -ForegroundColor Green
}
elseif ($selection -eq "none" -or $selection -eq "0") {
    # User cancelled
    Write-Host "`nExport cancelled." -ForegroundColor Yellow
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($oneNote) | Out-Null
    Write-Host "`nPress any key to exit..." -ForegroundColor Gray
    $null = Read-Host -Prompt "Press Enter to continue"
    exit 0
}
else {
    # Parse comma-separated numbers
    $selectedIndices = $selection -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }

    foreach ($idx in $selectedIndices) {
        $notebook = $notebookList | Where-Object { $_.Index -eq $idx }
        if ($notebook) {
            $selectedNotebooks += $notebook
        }
        else {
            Write-Host "Warning: Invalid notebook number '$idx' - skipping" -ForegroundColor Yellow
        }
    }

    if ($selectedNotebooks.Count -eq 0) {
        Write-Host "`nNo valid notebooks selected. Exiting..." -ForegroundColor Red
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($oneNote) | Out-Null
        Write-Host "`nPress any key to exit..." -ForegroundColor Gray
        $null = Read-Host -Prompt "Press Enter to continue"
        exit 0
    }

    Write-Host "`nSelected notebooks:" -ForegroundColor Green
    foreach ($nb in $selectedNotebooks) {
        Write-Host "  - $($nb.Name)" -ForegroundColor White
    }
}

Write-Host ""

# Statistics
$totalNotebooks = 0
$totalSectionGroups = 0
$totalSections = 0
$totalPages = 0
$totalSubpages = 0
$exportedPages = 0
$skippedBlankPages = 0
$failedPages = 0
$failedPagesList = @()  # Track failed pages with details
$encryptedSectionsList = @()  # Track encrypted sections to handle at the end

# Function to process a section
function ProcessSection {
    param(
        [object]$Section,
        [string]$NotebookName,
        [string]$SectionName,
        [string]$SectionGroupName = "",
        [string]$SectionUUID,
        [string]$SectionGroupUUID = "",
        [string]$OutputFolder,
        [object]$OneNoteApp,
        [ref]$PageMetadataList,
        [hashtable]$ExistingPageUUIDs = @{}
    )

    $script:totalSections++

    Write-Host "`n  Section: $SectionName" -ForegroundColor Cyan

    # Check if section is encrypted
    $isEncrypted = $false
    if ($Section.encrypted -eq "true") {
        $isEncrypted = $true
        Write-Host "    [!] This section is PASSWORD-PROTECTED - will handle at end" -ForegroundColor Yellow

        # Add to encrypted sections list for later processing
        $script:encryptedSectionsList += @{
            Section = $Section
            NotebookName = $NotebookName
            SectionName = $SectionName
            SectionGroupName = $SectionGroupName
            SectionUUID = $SectionUUID
            SectionGroupUUID = $SectionGroupUUID
            OutputFolder = $OutputFolder
        }

        return  # Skip for now, will process at the end
    }

    # Create section folder
    $sectionFolder = Join-Path $OutputFolder (Get-SafeFileName $SectionName)
    Write-Host "    Creating section folder: $sectionFolder" -ForegroundColor Gray

    if (!(Test-Path $sectionFolder)) {
        New-Item -ItemType Directory -Path $sectionFolder -Force | Out-Null
        Write-Host "    Section folder created" -ForegroundColor Green
    } else {
        Write-Host "    Section folder already exists" -ForegroundColor Gray
    }

    # Get pages in this section
    $pagesXml = ""
    try {
        $OneNoteApp.GetHierarchy($Section.ID, 4, [ref]$pagesXml)
    }
    catch {
        if ($_.Exception.Message -match "0x80042010") {
            Write-Host "    Section is still encrypted - skipping" -ForegroundColor Yellow
            return
        }
        throw
    }
    $pagesDoc = [xml]$pagesXml

    # Check if we have pages
    if ($null -eq $pagesDoc.Section) {
        Write-Host "    Warning: Could not retrieve section data" -ForegroundColor Yellow
        return
    }

    $pageList = @($pagesDoc.Section.Page)  # Force array
    if ($pageList.Count -eq 0) {
        Write-Host "    No pages in this section" -ForegroundColor Gray
        return
    }

    Write-Host "    Processing $($pageList.Count) page(s)..." -ForegroundColor Gray

    # Build a flat list of ALL pages (including nested subpages) with their pageLevel
    $allPages = @()
    function Get-AllPagesFlat {
        param([object]$PageNode, [ref]$PagesList)

        $PagesList.Value += $PageNode

        # Recursively get subpages
        if ($PageNode.Page) {
            foreach ($subpage in $PageNode.Page) {
                Get-AllPagesFlat -PageNode $subpage -PagesList $PagesList
            }
        }
    }

    foreach ($page in $pageList) {
        Get-AllPagesFlat -PageNode $page -PagesList ([ref]$allPages)
    }

    # Track parent names by pageLevel for building hierarchical filenames
    $parentNamesByLevel = @{}
    # Track parent UUIDs by pageLevel for reliable parent linking
    $parentUUIDsByLevel = @{}

    # Process all pages in order, using pageLevel to determine parent-child relationships
    foreach ($page in $allPages) {
        $script:totalPages++

        # Get the pageLevel attribute
        $currentPageLevel = $page.pageLevel
        if ([string]::IsNullOrWhiteSpace($currentPageLevel)) {
            $currentPageLevel = 1  # Default to level 1 if not specified
        } else {
            $currentPageLevel = [int]$currentPageLevel
        }

        # Determine parent page name based on pageLevel (for filename construction)
        $parentPageName = ""
        $parentPageUUID = ""
        if ($currentPageLevel -gt 1) {
            # This is a subpage - build parent hierarchy for filename
            $parentParts = @()
            for ($level = 1; $level -lt $currentPageLevel; $level++) {
                if ($parentNamesByLevel.ContainsKey($level)) {
                    $parentParts += $parentNamesByLevel[$level]
                }
            }
            $parentPageName = $parentParts -join "_"

            # Get the immediate parent's UUID (one level up)
            # If immediate parent doesn't exist (gap in hierarchy), search upward for closest ancestor
            $parentPageUUID = ""
            for ($level = ($currentPageLevel - 1); $level -ge 1; $level--) {
                if ($parentUUIDsByLevel.ContainsKey($level)) {
                    $parentPageUUID = $parentUUIDsByLevel[$level]
                    if ($level -ne ($currentPageLevel - 1)) {
                        # Found ancestor at non-immediate level (hierarchy gap detected)
                        Write-Host "      Note: Hierarchy gap detected - linking to ancestor at level $level" -ForegroundColor Yellow
                    }
                    break
                }
            }

            $script:totalSubpages++
        }

        # Export the page
        $result = Export-PageToPDF -Page $page -OneNoteApp $OneNoteApp -NotebookName $NotebookName `
                                  -SectionName $SectionName -SectionGroupName $SectionGroupName `
                                  -SectionUUID $SectionUUID -SectionGroupUUID $SectionGroupUUID `
                                  -ParentPageName $parentPageName -ParentPageUUID $parentPageUUID `
                                  -OutputFolder $sectionFolder `
                                  -PageLevel $currentPageLevel -ExistingPageUUIDs $ExistingPageUUIDs

        if ($result.Success) {
            $script:exportedPages++

            # Track blank pages separately (they export successfully but with placeholder content)
            if ($result.IsBlankPage) {
                $script:skippedBlankPages++
            }

            # Add metadata to collection
            $PageMetadataList.Value += $result

            # Store this page's name and UUID at its level for future children
            # CRITICAL: This ensures blank parent pages preserve hierarchy for their subpages
            $parentNamesByLevel[$currentPageLevel] = $result.PageName
            $parentUUIDsByLevel[$currentPageLevel] = $result.PageUUID

            # Clear any deeper levels (in case we're backtracking in the hierarchy)
            $keysToRemove = @()
            foreach ($key in $parentNamesByLevel.Keys) {
                if ($key -gt $currentPageLevel) {
                    $keysToRemove += $key
                }
            }
            foreach ($key in $keysToRemove) {
                $parentNamesByLevel.Remove($key)
                $parentUUIDsByLevel.Remove($key)
            }
        } else {
            # Export failed for some other reason
            $script:failedPages++
            # Add to failed pages list with details
            $script:failedPagesList += @{
                PageName = $result.PageName
                NotebookName = $result.NotebookName
                Section = $result.Section
                SectionGroup = $result.SectionGroup
                Error = $result.Error
            }

            # Still store the page name for potential children (but NOT UUID since it failed)
            if ($result.PageName) {
                $parentNamesByLevel[$currentPageLevel] = $result.PageName
            }
        }
    }
}

# Process each selected notebook
Write-Host "`nStarting export process..." -ForegroundColor Cyan
Write-Host "Output directory: $OutputPath`n" -ForegroundColor Yellow

foreach ($selectedNotebook in $selectedNotebooks) {
    $notebook = $selectedNotebook.Object
    # Clean notebook name (removes caret notation like ^J, ^M, ^I)
    $notebookName = Get-CleanedName $notebook.name
    $totalNotebooks++

    # Use display name for folder name if available, otherwise use regular name
    $folderName = $notebookName
    if (![string]::IsNullOrWhiteSpace($selectedNotebook.DisplayName)) {
        $folderName = Get-CleanedName $selectedNotebook.DisplayName
    }

    Write-Host "`nNOTEBOOK: $notebookName" -ForegroundColor Magenta
    if (![string]::IsNullOrWhiteSpace($selectedNotebook.DisplayName)) {
        Write-Host "Display Name: $($selectedNotebook.DisplayName)" -ForegroundColor Cyan
    }
    Write-Host ("=" * 50) -ForegroundColor Gray

    # Create notebook folder using display name
    $notebookFolder = Join-Path $OutputPath (Get-SafeFileName $folderName)
    if (!(Test-Path $notebookFolder)) {
        New-Item -ItemType Directory -Path $notebookFolder -Force | Out-Null
    }

    # Initialize page metadata collection for this notebook
    $notebookPageMetadata = @()

    # Load existing JSON to preserve UUIDs for unchanged items
    $existingJson = $null
    $existingSectionUUIDs = @{}
    $existingSectionGroupUUIDs = @{}
    $existingPageUUIDs = @{}

    $jsonIndexPath = Join-Path $notebookFolder "index.json"
    if (Test-Path $jsonIndexPath) {
        Write-Host "  Found existing JSON - preserving UUIDs for unchanged items" -ForegroundColor Yellow
        try {
            $existingJson = Get-Content $jsonIndexPath -Raw | ConvertFrom-Json

            # Load existing section group UUIDs
            if ($existingJson.SectionGroups) {
                foreach ($sg in $existingJson.SectionGroups) {
                    $key = "$folderName|$($sg.Name)"
                    $existingSectionGroupUUIDs[$key] = $sg.UUID
                }
            }

            # Load existing section UUIDs
            if ($existingJson.Sections) {
                foreach ($section in $existingJson.Sections) {
                    $sgName = if ($section.SectionGroup) { $section.SectionGroup } else { "" }
                    $key = "$folderName|$sgName|$($section.Name)"
                    $existingSectionUUIDs[$key] = $section.UUID
                }
            }

            # Load existing page UUIDs by PageID (OneNote's internal ID)
            if ($existingJson.Pages) {
                foreach ($page in $existingJson.Pages) {
                    # Use PageID as key since it's unique and stable
                    if ($page.PageID) {
                        $existingPageUUIDs[$page.PageID] = $page.PageUUID
                    }
                }
            }

            Write-Host "  Loaded $($existingSectionGroupUUIDs.Count) section group UUIDs, $($existingSectionUUIDs.Count) section UUIDs, $($existingPageUUIDs.Count) page UUIDs" -ForegroundColor Gray
        }
        catch {
            Write-Host "  Warning: Could not load existing JSON: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    # Create hashtables to track section and section group UUIDs
    # Key: "NotebookName|SectionGroupName|SectionName" or "NotebookName|SectionGroupName"
    # Value: UUID
    $sectionUUIDs = @{}
    $sectionGroupUUIDs = @{}

    # Arrays to store section/section group info for JSON
    $sectionsInfo = @()
    $sectionGroupsInfo = @()

    # Process each section group (if any)
    foreach ($sectionGroup in $notebook.SectionGroup) {
        $totalSectionGroups++

        # Clean section group name (removes caret notation like ^J, ^M, ^I)
        $sectionGroupName = Get-CleanedName $sectionGroup.name
        Write-Host "`n  Section Group: $sectionGroupName" -ForegroundColor Yellow

        # Reuse existing UUID or generate new one
        $sectionGroupKey = "$folderName|$sectionGroupName"
        if ($existingSectionGroupUUIDs.ContainsKey($sectionGroupKey)) {
            $sectionGroupUUID = $existingSectionGroupUUIDs[$sectionGroupKey]
            Write-Host "    Reusing existing UUID" -ForegroundColor Gray
        } else {
            $sectionGroupUUID = [System.Guid]::NewGuid().ToString()
            Write-Host "    Generated new UUID" -ForegroundColor Green
        }
        $sectionGroupUUIDs[$sectionGroupKey] = $sectionGroupUUID

        # Store section group info with original name
        $sectionGroupsInfo += @{
            Name = $sectionGroupName
            UUID = $sectionGroupUUID
        }

        Write-Host "    UUID: $sectionGroupUUID" -ForegroundColor DarkGray

        # Create section group folder
        $sectionGroupFolder = Join-Path $notebookFolder (Get-SafeFileName $sectionGroupName)
        Write-Host "    Creating section group folder: $sectionGroupFolder" -ForegroundColor Gray

        if (!(Test-Path $sectionGroupFolder)) {
            New-Item -ItemType Directory -Path $sectionGroupFolder -Force | Out-Null
            Write-Host "    Section group folder created" -ForegroundColor Green
        } else {
            Write-Host "    Section group folder already exists" -ForegroundColor Gray
        }

        # Process sections within this section group (clean caret notation)
        foreach ($section in $sectionGroup.Section) {
            $sectionName = Get-CleanedName $section.name

            # Reuse existing UUID or generate new one
            $sectionKey = "$folderName|$sectionGroupName|$sectionName"
            if ($existingSectionUUIDs.ContainsKey($sectionKey)) {
                $sectionUUID = $existingSectionUUIDs[$sectionKey]
            } else {
                $sectionUUID = [System.Guid]::NewGuid().ToString()
            }
            $sectionUUIDs[$sectionKey] = $sectionUUID

            # Store section info with original name
            $sectionsInfo += @{
                Name = $sectionName
                SectionGroupUUID = $sectionGroupUUID  # UUID for reliable linking (empty string if no group)
                UUID = $sectionUUID
            }

            ProcessSection -Section $section -NotebookName $folderName -SectionName $sectionName `
                          -SectionGroupName $sectionGroupName -SectionUUID $sectionUUID `
                          -SectionGroupUUID $sectionGroupUUID -OutputFolder $sectionGroupFolder `
                          -OneNoteApp $oneNote -PageMetadataList ([ref]$notebookPageMetadata) `
                          -ExistingPageUUIDs $existingPageUUIDs
        }
    }

    # Process direct sections (not in a section group, clean caret notation)
    foreach ($section in $notebook.Section) {
        $sectionName = Get-CleanedName $section.name

        # Reuse existing UUID or generate new one
        $sectionKey = "$folderName||$sectionName"  # Empty string for section group
        if ($existingSectionUUIDs.ContainsKey($sectionKey)) {
            $sectionUUID = $existingSectionUUIDs[$sectionKey]
        } else {
            $sectionUUID = [System.Guid]::NewGuid().ToString()
        }
        $sectionUUIDs[$sectionKey] = $sectionUUID

        # Store section info with original name
        $sectionsInfo += @{
            Name = $sectionName
            SectionGroupUUID = ""  # Empty string indicates no section group
            UUID = $sectionUUID
        }

        ProcessSection -Section $section -NotebookName $folderName -SectionName $sectionName `
                      -SectionGroupName "" -SectionUUID $sectionUUID `
                      -SectionGroupUUID "" -OutputFolder $notebookFolder `
                      -OneNoteApp $oneNote -PageMetadataList ([ref]$notebookPageMetadata) `
                      -ExistingPageUUIDs $existingPageUUIDs
    }

    # Generate JSON index file for this notebook
    Write-Host "`n  Generating JSON index for notebook: $folderName" -ForegroundColor Cyan

    $notebookIndex = @{
        NotebookName = $notebookName  # Original name from OneNote
        DisplayName = $folderName  # Display name (may be same as NotebookName)
        OneNotePath = $selectedNotebook.Path
        ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        TotalPages = $notebookPageMetadata.Count
        Sections = $sectionsInfo  # Array of section objects with Name, UUID, SectionGroupUUID
        SectionGroups = $sectionGroupsInfo  # Array of section group objects with Name, UUID
        Pages = @($notebookPageMetadata | ForEach-Object {
            @{
                PageName = $_.PageName
                UUID = $_.PageUUID  # UUID for sync tracking (prevent duplicates)
                SectionUUID = $_.SectionUUID  # UUID for reliable linking to section
                SectionGroupUUID = $_.SectionGroupUUID  # UUID for reliable linking to section group
                NotebookName = $_.NotebookName  # Original notebook name
                ParentPageUUID = $_.ParentPageUUID  # UUID of immediate parent (reliable linking)
                CreatedTime = $_.CreatedTime
                LastModifiedTime = $_.LastModifiedTime
                FilePath = $_.FilePath
                FileName = $_.FileName
                MarkdownFilePath = $_.MarkdownFilePath
                MarkdownFileName = $_.MarkdownFileName
                PageLevel = $_.PageLevel
            }
        })
    }

    # Save JSON index file in the notebook root folder
    $jsonIndexPath = Join-Path $notebookFolder "index.json"
    $notebookIndex | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonIndexPath -Encoding UTF8
    Write-Host "  JSON index created: $jsonIndexPath" -ForegroundColor Green
}

# Handle encrypted sections at the end
if ($encryptedSectionsList.Count -gt 0) {
    Write-Host ("`n" + "=" * 60) -ForegroundColor Yellow
    Write-Host "PASSWORD-PROTECTED SECTIONS FOUND" -ForegroundColor Yellow
    Write-Host ("=" * 60) -ForegroundColor Gray
    Write-Host "`nThe following sections are password-protected and were skipped:" -ForegroundColor Cyan

    for ($i = 0; $i -lt $encryptedSectionsList.Count; $i++) {
        $encSection = $encryptedSectionsList[$i]
        $location = "$($encSection.NotebookName)"
        if (![string]::IsNullOrWhiteSpace($encSection.SectionGroupName)) {
            $location = "{0} > {1}" -f $location, $encSection.SectionGroupName
        }
        Write-Host "  [$($i+1)] $($encSection.SectionName)" -ForegroundColor White
        Write-Host "      Location: $location" -ForegroundColor Gray
    }

    Write-Host "`n" + ("=" * 60) -ForegroundColor Gray
    Write-Host "`nTo export these sections:" -ForegroundColor Cyan
    Write-Host "  1. Open OneNote Desktop" -ForegroundColor White
    Write-Host "  2. Click on each protected section" -ForegroundColor White
    Write-Host "  3. Enter the password when prompted to unlock it" -ForegroundColor White
    Write-Host "  4. Come back here and press Enter" -ForegroundColor White
    Write-Host "`nOr press 'S' to skip these sections and finish the export." -ForegroundColor Yellow

    $response = Read-Host "`nUnlock sections now and press Enter (or 'S' to skip)"

    if ($response -ne 'S' -and $response -ne 's') {
        Write-Host "`nAttempting to export unlocked sections..." -ForegroundColor Cyan

        foreach ($encSection in $encryptedSectionsList) {
            Write-Host "`n  Trying section: $($encSection.SectionName)" -ForegroundColor Cyan

            # Test if section is now accessible
            $testXml = ""
            try {
                $oneNote.GetHierarchy($encSection.Section.ID, 4, [ref]$testXml)
                Write-Host "    [OK] Section is now unlocked - exporting..." -ForegroundColor Green

                # Initialize metadata list for this section
                $sectionPageMetadata = @()

                # Process the section now that it's unlocked
                ProcessSection -Section $encSection.Section `
                              -NotebookName $encSection.NotebookName `
                              -SectionName $encSection.SectionName `
                              -SectionGroupName $encSection.SectionGroupName `
                              -SectionUUID $encSection.SectionUUID `
                              -SectionGroupUUID $encSection.SectionGroupUUID `
                              -OutputFolder $encSection.OutputFolder `
                              -OneNoteApp $oneNote `
                              -PageMetadataList ([ref]$sectionPageMetadata) `
                              -ExistingPageUUIDs @{}

                # Note: We're not updating the JSON index here for simplicity
                # The user can re-run a full export if they want complete JSON

            } catch {
                if ($_.Exception.Message -match "0x80042010") {
                    Write-Host "    [X] Section is still locked - skipping" -ForegroundColor Yellow
                } else {
                    Write-Host "    [X] Error accessing section: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }

        Write-Host "`nEncrypted sections processing complete!" -ForegroundColor Green
    } else {
        Write-Host "`nSkipping encrypted sections." -ForegroundColor Yellow
    }
}

# Clean up
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($oneNote) | Out-Null
Remove-Variable oneNote

# Display failed pages details if any (BEFORE summary)
if ($failedPages -gt 0) {
    Write-Host ("`n" + "=" * 60) -ForegroundColor Red
    Write-Host "FAILED PAGES DETAILS" -ForegroundColor Red
    Write-Host ("=" * 60) -ForegroundColor Gray
    foreach ($failedPage in $failedPagesList) {
        $location = "$($failedPage.NotebookName)"
        if (![string]::IsNullOrWhiteSpace($failedPage.SectionGroup)) {
            $location = "{0} {1} {2}" -f $location, [char]62, $failedPage.SectionGroup
        }
        $location = "{0} {1} {2}" -f $location, [char]62, $failedPage.Section

        Write-Host "  Page: $($failedPage.PageName)" -ForegroundColor Yellow
        Write-Host "    Location: $location" -ForegroundColor Gray
        Write-Host "    Reason: $($failedPage.Error)" -ForegroundColor Red
        Write-Host ""
    }
}

# Display summary AFTER failed pages
Write-Host ("`n" + "=" * 60) -ForegroundColor Green
Write-Host "EXPORT COMPLETE!" -ForegroundColor Green
Write-Host ("=" * 60) -ForegroundColor Green

Write-Host "`nSummary:" -ForegroundColor Cyan
Write-Host "  Notebooks processed: $totalNotebooks" -ForegroundColor White
Write-Host "  Section groups: $totalSectionGroups" -ForegroundColor White
Write-Host "  Sections processed: $totalSections" -ForegroundColor White
if ($encryptedSectionsList.Count -gt 0) {
    Write-Host "    - Password-protected sections skipped: $($encryptedSectionsList.Count)" -ForegroundColor Yellow
}
Write-Host "  Total pages found: $totalPages" -ForegroundColor White
Write-Host "  Subpages found: $totalSubpages" -ForegroundColor White
Write-Host "  Successfully exported: $exportedPages" -ForegroundColor Green
Write-Host "    - Including blank pages (with placeholders): $skippedBlankPages" -ForegroundColor Gray
Write-Host "  Failed exports: $failedPages" -ForegroundColor $(if ($failedPages -gt 0) { "Red" } else { "Gray" })

Write-Host "`nOutput location: $OutputPath" -ForegroundColor Yellow

# Write to log before stopping transcript
# Log failed pages details FIRST
if ($failedPages -gt 0) {
    Write-Host "`n" + ("=" * 60) -ForegroundColor Gray
    Write-Host "FAILED PAGES" -ForegroundColor Red
    Write-Host ("=" * 60) -ForegroundColor Gray
    foreach ($failedPage in $failedPagesList) {
        $location = "$($failedPage.NotebookName)"
        if (![string]::IsNullOrWhiteSpace($failedPage.SectionGroup)) {
            $location = "{0} {1} {2}" -f $location, [char]62, $failedPage.SectionGroup
        }
        $location = "{0} {1} {2}" -f $location, [char]62, $failedPage.Section
        Write-Host "  - $($failedPage.PageName) [$location]" -ForegroundColor Yellow
        Write-Host "    Error: $($failedPage.Error)" -ForegroundColor Red
    }
}

# Then write summary
Write-Host "`n" + ("=" * 60) -ForegroundColor Gray
Write-Host "EXPORT SUMMARY" -ForegroundColor Cyan
Write-Host ("=" * 60) -ForegroundColor Gray
Write-Host "Export completed at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
Write-Host "Notebooks: $totalNotebooks" -ForegroundColor White
Write-Host "Section Groups: $totalSectionGroups" -ForegroundColor White
Write-Host "Sections: $totalSections" -ForegroundColor White
if ($encryptedSectionsList.Count -gt 0) {
    Write-Host "  Password-Protected Sections Skipped: $($encryptedSectionsList.Count)" -ForegroundColor Yellow
}
Write-Host "Total Pages: $totalPages" -ForegroundColor White
Write-Host "Subpages: $totalSubpages" -ForegroundColor White
Write-Host "Successfully Exported: $exportedPages" -ForegroundColor Green
Write-Host "  Including Blank Pages (placeholders): $skippedBlankPages" -ForegroundColor Gray
Write-Host "Failed: $failedPages" -ForegroundColor $(if ($failedPages -gt 0) { "Red" } else { "Gray" })

Write-Host "`nOutput Path: $OutputPath" -ForegroundColor Yellow
Write-Host ("=" * 60) -ForegroundColor Gray

# Cleanup: Stop the dialog closer background job
if ($dialogCloserJob) {
    Stop-Job -Job $dialogCloserJob -ErrorAction SilentlyContinue
    Remove-Job -Job $dialogCloserJob -Force -ErrorAction SilentlyContinue
    Write-Host "Dialog suppression stopped" -ForegroundColor Gray
}

# Stop transcript to finalize the log
Stop-Transcript | Out-Null

Write-Host "`nComplete log saved to: $logPath" -ForegroundColor Green
Write-Host "`nPress any key to exit..." -ForegroundColor Gray
$null = Read-Host -Prompt "Press Enter to continue"
