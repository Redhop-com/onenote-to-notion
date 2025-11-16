# Import OneNote to Notion Script
# Imports exported OneNote notebooks (PDFs + JSON index) into Notion database
# Uses the JSON index files to preserve metadata (dates, hierarchy, etc.)

param(
    [Parameter(Mandatory=$false)]
    [string]$SourcePath = "$env:USERPROFILE\Documents\OneNoteExports",

    [Parameter(Mandatory=$false)]
    [string]$NotionApiKey = "",

    [Parameter(Mandatory=$false)]
    [string]$NotionDatabaseId = "",

    [Parameter(Mandatory=$false)]
    [int]$MaxPages = 0  # 0 = no limit, otherwise limit to first N pages
)

# Set error handling
$ErrorActionPreference = "Stop"

# Check PowerShell version (requires 7.0+)
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "ERROR: This script requires PowerShell 7.0 or higher" -ForegroundColor Red
    Write-Host "Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "To install PowerShell 7:" -ForegroundColor Cyan
    Write-Host "  1. Download from: https://aka.ms/powershell" -ForegroundColor White
    Write-Host "  2. Or use winget: winget install Microsoft.PowerShell" -ForegroundColor White
    Write-Host ""
    Write-Host "After installing, run this script with 'pwsh' instead of 'powershell'" -ForegroundColor Yellow
    exit 1
}

# Display banner
Clear-Host
Write-Host @"
==========================================================
        Import OneNote to Notion Tool

  This script imports exported OneNote notebooks into
  a Notion database using the JSON index files
==========================================================
"@ -ForegroundColor Cyan

Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Gray
Write-Host ""

# Start transcript to log all console output
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = Join-Path $SourcePath "import_log_$timestamp.txt"
Start-Transcript -Path $logPath -Force | Out-Null
Write-Host "Logging to: $logPath" -ForegroundColor Gray
Write-Host "Import started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
Write-Host ""

# Function to validate path
function Test-ExportPath {
    param([string]$Path)

    if (!(Test-Path $Path)) {
        Write-Host "Error: Source path does not exist: $Path" -ForegroundColor Red
        return $false
    }

    return $true
}

# Function to find notebooks (folders with index.json)
function Get-NotebookFolders {
    param([string]$BasePath)

    # Use ArrayList instead of array for better object handling
    $notebooks = New-Object System.Collections.ArrayList

    # Find all index.json files
    $indexFiles = Get-ChildItem -Path $BasePath -Filter "index.json" -Recurse -ErrorAction SilentlyContinue

    foreach ($indexFile in $indexFiles) {
        # Get the notebook folder (parent of index.json)
        $notebookFolder = $indexFile.Directory.FullName

        # Read the index file to get notebook info
        try {
            $indexContent = Get-Content $indexFile.FullName -Raw | ConvertFrom-Json

            # Create PSCustomObject instead of hashtable for more reliable handling
            $notebookInfo = [PSCustomObject]@{
                Path = $notebookFolder
                Name = $indexContent.NotebookName
                DisplayName = $indexContent.DisplayName
                TotalPages = $indexContent.TotalPages
                ExportDate = $indexContent.ExportDate
                IndexFile = $indexFile.FullName
                IndexData = $indexContent
            }

            [void]$notebooks.Add($notebookInfo)  # Use .Add() method
        }
        catch {
            Write-Host "Warning: Could not read index file: $($indexFile.FullName)" -ForegroundColor Yellow
        }
    }

    return ,$notebooks  # Return as single object
}

# Function to search for databases accessible to the integration
function Get-NotionDatabases {
    param(
        [string]$ApiKey
    )

    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }

    try {
        # Search for all databases accessible to this integration
        $searchBody = @{
            "filter" = @{
                "value" = "database"
                "property" = "object"
            }
        } | ConvertTo-Json

        $response = Invoke-RestMethod -Uri "https://api.notion.com/v1/search" `
                                     -Method Post `
                                     -Headers $headers `
                                     -Body $searchBody

        # Use ArrayList to avoid hashtable unwrapping issues
        $databases = New-Object System.Collections.ArrayList

        foreach ($db in $response.results) {
            # Extract database title
            $title = ""
            if ($db.title -and $db.title.Count -gt 0) {
                $titleParts = @()
                foreach ($titleBlock in $db.title) {
                    if ($titleBlock.plain_text) {
                        $titleParts += $titleBlock.plain_text
                    }
                }
                $title = $titleParts -join ""
            }

            if ([string]::IsNullOrWhiteSpace($title)) {
                $title = "Untitled Database"
            }

            # Create PSCustomObject instead of hashtable
            $dbInfo = [PSCustomObject]@{
                Id = $db.id
                Title = $title
                Url = $db.url
                LastEditedTime = $db.last_edited_time
            }

            [void]$databases.Add($dbInfo)
        }

        return ,$databases  # Comma prevents unwrapping
    }
    catch {
        Write-Host "Error searching for databases: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# Function to validate database has all required properties
function Test-NotionDatabaseProperties {
    param(
        [string]$DatabaseId,
        [string]$ApiKey
    )

    $requiredProperties = @{
        "Name" = "title"
        "Label" = "multi_select"
        "Type" = "select"
        "Notebook" = "relation"  # Now a relation to Notebooks database
        "Imported" = "date"
        "Last Modified" = "last_edited_time"
        "Attachment" = "files"
        "Parent item" = "relation"
    }

    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }

    try {
        # Get database schema
        $database = Invoke-RestMethod -Uri "https://api.notion.com/v1/databases/$DatabaseId" `
                                     -Method Get `
                                     -Headers $headers

        # Check each required property
        $missingProperties = New-Object System.Collections.ArrayList
        $wrongTypeProperties = New-Object System.Collections.ArrayList

        foreach ($propName in $requiredProperties.Keys) {
            $expectedType = $requiredProperties[$propName]

            if (-not $database.properties.PSObject.Properties[$propName]) {
                [void]$missingProperties.Add([PSCustomObject]@{
                    Name = $propName
                    ExpectedType = $expectedType
                })
            }
            else {
                $actualType = $database.properties.$propName.type
                if ($actualType -ne $expectedType) {
                    [void]$wrongTypeProperties.Add([PSCustomObject]@{
                        Name = $propName
                        ExpectedType = $expectedType
                        ActualType = $actualType
                    })
                }
            }
        }

        # Report results
        if ($missingProperties.Count -eq 0 -and $wrongTypeProperties.Count -eq 0) {
            return @{
                Valid = $true
                DatabaseTitle = if ($database.title) { ($database.title | ForEach-Object { $_.plain_text }) -join "" } else { "Untitled" }
            }
        }
        else {
            return @{
                Valid = $false
                DatabaseTitle = if ($database.title) { ($database.title | ForEach-Object { $_.plain_text }) -join "" } else { "Untitled" }
                Missing = $missingProperties
                WrongType = $wrongTypeProperties
            }
        }
    }
    catch {
        Write-Host "Error validating database: $($_.Exception.Message)" -ForegroundColor Red
        return @{
            Valid = $false
            Error = $_.Exception.Message
        }
    }
}

# Function to load cached Notion configuration
function Get-CachedConfig {
    param([string]$ConfigPath)

    if (Test-Path $ConfigPath) {
        try {
            $cached = Get-Content $ConfigPath -Raw | ConvertFrom-Json
            return @{
                ApiKey = $cached.ApiKey
                DatabaseId = $cached.DatabaseId
                NotebooksDatabaseId = $cached.NotebooksDatabaseId
            }
        }
        catch {
            Write-Host "Warning: Could not read cached config" -ForegroundColor Yellow
            return $null
        }
    }
    return $null
}

# Function to save Notion configuration to cache
function Save-ConfigCache {
    param(
        [string]$ConfigPath,
        [string]$ApiKey,
        [string]$DatabaseId,
        [string]$NotebooksDatabaseId = ""
    )

    try {
        $config = @{
            ApiKey = $ApiKey
            DatabaseId = $DatabaseId
            NotebooksDatabaseId = $NotebooksDatabaseId
            LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        $config | ConvertTo-Json | Out-File -FilePath $ConfigPath -Encoding UTF8
        Write-Host "`nConfiguration saved for future use" -ForegroundColor Green
    }
    catch {
        Write-Host "Warning: Could not save configuration cache" -ForegroundColor Yellow
    }
}

# Function to get Notion API configuration
function Get-NotionConfig {
    param(
        [string]$ApiKey,
        [string]$DatabaseId,
        [string]$ConfigCachePath
    )

    $config = @{}
    $NotebooksDatabaseId = ""

    # Try to load from cache
    $cachedConfig = Get-CachedConfig -ConfigPath $ConfigCachePath

    if ($cachedConfig) {
        Write-Host "`nFound cached Notion configuration" -ForegroundColor Green

        if ([string]::IsNullOrWhiteSpace($ApiKey) -and ![string]::IsNullOrWhiteSpace($cachedConfig.ApiKey)) {
            $ApiKey = $cachedConfig.ApiKey
            Write-Host "  Using cached API Key" -ForegroundColor Gray
        }

        if ([string]::IsNullOrWhiteSpace($DatabaseId) -and ![string]::IsNullOrWhiteSpace($cachedConfig.DatabaseId)) {
            $DatabaseId = $cachedConfig.DatabaseId
            Write-Host "  Using cached Pages Database ID" -ForegroundColor Gray
        }

        # Also load Notebooks Database ID from cache
        if (![string]::IsNullOrWhiteSpace($cachedConfig.NotebooksDatabaseId)) {
            $NotebooksDatabaseId = $cachedConfig.NotebooksDatabaseId
            Write-Host "  Using cached Notebooks Database ID" -ForegroundColor Gray
        }

        # Ask if user wants to use cached config or enter new
        Write-Host ""
        $useCached = Read-Host "Use cached configuration? (Y/n)"
        if ($useCached -eq "n" -or $useCached -eq "N") {
            $ApiKey = ""
            $DatabaseId = ""
            $NotebooksDatabaseId = ""
            Write-Host "Proceeding with new configuration..." -ForegroundColor Yellow
        }
    }

    # Get API Key
    if ([string]::IsNullOrWhiteSpace($ApiKey)) {
        Write-Host "`nNotion API Configuration Required" -ForegroundColor Yellow
        Write-Host ("=" * 60) -ForegroundColor Gray
        Write-Host "To use this script, you need a Notion Integration (API Key)" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Create an integration at: https://www.notion.so/my-integrations" -ForegroundColor Gray
        Write-Host "Then share your database(s) with the integration" -ForegroundColor Gray
        Write-Host ""

        $ApiKey = Read-Host "Enter your Notion API Key (or press ENTER to exit)"
        if ([string]::IsNullOrWhiteSpace($ApiKey)) {
            return $null
        }
    }

    # Search for available databases
    if ([string]::IsNullOrWhiteSpace($DatabaseId)) {
        Write-Host "`nSearching for databases accessible to your integration..." -ForegroundColor Cyan

        $databases = Get-NotionDatabases -ApiKey $ApiKey

        if ($databases.Count -eq 0) {
            Write-Host "`nNo databases found!" -ForegroundColor Red
            Write-Host "Make sure you have:" -ForegroundColor Yellow
            Write-Host "  1. Created at least one database in Notion" -ForegroundColor White
            Write-Host "  2. Shared the database(s) with your integration" -ForegroundColor White
            Write-Host "     (Click '...' menu â†’ 'Add connections' â†’ Select your integration)" -ForegroundColor Gray
            Write-Host ""
            return $null
        }

        # Display available databases
        Write-Host "`n" + ("=" * 60) -ForegroundColor Cyan
        Write-Host "AVAILABLE DATABASES" -ForegroundColor Cyan
        Write-Host ("=" * 60) -ForegroundColor Cyan

        $dbIndex = 1
        foreach ($db in $databases) {
            Write-Host "[$dbIndex] $($db.Title)" -ForegroundColor White
            Write-Host "    ID: $($db.Id)" -ForegroundColor Gray
            Write-Host "    URL: $($db.Url)" -ForegroundColor DarkGray
            Write-Host ""
            $dbIndex++
        }

        # Prompt for database selection
        Write-Host ("=" * 60) -ForegroundColor Gray
        Write-Host "Select a database:" -ForegroundColor Yellow
        Write-Host "  - Enter the number of the database to use" -ForegroundColor White
        Write-Host "  - Press ENTER to cancel" -ForegroundColor White

        $selection = Read-Host "`nYour selection"

        if ([string]::IsNullOrWhiteSpace($selection)) {
            return $null
        }

        # Validate selection
        if ($selection -match '^\d+$') {
            $selectedIndex = [int]$selection
            if ($selectedIndex -ge 1 -and $selectedIndex -le $databases.Count) {
                $selectedDb = $databases[$selectedIndex - 1]

                $DatabaseId = $selectedDb.Id
                $DatabaseTitle = $selectedDb.Title
                Write-Host "`nSelected database: $DatabaseTitle" -ForegroundColor Green
                Write-Host "Database ID: $DatabaseId" -ForegroundColor Gray
            }
            else {
                Write-Host "`nInvalid selection!" -ForegroundColor Red
                return $null
            }
        }
        else {
            Write-Host "`nInvalid selection!" -ForegroundColor Red
            return $null
        }
    }

    # Search for Notebooks database
    if ([string]::IsNullOrWhiteSpace($NotebooksDatabaseId)) {
        Write-Host "`nSearching for Notebooks database..." -ForegroundColor Cyan

        $databases = Get-NotionDatabases -ApiKey $ApiKey

        if ($databases.Count -eq 0) {
            Write-Host "`nNo databases found for Notebooks!" -ForegroundColor Red
            return $null
        }

        # Display available databases for Notebooks
        Write-Host "`n" + ("=" * 60) -ForegroundColor Cyan
        Write-Host "AVAILABLE DATABASES (for Notebooks)" -ForegroundColor Cyan
        Write-Host ("=" * 60) -ForegroundColor Cyan

        $dbIndex = 1
        foreach ($db in $databases) {
            Write-Host "[$dbIndex] $($db.Title)" -ForegroundColor White
            Write-Host "    ID: $($db.Id)" -ForegroundColor Gray
            Write-Host "    URL: $($db.Url)" -ForegroundColor DarkGray
            Write-Host ""
            $dbIndex++
        }

        # Prompt for Notebooks database selection
        Write-Host ("=" * 60) -ForegroundColor Gray
        Write-Host "Select the Notebooks database:" -ForegroundColor Yellow
        Write-Host "  - Enter the number of the database to use" -ForegroundColor White
        Write-Host "  - Press ENTER to cancel" -ForegroundColor White

        $selection = Read-Host "`nYour selection"

        if ([string]::IsNullOrWhiteSpace($selection)) {
            return $null
        }

        # Validate selection
        if ($selection -match '^\d+$') {
            $selectedIndex = [int]$selection
            if ($selectedIndex -ge 1 -and $selectedIndex -le $databases.Count) {
                $selectedDb = $databases[$selectedIndex - 1]

                $NotebooksDatabaseId = $selectedDb.Id
                $NotebooksDatabaseTitle = $selectedDb.Title
                Write-Host "`nSelected Notebooks database: $NotebooksDatabaseTitle" -ForegroundColor Green
                Write-Host "Database ID: $NotebooksDatabaseId" -ForegroundColor Gray
            }
            else {
                Write-Host "`nInvalid selection!" -ForegroundColor Red
                return $null
            }
        }
        else {
            Write-Host "`nInvalid selection!" -ForegroundColor Red
            return $null
        }
    }

    $config.ApiKey = $ApiKey
    $config.DatabaseId = $DatabaseId
    $config.NotebooksDatabaseId = $NotebooksDatabaseId

    # Save configuration to cache
    Save-ConfigCache -ConfigPath $ConfigCachePath -ApiKey $ApiKey -DatabaseId $DatabaseId -NotebooksDatabaseId $NotebooksDatabaseId

    return $config
}

# Function to upload file to Notion using File Upload API (3-step process)
function Upload-FileToNotion {
    param(
        [string]$FilePath,
        [string]$ApiKey,
        [string]$PageId
    )

    if (!(Test-Path $FilePath)) {
        Write-Host "    Warning: PDF file not found: $FilePath" -ForegroundColor Yellow
        return $null
    }

    try {
        # Get file info
        $fileInfo = Get-Item $FilePath
        $fileName = $fileInfo.Name
        $fileSize = $fileInfo.Length

        # Check file size limits (5MB for free, 5GB for paid, 20MB for single-part)
        $maxSinglePartSize = 20MB
        if ($fileSize -gt $maxSinglePartSize) {
            Write-Host "    Warning: PDF file too large for single-part upload ($([math]::Round($fileSize/1MB, 2))MB). Multipart upload not yet implemented." -ForegroundColor Yellow
            return $null
        }

        # Shorten filename if needed (max 900 bytes)
        $fileNameBytes = [System.Text.Encoding]::UTF8.GetByteCount($fileName)
        if ($fileNameBytes -gt 900) {
            $extension = [System.IO.Path]::GetExtension($fileName)
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
            $maxBaseLength = 900 - [System.Text.Encoding]::UTF8.GetByteCount($extension) - 10
            $fileName = $baseName.Substring(0, [Math]::Min($baseName.Length, $maxBaseLength)) + $extension
        }

        $headers = @{
            "Authorization" = "Bearer $ApiKey"
            "Notion-Version" = "2022-06-28"
        }

        # STEP 1: Create file upload
        Write-Host "      Step 1/3: Initiating upload..." -ForegroundColor DarkGray
        $createBody = @{
            "filename" = $fileName
            "content_type" = "application/pdf"
        } | ConvertTo-Json

        $headers["Content-Type"] = "application/json"
        $createResponse = Invoke-RestMethod -Uri "https://api.notion.com/v1/file_uploads" `
                                           -Method Post `
                                           -Headers $headers `
                                           -Body $createBody

        $fileUploadId = $createResponse.id
        $uploadUrl = $createResponse.upload_url

        # Debug: Check what we got back
        if (!$uploadUrl) {
            Write-Host "    Warning: Failed to get upload URL from Notion" -ForegroundColor Yellow
            Write-Host "    Response: $($createResponse | ConvertTo-Json -Depth 5)" -ForegroundColor DarkGray
            return $null
        }

        # Debug: Show upload URL (first 80 chars)
        Write-Host "      Upload URL: $($uploadUrl.Substring(0, [Math]::Min(80, $uploadUrl.Length)))..." -ForegroundColor DarkGray

        # STEP 2: Upload file data to the upload URL
        Write-Host "      Step 2/3: Uploading file data..." -ForegroundColor DarkGray

        # Use HttpClient for proper multipart/form-data handling
        $httpClient = New-Object System.Net.Http.HttpClient
        $httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer $ApiKey")
        $httpClient.DefaultRequestHeaders.Add("Notion-Version", "2022-06-28")

        try {
            # Create multipart form content
            $multipartContent = New-Object System.Net.Http.MultipartFormDataContent

            # Read file and create file content
            $fileStream = [System.IO.File]::OpenRead($FilePath)
            $fileContent = New-Object System.Net.Http.StreamContent($fileStream)

            # Set the content type to application/pdf
            $fileContent.Headers.ContentType = New-Object System.Net.Http.Headers.MediaTypeHeaderValue("application/pdf")

            # Add file to multipart content with correct filename
            $multipartContent.Add($fileContent, "file", $fileName)

            # Upload
            $response = $httpClient.PostAsync($uploadUrl, $multipartContent).Result

            if (-not $response.IsSuccessStatusCode) {
                $errorContent = $response.Content.ReadAsStringAsync().Result
                throw "Upload failed: $errorContent"
            }
        }
        finally {
            if ($fileStream) { $fileStream.Close() }
            if ($httpClient) { $httpClient.Dispose() }
        }

        # STEP 3: Return the file upload ID to be used in blocks
        Write-Host "      Step 3/3: Upload complete" -ForegroundColor DarkGray
        return $fileUploadId
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($_.ErrorDetails.Message) {
            try {
                $errorJson = $_.ErrorDetails.Message | ConvertFrom-Json
                $errorMsg = $errorJson.message
            }
            catch {}
        }
        Write-Host "    Warning: Failed to upload PDF: $errorMsg" -ForegroundColor Yellow
        return $null
    }
}

# Function to check if a page already exists in the database
function Find-ExistingNotionPage {
    param(
        [string]$SyncUUID,
        [string]$ApiKey,
        [string]$DatabaseId
    )

    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }

    # Query the database for pages with matching Sync-UUID
    $queryBody = @{
        "filter" = @{
            "property" = "Sync-UUID"
            "rich_text" = @{
                "equals" = $SyncUUID
            }
        }
    } | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri "https://api.notion.com/v1/databases/$DatabaseId/query" `
                                     -Method Post `
                                     -Headers $headers `
                                     -Body $queryBody

        if ($response.results -and $response.results.Count -gt 0) {
            return $response.results[0]
        }
        return $null
    }
    catch {
        Write-Host "      Warning: Could not query for existing page: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}

# Function to get or create a Notebook entry in the Notebooks database
function Get-OrCreateNotebook {
    param(
        [string]$NotebookName,
        [string]$NotebooksDatabaseId,
        [string]$ApiKey
    )

    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }

    # Search for existing notebook by name
    $queryBody = @{
        "filter" = @{
            "property" = "Name"
            "title" = @{
                "equals" = $NotebookName
            }
        }
    } | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri "https://api.notion.com/v1/databases/$NotebooksDatabaseId/query" `
                                     -Method Post `
                                     -Headers $headers `
                                     -Body $queryBody

        # If found, return existing notebook page ID
        if ($response.results -and $response.results.Count -gt 0) {
            return $response.results[0].id
        }

        # If not found, create new notebook entry
        $createBody = @{
            "parent" = @{
                "database_id" = $NotebooksDatabaseId
            }
            "properties" = @{
                "Name" = @{
                    "title" = @(
                        @{
                            "text" = @{
                                "content" = $NotebookName
                            }
                        }
                    )
                }
            }
        } | ConvertTo-Json -Depth 10

        $newNotebook = Invoke-RestMethod -Uri "https://api.notion.com/v1/pages" `
                                        -Method Post `
                                        -Headers $headers `
                                        -Body $createBody

        Write-Host "  Created notebook entry: $NotebookName" -ForegroundColor Green
        return $newNotebook.id
    }
    catch {
        Write-Host "  Warning: Could not get/create notebook '$NotebookName': $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}

# Function to create a simple hierarchical container page (Section Group or Section)
function New-NotionContainerPage {
    param(
        [string]$PageName,
        [string]$ContainerType,  # "Section Group" or "Section"
        [string]$NotebookPageId,  # Notion page ID of the notebook in Notebooks database
        [string]$SyncUUID,  # UUID for sync tracking
        [string]$ApiKey,
        [string]$DatabaseId,
        [string]$ParentNotionPageId = ""
    )

    # Check if page already exists
    $existingPage = Find-ExistingNotionPage -SyncUUID $SyncUUID -ApiKey $ApiKey -DatabaseId $DatabaseId
    if ($existingPage) {
        Write-Host "    Already exists (updating)" -ForegroundColor Yellow
        # Return existing page info
        return @{
            Success = $true
            PageId = $existingPage.id
            Url = $existingPage.url
            IsUpdate = $true
        }
    }

    # Build properties
    $properties = @{
        "Name" = @{
            "title" = @(
                @{
                    "text" = @{
                        "content" = $PageName
                    }
                }
            )
        }
        "Notebook" = @{
            "relation" = @(
                @{
                    "id" = $NotebookPageId
                }
            )
        }
        "Label" = @{
            # Label property logic:
            # - All containers (Section Groups and Sections) use their own name
            "multi_select" = @(
                @{
                    "name" = $PageName
                }
            )
        }
        "Type" = @{
            "select" = @{
                "name" = "Label"
            }
        }
        "Sync-UUID" = @{
            "rich_text" = @(@{ "text" = @{ "content" = $SyncUUID } })
        }
    }

    # Add Parent item relation if ParentNotionPageId is provided
    if (![string]::IsNullOrWhiteSpace($ParentNotionPageId)) {
        $properties["Parent item"] = @{
            "relation" = @(
                @{
                    "id" = $ParentNotionPageId
                }
            )
        }
    }

    # Create the page object
    $page = @{
        "parent" = @{
            "database_id" = $DatabaseId
        }
        "properties" = $properties
    }

    # Convert to JSON
    $body = $page | ConvertTo-Json -Depth 10

    # Make API request
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }

    try {
        $response = Invoke-RestMethod -Uri "https://api.notion.com/v1/pages" `
                                     -Method Post `
                                     -Headers $headers `
                                     -Body $body

        return @{
            Success = $true
            PageId = $response.id
            Url = $response.url
        }
    }
    catch {
        $errorDetails = $_.Exception.Message
        if ($_.ErrorDetails.Message) {
            try {
                $errorJson = $_.ErrorDetails.Message | ConvertFrom-Json
                $errorDetails = "Notion API Error: $($errorJson.message)"
            }
            catch {
                $errorDetails = $_.ErrorDetails.Message
            }
        }

        return @{
            Success = $false
            Error = $errorDetails
        }
    }
}

# Function to create a page in Notion
function New-NotionPage {
    param(
        [hashtable]$PageData,
        [string]$ApiKey,
        [string]$DatabaseId,
        [string]$PdfFilePath,
        [string]$MarkdownFilePath,
        [string]$NotebookPageId,
        [string]$ParentNotionPageId = ""
    )

    # Check if page already exists using PageUUID
    if (![string]::IsNullOrWhiteSpace($PageData.PageUUID)) {
        $existingPage = Find-ExistingNotionPage -SyncUUID $PageData.PageUUID -ApiKey $ApiKey -DatabaseId $DatabaseId
        if ($existingPage) {
            Write-Host "      Already exists (UUID: $($PageData.PageUUID))" -ForegroundColor Yellow
            # Return existing page info
            return @{
                Success = $true
                PageId = $existingPage.id
                Url = $existingPage.url
                IsUpdate = $true
            }
        }
    }

    # Prepare the page properties
    $properties = @{
        "Name" = @{
            "title" = @(
                @{
                    "text" = @{
                        "content" = $PageData.PageName
                    }
                }
            )
        }
        "Label" = @{
            "multi_select" = @(
                @{
                    "name" = $PageData.Section
                }
            )
        }
        "Notebook" = @{
            "relation" = @(
                @{
                    "id" = $NotebookPageId
                }
            )
        }
    }

    # Note: Label property contains the immediate parent Section name
    # Type property is NOT set for regular OneNote pages (only set for containers)

    # Add Imported Time if it exists (from OneNote CreatedTime)
    if (![string]::IsNullOrWhiteSpace($PageData.CreatedTime)) {
        try {
            # Parse the date and convert to ISO 8601 format
            # Handle both formats: "04/25/2009 22:02:26" and "2009-04-25T22:02:26.000Z"
            if ($PageData.CreatedTime -match '^\d{4}-\d{2}-\d{2}T') {
                # Already in ISO 8601 format
                $createdTimeFormatted = $PageData.CreatedTime
            } else {
                # Parse as DateTime and convert to ISO 8601
                # Use ParseExact for MM/dd/yyyy HH:mm:ss format
                $dateTime = [DateTime]::ParseExact($PageData.CreatedTime, "MM/dd/yyyy HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
                $createdTimeFormatted = $dateTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.000Z")
            }

            $properties["Imported"] = @{
                "date" = @{
                    "start" = $createdTimeFormatted
                }
            }
        }
        catch {
            Write-Host "      Warning: Could not parse imported date: $($PageData.CreatedTime)" -ForegroundColor Yellow
        }
    }

    # Note: Last Modified uses Notion's automatic "last_edited_time" property
    # It is not set manually during page creation

    # Add Sync-UUID if it exists
    if (![string]::IsNullOrWhiteSpace($PageData.PageUUID)) {
        $properties["Sync-UUID"] = @{
            "rich_text" = @(
                @{
                    "text" = @{
                        "content" = $PageData.PageUUID
                    }
                }
            )
        }
    }

    # Add Parent item relation if ParentNotionPageId is provided
    if (![string]::IsNullOrWhiteSpace($ParentNotionPageId)) {
        $properties["Parent item"] = @{
            "relation" = @(
                @{
                    "id" = $ParentNotionPageId
                }
            )
        }
        Write-Host "      Linking to parent page..." -ForegroundColor DarkGray
    }

    # Upload PDF file and add to Attachment property if file exists
    if (Test-Path $PdfFilePath) {
        Write-Host "      Uploading PDF file..." -ForegroundColor Gray
        $fileUploadId = Upload-FileToNotion -FilePath $PdfFilePath -ApiKey $ApiKey -PageId ""

        if ($fileUploadId) {
            # Add file to Attachment property
            # Notion has a 100 character limit on attachment filenames
            $fileName = [System.IO.Path]::GetFileName($PdfFilePath)
            if ($fileName.Length -gt 100) {
                # Truncate while preserving extension
                $extension = [System.IO.Path]::GetExtension($fileName)
                $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
                $maxNameLength = 100 - $extension.Length
                $fileName = $nameWithoutExt.Substring(0, $maxNameLength) + $extension
            }

            $properties["Attachment"] = @{
                "files" = @(
                    @{
                        "type" = "file_upload"
                        "name" = $fileName
                        "file_upload" = @{
                            "id" = $fileUploadId
                        }
                    }
                )
            }
            Write-Host "      PDF uploaded successfully" -ForegroundColor Green
        }
    }

    # Create the page object
    $page = @{
        "parent" = @{
            "database_id" = $DatabaseId
        }
        "properties" = $properties
    }

    # Add page content from markdown file
    $children = New-Object System.Collections.ArrayList

    if (![string]::IsNullOrWhiteSpace($MarkdownFilePath) -and (Test-Path $MarkdownFilePath)) {
        try {
            # Read markdown content
            $mdContent = Get-Content -Path $MarkdownFilePath -Raw -Encoding UTF8

            # Remove metadata section (everything before the --- separator)
            # The markdown format is: # Title, **Metadata**, ---, Content
            if ($mdContent -match '(?s)^.*?---\s*\n\n(.*)$') {
                # Extract content after the --- separator
                $mdContent = $matches[1]
            }
            # Fallback: if no separator found, remove lines starting with # or **
            else {
                $lines = $mdContent -split "`n"
                $contentLines = New-Object System.Collections.ArrayList
                $inContent = $false

                foreach ($line in $lines) {
                    # Skip title line (# ...)
                    if ($line -match '^#\s+') {
                        continue
                    }
                    # Skip metadata lines (**Field:** ...)
                    if ($line -match '^\*\*.*\*\*:') {
                        continue
                    }
                    # Skip separator line
                    if ($line -match '^---+\s*$') {
                        $inContent = $true
                        continue
                    }
                    # Add content lines
                    if ($inContent -or (!($line -match '^#') -and !($line -match '^\*\*'))) {
                        [void]$contentLines.Add($line)
                    }
                }

                $mdContent = $contentLines -join "`n"
            }

            # Split into paragraphs (double newline separated)
            $paragraphs = $mdContent -split "`n`n" | Where-Object { ![string]::IsNullOrWhiteSpace($_) }

            foreach ($para in $paragraphs) {
                # Clean up the text
                $cleanText = $para.Trim()
                if ([string]::IsNullOrWhiteSpace($cleanText)) {
                    continue
                }

                # Skip the "no content" placeholder
                if ($cleanText -match '^\[No text content found') {
                    continue
                }

                # Notion has a 2000 character limit per text block
                if ($cleanText.Length -gt 2000) {
                    $cleanText = $cleanText.Substring(0, 2000)
                }

                # Add as paragraph block
                [void]$children.Add(@{
                    "object" = "block"
                    "type" = "paragraph"
                    "paragraph" = @{
                        "rich_text" = @(
                            @{
                                "type" = "text"
                                "text" = @{
                                    "content" = $cleanText
                                }
                            }
                        )
                    }
                })
            }
        }
        catch {
            Write-Host "      Warning: Could not read markdown file: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    # If no markdown content was added, add PDF path as fallback
    if ($children.Count -eq 0) {
        [void]$children.Add(@{
            "object" = "block"
            "type" = "paragraph"
            "paragraph" = @{
                "rich_text" = @(
                    @{
                        "type" = "text"
                        "text" = @{
                            "content" = "PDF File: $PdfFilePath"
                        }
                    }
                )
            }
        })
    }

    # Notion API limit: 100 blocks per request
    # Split children into batches if needed
    $firstBatch = $children | Select-Object -First 100
    $remainingBlocks = $children | Select-Object -Skip 100

    $page["children"] = @($firstBatch)

    # Convert to JSON
    $body = $page | ConvertTo-Json -Depth 10

    # Make API request
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }

    try {
        # Write-Host $body -ForegroundColor Gray

        # Create the page first with first 100 blocks
        $response = Invoke-RestMethod -Uri "https://api.notion.com/v1/pages" `
                                     -Method Post `
                                     -Headers $headers `
                                     -Body $body

        $pageId = $response.id
        $pageUrl = $response.url

        # If there are remaining blocks, append them in batches of 100
        if ($remainingBlocks.Count -gt 0) {
            Write-Host "      Page has $($children.Count) blocks, appending remaining $($remainingBlocks.Count) in batches..." -ForegroundColor Yellow

            for ($i = 0; $i -lt $remainingBlocks.Count; $i += 100) {
                $batchSize = [Math]::Min(100, $remainingBlocks.Count - $i)
                $batch = $remainingBlocks | Select-Object -Skip $i -First $batchSize

                $appendBody = @{
                    "children" = @($batch)
                } | ConvertTo-Json -Depth 10

                Write-Host "      Appending batch: blocks $($i + 1)-$($i + $batchSize) of $($remainingBlocks.Count)" -ForegroundColor Gray

                [void](Invoke-RestMethod -Uri "https://api.notion.com/v1/blocks/$pageId/children" `
                                        -Method Patch `
                                        -Headers $headers `
                                        -Body $appendBody)

                # Rate limiting
                Start-Sleep -Milliseconds 350
            }

            Write-Host "      All blocks appended successfully" -ForegroundColor Green
        }

        # PDF already uploaded and attached via Attachment property

        return @{
            Success = $true
            PageId = $pageId
            Url = $pageUrl
        }
    }
    catch {
        # Try to get detailed error from Notion API
        $errorDetails = $_.Exception.Message
        if ($_.ErrorDetails.Message) {
            try {
                $errorJson = $_.ErrorDetails.Message | ConvertFrom-Json
                $errorDetails = "Notion API Error: $($errorJson.message)"
                if ($errorJson.code) {
                    $errorDetails += " (Code: $($errorJson.code))"
                }
            }
            catch {
                $errorDetails = $_.ErrorDetails.Message
            }
        }

        return @{
            Success = $false
            Error = $errorDetails
        }
    }
}

# Main script execution
Write-Host "Scanning for exported notebooks..." -ForegroundColor Cyan
Write-Host "Source path: $SourcePath`n" -ForegroundColor Gray

# Validate source path
if (!(Test-ExportPath -Path $SourcePath)) {
    Write-Host "`nPress any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# Find notebooks
$notebooks = Get-NotebookFolders -BasePath $SourcePath

if ($notebooks.Count -eq 0) {
    Write-Host "No notebooks found in the source path!" -ForegroundColor Red
    Write-Host "Make sure you have exported notebooks with index.json files." -ForegroundColor Yellow
    Write-Host "`nPress any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# Display found notebooks
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host "FOUND NOTEBOOKS" -ForegroundColor Cyan
Write-Host ("=" * 60) -ForegroundColor Cyan

$notebookIndex = 1
foreach ($notebook in $notebooks) {
    Write-Host "[$notebookIndex] $($notebook.DisplayName)" -ForegroundColor White
    Write-Host "    Path: $($notebook.Path)" -ForegroundColor Gray
    Write-Host "    Pages: $($notebook.TotalPages)" -ForegroundColor Cyan
    Write-Host "    Exported: $($notebook.ExportDate)" -ForegroundColor DarkGray
    Write-Host ""
    $notebookIndex++
}

# Prompt for notebook selection
Write-Host ("=" * 60) -ForegroundColor Gray
Write-Host "Selection Options:" -ForegroundColor Yellow
Write-Host "  - Press ENTER to import ALL notebooks" -ForegroundColor White
Write-Host "  - Enter numbers separated by commas (e.g., 1,3,5) to select specific notebooks" -ForegroundColor White
Write-Host "  - Enter 'none' or '0' to cancel" -ForegroundColor White

$selection = Read-Host "`nYour selection"

# Process selection
$selectedNotebooks = New-Object System.Collections.ArrayList

if ([string]::IsNullOrWhiteSpace($selection)) {
    # Import all notebooks (default)
    foreach ($nb in $notebooks) {
        [void]$selectedNotebooks.Add($nb)
    }
    Write-Host "`nImporting ALL notebooks..." -ForegroundColor Green
}
elseif ($selection -eq "none" -or $selection -eq "0") {
    # User cancelled
    Write-Host "`nImport cancelled." -ForegroundColor Yellow
    Write-Host "`nPress any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 0
}
else {
    # Parse selection
    $indices = $selection -split ',' | ForEach-Object { $_.Trim() }

    foreach ($index in $indices) {
        if ($index -match '^\d+$') {
            $idx = [int]$index

            if ($idx -ge 1 -and $idx -le $notebooks.Count) {
                $notebookToAdd = $notebooks[$idx - 1]
                [void]$selectedNotebooks.Add($notebookToAdd)  # Use .Add() method
            }
        }
    }

    if ($selectedNotebooks.Count -eq 0) {
        Write-Host "`nNo valid notebooks selected!" -ForegroundColor Red
        Write-Host "`nPress any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }

    Write-Host "`nImporting $($selectedNotebooks.Count) notebook(s)..." -ForegroundColor Green
}

# Define cache file path
$configCachePath = Join-Path $PSScriptRoot "notion_config.json"

# Get Notion configuration
Write-Host ""
$notionConfig = Get-NotionConfig -ApiKey $NotionApiKey `
                                 -DatabaseId $NotionDatabaseId `
                                 -ConfigCachePath $configCachePath

if ($null -eq $notionConfig) {
    Write-Host "`nImport cancelled - Notion configuration not provided." -ForegroundColor Yellow
    Write-Host "`nPress any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 0
}

Write-Host "`nNotion Configuration:" -ForegroundColor Cyan
Write-Host "  API Key: $($notionConfig.ApiKey.Substring(0, 10))..." -ForegroundColor Gray
Write-Host "  Database ID: $($notionConfig.DatabaseId)" -ForegroundColor Gray

# Validate database has required properties
Write-Host "`nValidating database properties..." -ForegroundColor Cyan
$validation = Test-NotionDatabaseProperties -DatabaseId $notionConfig.DatabaseId -ApiKey $notionConfig.ApiKey

if ($validation.Valid) {
    Write-Host " Database '$($validation.DatabaseTitle)' has all required properties" -ForegroundColor Green
}
else {
    Write-Host "`n" + ("=" * 60) -ForegroundColor Red
    Write-Host "DATABASE VALIDATION FAILED" -ForegroundColor Red
    Write-Host ("=" * 60) -ForegroundColor Red
    Write-Host ""
    Write-Host "Database: $($validation.DatabaseTitle)" -ForegroundColor Yellow
    Write-Host ""

    if ($validation.Missing -and $validation.Missing.Count -gt 0) {
        Write-Host "Missing Properties ($($validation.Missing.Count)):" -ForegroundColor Red
        foreach ($prop in $validation.Missing) {
            Write-Host "  âœ— $($prop.Name) (Type: $($prop.ExpectedType))" -ForegroundColor Red
        }
        Write-Host ""
    }

    if ($validation.WrongType -and $validation.WrongType.Count -gt 0) {
        Write-Host "Wrong Property Types ($($validation.WrongType.Count)):" -ForegroundColor Yellow
        foreach ($prop in $validation.WrongType) {
            Write-Host "  ! $($prop.Name)" -ForegroundColor Yellow
            Write-Host "    Expected: $($prop.ExpectedType)" -ForegroundColor Gray
            Write-Host "    Actual: $($prop.ActualType)" -ForegroundColor Gray
        }
        Write-Host ""
    }

    Write-Host "How to fix:" -ForegroundColor Cyan
    Write-Host "  Option 1: Run .\Run-CreateDatabase.ps1 to create a new database" -ForegroundColor White
    Write-Host "  Option 2: Manually add missing properties to your database in Notion" -ForegroundColor White
    Write-Host "  Option 3: See Setup-NotionDatabase.md for detailed instructions" -ForegroundColor White
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# Statistics
$totalPages = 0
$importedPages = 0
$failedPages = 0
$failedPagesList = @()

# Cache for page UUID to Notion page ID mapping
# This allows us to look up parent pages when creating sub-items using reliable UUIDs
$pageUUIDToIdCache = @{}

# Cache for Section Group and Section to Notion page ID mapping
# Key format: "NotebookName|SectionGroup" or "NotebookName|SectionGroup|Section"
$containerCache = @{}

# Track pages imported per level (for MaxPages per-level limiting)
$pageLevelCounts = @{}

# Import each selected notebook
Write-Host "`n" + ("=" * 60) -ForegroundColor Cyan
Write-Host "STARTING IMPORT PROCESS" -ForegroundColor Cyan
Write-Host ("=" * 60) -ForegroundColor Cyan

# Display MaxPages limit if set
if ($MaxPages -gt 0) {
    Write-Host "`n⚠️  MaxPages limit enabled: Will import only the first $MaxPages pages" -ForegroundColor Yellow
    Write-Host ""
}

foreach ($notebook in $selectedNotebooks) {
    Write-Host "`nNotebook: $($notebook.DisplayName)" -ForegroundColor Magenta
    Write-Host "Pages to import: $($notebook.TotalPages)" -ForegroundColor Cyan

    if ($notebook.TotalPages -eq 0) {
        Write-Host "  Skipping - no pages to import" -ForegroundColor Gray
        continue
    }

    # Extract Section Groups and Sections from the JSON metadata arrays
    # These arrays were populated during export with UUIDs for reliable identification
    $sectionGroups = @{}
    $sections = @{}

    # Get notebook name
    $notebookName = if (![string]::IsNullOrWhiteSpace($notebook.IndexData.NotebookName)) {
        $notebook.IndexData.NotebookName
    } else {
        $notebook.DisplayName
    }

    # Get or create the notebook entry in the Notebooks database
    Write-Host "  Getting/creating notebook entry: $notebookName" -ForegroundColor DarkGray
    $notebookPageId = Get-OrCreateNotebook -NotebookName $notebookName `
                                           -NotebooksDatabaseId $notionConfig.NotebooksDatabaseId `
                                           -ApiKey $notionConfig.ApiKey

    if ([string]::IsNullOrWhiteSpace($notebookPageId)) {
        Write-Host "  ERROR: Failed to get/create notebook entry" -ForegroundColor Red
        continue
    }
    Write-Host "  Notebook Page ID: $($notebookPageId.Substring(0,8))..." -ForegroundColor DarkGray

    # Load section groups from the SectionGroups array
    if ($notebook.IndexData.SectionGroups) {
        foreach ($sg in $notebook.IndexData.SectionGroups) {
            $sgUUID = $sg.UUID
            $sectionGroups[$sgUUID] = @{
                Name = $sg.Name
                NotebookName = $notebookName
                UUID = $sgUUID
            }
        }
    }

    # Load sections from the Sections array
    if ($notebook.IndexData.Sections) {
        foreach ($section in $notebook.IndexData.Sections) {
            $sectionUUID = $section.UUID
            $sections[$sectionUUID] = @{
                Name = $section.Name
                SectionGroupUUID = $section.SectionGroupUUID  # UUID for reliable linking (empty if no group)
                NotebookName = $notebookName
                UUID = $sectionUUID
            }
        }
    }

    # Step 1: Import Section Groups first
    Write-Host "`n  === Importing Section Groups ===" -ForegroundColor Yellow
    foreach ($sgUUID in $sectionGroups.Keys) {
        $sg = $sectionGroups[$sgUUID]
        Write-Host "  [SG] Creating: $($sg.Name)" -ForegroundColor Cyan
        Write-Host "      UUID: $sgUUID" -ForegroundColor DarkGray

        $result = New-NotionContainerPage -PageName $sg.Name `
                                          -ContainerType "Section Group" `
                                          -NotebookPageId $notebookPageId `
                                          -SyncUUID $sgUUID `
                                          -ApiKey $notionConfig.ApiKey `
                                          -DatabaseId $notionConfig.DatabaseId

        if ($result.Success) {
            # Store by UUID instead of name-based key
            $containerCache[$sgUUID] = $result.PageId
            Write-Host "    Created (Notion ID: $($result.PageId.Substring(0,8))...)" -ForegroundColor Green
        }
        else {
            Write-Host "    Failed: $($result.Error)" -ForegroundColor Red
        }

        Start-Sleep -Milliseconds 350
    }

    # Step 2: Import Sections
    Write-Host "`n  === Importing Sections ===" -ForegroundColor Yellow
    foreach ($sectionUUID in $sections.Keys) {
        $section = $sections[$sectionUUID]
        Write-Host "  [Section] Creating: $($section.Name)" -ForegroundColor Cyan
        Write-Host "      UUID: $sectionUUID" -ForegroundColor DarkGray

        # Look up parent Section Group by UUID if exists
        $parentSGId = ""
        if (![string]::IsNullOrWhiteSpace($section.SectionGroupUUID)) {
            $sgUUID = $section.SectionGroupUUID
            if ($containerCache.ContainsKey($sgUUID)) {
                $parentSGId = $containerCache[$sgUUID]
                $sectionGroupName = if ($sectionGroups.ContainsKey($sgUUID)) { $sectionGroups[$sgUUID].Name } else { "" }
                Write-Host "    Parent SG: $sectionGroupName (UUID: $sgUUID)" -ForegroundColor DarkGray
            }
        }

        $result = New-NotionContainerPage -PageName $section.Name `
                                          -ContainerType "Section" `
                                          -NotebookPageId $notebookPageId `
                                          -SyncUUID $sectionUUID `
                                          -ApiKey $notionConfig.ApiKey `
                                          -DatabaseId $notionConfig.DatabaseId `
                                          -ParentNotionPageId $parentSGId

        if ($result.Success) {
            # Store by UUID instead of name-based key
            $containerCache[$sectionUUID] = $result.PageId
            Write-Host "    Created (Notion ID: $($result.PageId.Substring(0,8))...)" -ForegroundColor Green
        }
        else {
            Write-Host "    Failed: $($result.Error)" -ForegroundColor Red
        }

        Start-Sleep -Milliseconds 350
    }

    # Step 3: Import Pages sorted by PageLevel (0, then 1, then 2, etc.)
    Write-Host "`n  === Importing Pages ===" -ForegroundColor Yellow
    $sortedPages = $notebook.IndexData.Pages | Sort-Object PageLevel

    # Pre-scan: Find which parent pages are needed by subpages that will be imported
    $requiredParents = @{}
    if ($MaxPages -gt 0) {
        $tempLevelCounts = @{}
        foreach ($pageData in $sortedPages) {
            $currentLevel = $pageData.PageLevel
            if (!$tempLevelCounts.ContainsKey($currentLevel)) {
                $tempLevelCounts[$currentLevel] = 0
            }

            # If this page will be imported (within limit)
            if ($tempLevelCounts[$currentLevel] -lt $MaxPages) {
                $tempLevelCounts[$currentLevel]++

                # If it's a subpage, mark its parent as required (by UUID)
                if ($currentLevel -gt 0 -and ![string]::IsNullOrWhiteSpace($pageData.ParentPageUUID)) {
                    $requiredParents[$pageData.ParentPageUUID] = $true
                }
            }
        }
        Write-Host "  Pre-scan: Found $($requiredParents.Count) parent pages required by subpages" -ForegroundColor DarkCyan
    }

    $pageNumber = 1
    foreach ($pageData in $sortedPages) {
        $totalPages++

        # Check if we've reached the MaxPages limit for this page level
        $currentLevel = $pageData.PageLevel
        if (!$pageLevelCounts.ContainsKey($currentLevel)) {
            $pageLevelCounts[$currentLevel] = 0
        }

        # Check if this page should be imported
        $shouldImport = $false
        $skipReason = ""
        $isRequiredParent = $false

        if ($MaxPages -gt 0 -and $pageLevelCounts[$currentLevel] -ge $MaxPages) {
            # Check if this is a required parent page (by UUID)
            if (![string]::IsNullOrWhiteSpace($pageData.UUID) -and $requiredParents.ContainsKey($pageData.UUID)) {
                $shouldImport = $true
                $isRequiredParent = $true
                $skipReason = "(Required parent - importing despite limit)"
            } else {
                $shouldImport = $false
                $skipReason = "MaxPages limit ($MaxPages) reached for level $currentLevel"
            }
        } else {
            $shouldImport = $true
        }

        if (-not $shouldImport) {
            Write-Host "  [$pageNumber/$($notebook.TotalPages)] [Level $currentLevel] Skipping: $($pageData.PageName) - $skipReason" -ForegroundColor Yellow
            $pageNumber++
            continue
        }

        $statusMsg = if ($isRequiredParent) { "$skipReason" } else { "" }
        Write-Host "  [$pageNumber/$($notebook.TotalPages)] [Level $currentLevel] Importing: $($pageData.PageName) $statusMsg" -ForegroundColor White

        # Get notebook name from page data (original OneNote name, not folder name)
        $pageNotebookName = if (![string]::IsNullOrWhiteSpace($pageData.NotebookName)) {
            $pageData.NotebookName
        } else {
            $notebook.DisplayName
        }

        # Determine Label property: Use immediate parent Section name
        $sectionPropertyValue = ""
        if (![string]::IsNullOrWhiteSpace($pageData.SectionUUID) -and $sections.ContainsKey($pageData.SectionUUID)) {
            $sectionData = $sections[$pageData.SectionUUID]
            # Always use the Section name (immediate parent)
            $sectionPropertyValue = $sectionData.Name
        }

        # Create hashtable for page data with notebook name
        $pageInfo = @{
            PageName = $pageData.PageName
            PageUUID = $pageData.UUID  # UUID for sync tracking
            Section = $sectionPropertyValue  # Section name (immediate parent)
            CreatedTime = $pageData.CreatedTime
            LastModifiedTime = $pageData.LastModifiedTime
            PageLevel = $pageData.PageLevel
            NotebookName = $pageNotebookName  # Use original OneNote name
        }

        # Determine parent based on page level
        $parentNotionPageId = ""

        # Level 1 pages (top-level in section) link to their Section
        # Note: OneNote uses pageLevel=1 for top-level pages, not 0
        if ($pageData.PageLevel -eq 1) {
            # Look up section using UUID for reliable linking (handles duplicate names)
            if (![string]::IsNullOrWhiteSpace($pageData.SectionUUID)) {
                $sectionUUID = $pageData.SectionUUID

                if ($containerCache.ContainsKey($sectionUUID)) {
                    $parentNotionPageId = $containerCache[$sectionUUID]
                    Write-Host "    Parent: Section '$sectionName' (UUID: $($sectionUUID.Substring(0,8))...)" -ForegroundColor DarkGray
                } else {
                    Write-Host "    Warning: Section UUID '$sectionUUID' not found in cache" -ForegroundColor Yellow
                }
            } else {
                Write-Host "    Warning: Page missing SectionUUID - cannot link to parent" -ForegroundColor Yellow
            }
        }
        # Subpages (level > 1) link to their parent page using UUID
        elseif ($pageData.PageLevel -gt 1) {
            if (![string]::IsNullOrWhiteSpace($pageData.ParentPageUUID)) {
                if ($pageUUIDToIdCache.ContainsKey($pageData.ParentPageUUID)) {
                    $parentNotionPageId = $pageUUIDToIdCache[$pageData.ParentPageUUID]
                    Write-Host "    Parent: Page UUID '$($pageData.ParentPageUUID.Substring(0,8))...' (Level $($pageData.PageLevel))" -ForegroundColor DarkGray
                } else {
                    # Parent page not found in cache - this can happen if hierarchy has gaps
                    # Fall back to linking to the section instead
                    Write-Host "    Warning: Parent page UUID not found - linking to section instead" -ForegroundColor Yellow

                    if (![string]::IsNullOrWhiteSpace($pageData.SectionUUID)) {
                        if ($containerCache.ContainsKey($pageData.SectionUUID)) {
                            $parentNotionPageId = $containerCache[$pageData.SectionUUID]
                            Write-Host "    Fallback: Linking to Section (UUID: $($pageData.SectionUUID.Substring(0,8))...)" -ForegroundColor Cyan
                        } else {
                            Write-Host "    Error: Section UUID also not found in cache" -ForegroundColor Red
                        }
                    }
                }
            } else {
                # No ParentPageUUID - link to section
                Write-Host "    No ParentPageUUID - linking to section" -ForegroundColor Gray

                if (![string]::IsNullOrWhiteSpace($pageData.SectionUUID)) {
                    if ($containerCache.ContainsKey($pageData.SectionUUID)) {
                        $parentNotionPageId = $containerCache[$pageData.SectionUUID]
                        Write-Host "    Linking to Section (UUID: $($pageData.SectionUUID.Substring(0,8))...)" -ForegroundColor Cyan
                    }
                }
            }
        }

        # Create new page
        $result = New-NotionPage -PageData $pageInfo `
                                 -ApiKey $notionConfig.ApiKey `
                                 -DatabaseId $notionConfig.DatabaseId `
                                 -PdfFilePath $pageData.FilePath `
                                 -MarkdownFilePath $pageData.MarkdownFilePath `
                                 -NotebookPageId $notebookPageId `
                                 -ParentNotionPageId $parentNotionPageId

        if ($result.Success) {
            $importedPages++
            # Only increment level counter if this is NOT a required parent imported beyond limit
            if (-not $isRequiredParent) {
                $pageLevelCounts[$currentLevel]++
            }
            Write-Host "    Created - Notion URL: $($result.Url)" -ForegroundColor Green

            # Cache the page ID by UUID for reliable parent lookups
            if (![string]::IsNullOrWhiteSpace($pageData.UUID)) {
                $pageUUIDToIdCache[$pageData.UUID] = $result.PageId
            }
        }
        else {
            $failedPages++
            Write-Host "    Failed: $($result.Error)" -ForegroundColor Red
        }

        if (-not $result.Success) {

            $failedPagesList += @{
                PageName = $pageData.PageName
                NotebookName = $notebook.DisplayName
                Section = $sectionName
                SectionGroup = $sectionGroupName
                Error = $result.Error
            }
        }

        $pageNumber++

        # Rate limiting - Notion allows 3 requests per second
        Start-Sleep -Milliseconds 350
    }

    # Break out of notebook loop if MaxPages limit reached
    if ($MaxPages -gt 0 -and $totalPages -ge $MaxPages) {
        break
    }
}

# Display summary
Write-Host "`n" + ("=" * 60) -ForegroundColor Green
Write-Host "IMPORT COMPLETE!" -ForegroundColor Green
Write-Host ("=" * 60) -ForegroundColor Green

Write-Host "`nSummary:" -ForegroundColor Cyan
Write-Host "  Notebooks processed: $($selectedNotebooks.Count)" -ForegroundColor White
Write-Host "  Total pages: $totalPages" -ForegroundColor White
Write-Host "  Successfully imported: $importedPages" -ForegroundColor Green
Write-Host "  Failed imports: $failedPages" -ForegroundColor $(if ($failedPages -gt 0) { "Red" } else { "Gray" })

# Display failed pages details if any
if ($failedPages -gt 0) {
    Write-Host "`nFailed Pages Details:" -ForegroundColor Red
    Write-Host ("=" * 60) -ForegroundColor Gray
    foreach ($failedPage in $failedPagesList) {
        $locationParts = @($failedPage.NotebookName)
        if (![string]::IsNullOrWhiteSpace($failedPage.SectionGroup)) {
            $locationParts += $failedPage.SectionGroup
        }
        $locationParts += $failedPage.Section
        $location = $locationParts -join ' -> '

        Write-Host "  Page: $($failedPage.PageName)" -ForegroundColor Yellow
        Write-Host "    Location: $location" -ForegroundColor Gray
        Write-Host "    Reason: $($failedPage.Error)" -ForegroundColor Red
        Write-Host ""
    }
}

# Stop transcript to finalize the log
Stop-Transcript | Out-Null

# End of script
Write-Host "`nComplete log saved to: $logPath" -ForegroundColor Green
Write-Host ""
Write-host "=== Import Complete === " -ForegroundColor Cyan
Start-Sleep -Seconds 2
