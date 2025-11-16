# OneNote to Notion Migration Tool

This PowerShell toolset exports all your OneNote notebooks to PDF files and imports them into Notion, maintaining the folder structure and hierarchy.

## Features

### Export Features
- Exports ALL OneNote notebooks automatically
- Creates one PDF per page with JSON metadata
- Maintains folder structure (Notebook → Section Group → Section → Page)
- Handles subpages with proper naming convention
- Generates JSON index files with page metadata (titles, dates, hierarchy)
- Progress tracking and detailed logging
- Exports Password protected pages, if unlocked. Note - they are exported unencrypted.
- Error handling with continuation on failures

### Import Features
- Uploads PDFs to Notion database with proper hierarchy
- Preserves parent-child relationships between pages
- Maintains creation and modification dates from OneNote
- Creates proper page structure in Notion
- UUID-based sync system for incremental updates
- Handles nested subpages (unlimited depth)
- Automatic database setup script included

## Requirements

### For Export
- Windows 10 or later
- OneNote Desktop application (NOT the Windows Store version)
  - Download: https://www.onenote.com/download
- Windows PowerShell 5.1 (comes with Windows)
- Sufficient disk space for PDFs

### For Import
- Windows 10 or later
- PowerShell 7.0 or higher (must install separately)
  - Download: https://aka.ms/powershell
- Notion API key (create at https://www.notion.so/my-integrations)
- Notion workspace with permission to create databases

## Installation

1. Download all essential files to the same folder:
   - **Export Scripts**: `Export-OneNoteToPDF.ps1`, `Run-Export.ps1`
   - **Import Scripts**: `Import-OneNoteToNotion.ps1`, `Run-Import.ps1`
   - **Testing**: `Test-OneNoteConnection.ps1`
   - **Optional Launchers**: `Export.cmd`, `Import.cmd`, `Test.cmd`

2. Install PowerShell 7 for import functionality:
   ```powershell
   winget install Microsoft.PowerShell
   ```

## Quick Start

### Step 1: Set Up Notion Database (First Time Only)

1. Create a Notion integration and get your API key:
   - Go to https://www.notion.so/my-integrations
   - Click "New integration"
   - Give it a name and select your workspace
   - Copy the API key

2. Use the Notion template to create your database:
   - Duplicate the provided Notion template to your workspace
   - Share the database with your integration
   - Copy the database ID from the URL

### Step 2: Export from OneNote

**Easy method** (double-click):
- Double-click `Export.cmd` or `Run-Export.ps1`

**Advanced method**:
```powershell
.\Export-OneNoteToPDF.ps1 -OutputPath "D:\MyExports" -IncludeSubpages $true -ShowProgress $true
```

This will create:
- PDF files for each page (one per page)
- JSON index file per notebook with metadata and UUIDs
- Export located at: `Documents\OneNoteExports\` (or your custom path)

### Step 3: Import to Notion

**Easy method** (double-click):
- Double-click `Import.cmd` or `Run-Import.ps1`

**Advanced method**:
```powershell
pwsh -File Import-OneNoteToNotion.ps1 -MaxPages 10  # Test with 10 pages first
```

The import script will:
- Read the exported PDFs and JSON files
- Upload to your Notion database
- Preserve dates and hierarchy
- Create parent-child relationships

## Usage

### Export Parameters

```powershell
.\Export-OneNoteToPDF.ps1 [parameters]
```

- `-OutputPath`: Specify custom export location (default: `Documents\OneNoteExports`)
- `-IncludeSubpages`: Include subpages in export (default: `$true`)
- `-ShowProgress`: Show detailed progress output (default: `$true`)

### Import Parameters

```powershell
pwsh -File Import-OneNoteToNotion.ps1 [parameters]
```

- `-NotionApiKey`: Your Notion API key (optional, will prompt if not provided)
- `-NotionDatabaseId`: Your Notion database ID (optional, will prompt if not provided)
- `-MaxPages`: Limit import to N pages (useful for testing)
- `-ExportPath`: Path to the exported files (default: `Documents\OneNoteExports`)

### Testing Connection

Before exporting, test your OneNote connection:
```powershell
.\Test-OneNoteConnection.ps1
```

This will:
- Verify OneNote Desktop is installed
- Check if notebooks are accessible
- Count total pages to export
- Estimate time and disk space needed

## Output Structure

### Export Output

```
OneNoteExports/
├── NotebookName1/
│   ├── SectionGroupName/          # Section groups create folders
│   │   ├── SectionName/
│   │   │   ├── Page1.pdf
│   │   │   ├── Page2.pdf
│   │   │   └── Page2_Subpage1.pdf  # Subpages use underscore naming
│   │   └── SectionName2/
│   │       └── TopLevelPage_ChildPage_GrandchildPage.pdf
│   ├── DirectSectionName/         # Sections not in groups
│   │   ├── Page1.pdf
│   │   └── ParentPage_SubPage.pdf
│   └── index.json                 # Notebook metadata with UUIDs
└── export_log_YYYYMMDD_HHMMSS.txt
```

### Notion Database Structure

The import creates entries in two Notion databases:

#### Notebooks Database
- **Name** (Title): Notebook name

#### Pages Database
The import creates pages with these properties:
- **Name** (Title): Page name
- **Label** (Multi-select): Contains the immediate parent Section name
- **Notebook** (Relation): Links to entry in Notebooks database
- **Type** (Select): Only set for containers (Section Groups, Sections) - not for regular pages
- **Imported** (Date): OneNote creation date (from CreatedTime)
- **Sync-UUID** (Rich Text): Unique identifier for incremental sync
- **Parent item** (Relation): Links to parent page for hierarchy
- **Last Edited Time** (Automatic): Notion's automatic last modified timestamp
- **PDF** (Files): Attached PDF file of the page content

## Workflow Examples

### Full Migration

1. **Set up Notion database** (first time only):
   - Use the Notion template to create your database
   - Get your API key and database ID

2. **Export from OneNote**:
   ```powershell
   .\Run-Export.ps1
   ```

3. **Test import with a few pages**:
   ```powershell
   pwsh -File Import-OneNoteToNotion.ps1 -MaxPages 10
   ```

4. **Import everything**:
   ```powershell
   pwsh -File Import-OneNoteToNotion.ps1
   ```

### Incremental Updates

The system supports incremental updates via UUID-based sync:

1. Make changes in OneNote
2. Re-run export: `.\Run-Export.ps1`
3. Re-run import: `pwsh -File Import-OneNoteToNotion.ps1`

The import will:
- Skip unchanged pages
- Update modified pages
- Add new pages
- Preserve existing Notion page IDs

## Troubleshooting

### Export Issues

**Error: "Cannot find OneNote.Application"**
- Install OneNote Desktop (not Windows Store version)
- Download: https://www.onenote.com/download

**Error: "Execution Policy"**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**Error: "COM component error when using PowerShell 7"**
- Use Windows PowerShell 5.1 for export (use `Export.cmd` launcher)
- Export requires the Desktop edition, not Core

### Import Issues

**Error: "PowerShell 7 is not installed"**
- Download and install from: https://aka.ms/powershell
- Or use: `winget install Microsoft.PowerShell`

**Error: "Invalid API key"**
- Verify your API key at https://www.notion.so/my-integrations
- Ensure the integration has access to your workspace

**Error: "Database not found"**
- Ensure you've shared the database with your integration
- Verify the database ID is correct

**Error: "Missing required properties"**
- The database schema is incorrect
- Use the Notion template which has all required properties
- Or manually add the required properties (see error message for details)

**Import is slow**
- Notion API has rate limits
- Large PDFs take longer to upload
- Consider using `-MaxPages` to test with smaller batches first

### General Issues

**Script runs but no output**
- Check the export_log_*.txt file for errors
- Ensure OneNote is properly installed and synced
- The script shows detailed progress by default with `-ShowProgress $true`

**Some pages fail**
- Pages with special characters may have issues
- Very large pages (>100MB) may timeout
- Check the log file for specific errors

## File Naming

The export script sanitizes file names by:
- Replacing invalid characters (: / \ * ? " < > |) with underscores
- Limiting length to 100 characters
- Removing leading/trailing spaces and dots
- Preserving the original names in JSON metadata

## Sync System

The toolset uses UUID-based synchronization:
- Each OneNote item (notebook, section, page) gets a unique UUID
- UUIDs are preserved in the index.json file across exports
- UUIDs are stored in Notion's "Sync-UUID" property
- Subsequent imports use UUIDs to detect and update existing pages
- This enables incremental updates without duplicating pages
- See `COMPLETE_SYNC_SYSTEM.md` for detailed documentation

## Log Files

### Export Log (`export_log_YYYYMMDD_HHMMSS.txt`)
- Complete transcript of export session
- Export summary with counts (notebooks, sections, pages)
- Number of successful/failed exports
- Timestamp and duration
- Located in the export output folder

### Import Log (console output)
- Real-time progress display
- Pages imported/updated/skipped counts
- Upload progress with time estimates
- Errors and warnings
- Final summary statistics

## Advanced Topics

### Custom Database Setup
If you want to use an existing Notion database:
1. Ensure it has all required properties (see Notion template for reference)
2. Share the database with your integration
3. Use the database ID when importing

### Filtering Notebooks
To export specific notebooks only, modify `Export-OneNoteToPDF.ps1` line 148:
```powershell
foreach ($notebook in $xml.Notebooks.Notebook | Where-Object {$_.name -match "Work"}) {
```

### Batch Processing
For very large imports, use `-MaxPages` to process in batches:
```powershell
pwsh -File Import-OneNoteToNotion.ps1 -MaxPages 100
```

## Performance

- **Export**: ~1-3 seconds per page (depends on page size)
- **Import**: ~2-5 seconds per page (depends on PDF size and Notion API)
- Large notebooks (500+ pages) may take 30+ minutes per operation

## Support

For detailed documentation, see:
- `START_HERE.txt` - Quick start guide
- `QUICK_START.md` - Detailed setup instructions
- `COMPLETE_SYNC_SYSTEM.md` - How the sync system works
- `PARENT_PAGE_UUID_LINKING.md` - How parent-child linking works

## License

This tool is provided as-is for personal use.

## Version

Version 2.0
- Full export and import functionality
- UUID-based sync system
- Parent-child hierarchy preservation
- Incremental update support
- Automated database setup
