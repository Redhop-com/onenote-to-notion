# Quick Start Guide

This guide shows you the easiest way to run the OneNote export and import scripts.

## TL;DR - Just Double-Click These Files

### For Exporting OneNote to PDF
**Double-click:** `Export.cmd`

### For Importing to Notion
**Double-click:** `Import.cmd`

### For Testing OneNote Connection
**Double-click:** `Test.cmd`

That's it! The `.cmd` launcher files automatically use the correct PowerShell version.

---

## Why Do We Need Different PowerShell Versions?

### Export Scripts → Windows PowerShell 5.1
- **Export.cmd** launches: `Export-OneNoteToPDF.ps1`
- **Test.cmd** launches: `Test-OneNoteConnection.ps1`
- **Why?** OneNote COM automation only works in Windows PowerShell 5.1 (Desktop edition)
- Uses: `powershell.exe`

### Import Script → PowerShell 7
- **Import.cmd** launches: `Import-OneNoteToNotion.ps1`
- **Why?** Uses modern .NET HTTP APIs for Notion file uploads
- Uses: `pwsh.exe`
- **Note:** You must install PowerShell 7 from https://aka.ms/powershell

---

## File Overview

### Launcher Files (Easy - Just Double-Click!)
```
Export.cmd         → Export OneNote notebooks to PDF
Import.cmd         → Import PDFs to Notion database
Test.cmd           → Test OneNote connection
```

### PowerShell Scripts (Advanced - Use if you know PowerShell)
```
Export-OneNoteToPDF.ps1          → Main export script (requires PS 5.1)
Import-OneNoteToNotion.ps1       → Main import script (requires PS 7)
Test-OneNoteConnection.ps1       → Connection tester (requires PS 5.1)
Run-Export.ps1                   → Alternative export launcher (requires PS 5.1)
```

---

## Troubleshooting

### "PowerShell 7 is not installed" Error
If you see this when running `Import.cmd`:

1. **Download PowerShell 7:**
   - Visit: https://aka.ms/powershell
   - Or run: `winget install Microsoft.PowerShell`

2. **After installing**, restart your terminal and try again

### "Script cannot be loaded" Error
If you see an execution policy error:

**Option 1 (Recommended):** Use the `.cmd` launchers - they handle this automatically

**Option 2:** Enable script execution:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Export Script Shows "COM component error"
You're probably using PowerShell 7 instead of Windows PowerShell 5.1.

**Solution:** Use `Export.cmd` instead of running the script directly.

### How to Check Your PowerShell Version
```powershell
$PSVersionTable
```

**Windows PowerShell 5.1 looks like:**
```
PSVersion      : 5.1.xxxxx
PSEdition      : Desktop
```

**PowerShell 7 looks like:**
```
PSVersion      : 7.x.x
PSEdition      : Core
```

---

## Requirements

### For Export Scripts
- ✅ Windows 10 or later
- ✅ Windows PowerShell 5.1 (comes with Windows)
- ✅ OneNote Desktop (NOT Windows Store version)
  - Download: https://www.onenote.com/download

### For Import Script
- ✅ Windows 10 or later
- ✅ PowerShell 7.0 or higher
  - Download: https://aka.ms/powershell
- ✅ Notion API key and database
- ✅ Exported OneNote files (from running Export.cmd first)

---

## Workflow

### Complete OneNote → Notion Migration

1. **Export from OneNote**
   ```
   Double-click: Export.cmd
   ```
   - Exports all notebooks to: `Documents\OneNoteExports\`
   - Creates PDF files for each page
   - Generates JSON index files with metadata

2. **Import to Notion**
   ```
   Double-click: Import.cmd
   ```
   - Reads the exported PDFs and JSON files
   - Uploads to your Notion database
   - Preserves dates, hierarchy, and metadata

---

## Advanced Usage

### Running with Custom Parameters

#### Export with Custom Path
```cmd
powershell.exe -File Export-OneNoteToPDF.ps1 -OutputPath "D:\MyExports"
```

#### Export without Subpages
```cmd
powershell.exe -File Export-OneNoteToPDF.ps1 -IncludeSubpages $false
```

#### Import with API Key as Parameter
```cmd
pwsh.exe -File Import-OneNoteToNotion.ps1 -NotionApiKey "your_key" -NotionDatabaseId "your_db_id"
```

#### Limit Import to First 10 Pages (for testing)
```cmd
pwsh.exe -File Import-OneNoteToNotion.ps1 -MaxPages 10
```

---

## Getting Help

### Test Your OneNote Connection
Before exporting, test if OneNote is accessible:
```
Double-click: Test.cmd
```

This will:
- ✅ Verify OneNote Desktop is installed
- ✅ Check if notebooks are accessible
- ✅ Count total pages to export
- ✅ Estimate time and disk space needed

### Common Issues
See `POWERSHELL_VERSION_FIX.md` for detailed troubleshooting

---

## Summary

**Simplest approach:**
1. Export: Double-click `Export.cmd`
2. Import: Double-click `Import.cmd`

The launcher files handle all the complexity of using the correct PowerShell version!
