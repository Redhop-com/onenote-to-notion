# Quick Launch Script for OneNote to PDF Export
# Run this to export with default settings

#Requires -PSEdition Desktop

# Set the output path (change this to your preferred location)
$exportPath = "$env:USERPROFILE\Documents\OneNoteExports"

# Display banner
Clear-Host
Write-Host @"
==========================================================
          OneNote to PDF Batch Export Tool

  Export your OneNote notebooks to PDF files
  Each page will be saved as an individual PDF file
==========================================================
"@ -ForegroundColor Cyan

Write-Host "`nExport Settings:" -ForegroundColor Yellow
Write-Host "  - Output folder: $exportPath" -ForegroundColor White
Write-Host "  - Include subpages: Yes" -ForegroundColor White
Write-Host "  - Naming format: NotebookName_SectionName_PageName.pdf" -ForegroundColor White
Write-Host "  - Folder structure: Notebook/SectionGroup/Section/" -ForegroundColor White
Write-Host "  - Interactive selection: Choose which notebooks to export" -ForegroundColor White

Write-Host "`nIMPORTANT:" -ForegroundColor Red
Write-Host "  - Requires Windows PowerShell 5.1 (NOT PowerShell Core 7.x)" -ForegroundColor White
Write-Host "  - Make sure OneNote Desktop is installed (NOT Windows Store version)" -ForegroundColor White
Write-Host "  - This may take a while depending on the number of notebooks" -ForegroundColor White
Write-Host "  - Do not close this window until the export is complete" -ForegroundColor White

Write-Host "`nPress Enter to start the export or Ctrl+C to cancel..." -ForegroundColor Green
Read-Host

# Run the main export script
$scriptPath = Join-Path $PSScriptRoot "Export-OneNoteToPDF.ps1"

if (Test-Path $scriptPath) {
    & $scriptPath -OutputPath $exportPath -IncludeSubpages $true -ShowProgress $true
} else {
    Write-Host "Error: Cannot find Export-OneNoteToPDF.ps1" -ForegroundColor Red
    Write-Host "Make sure both scripts are in the same folder." -ForegroundColor Red
    Read-Host "Press Enter to exit"
}
