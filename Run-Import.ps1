# Quick Launcher for Import-OneNoteToNotion.ps1
# Simplified script to run the import with default settings

Clear-Host
Write-Host @"
==========================================================
        Quick Import to Notion Launcher

  This will import your exported OneNote notebooks
  into a Notion database
==========================================================
"@ -ForegroundColor Green

Write-Host ""

# Run the main import script with default parameters
& "$PSScriptRoot\Import-OneNoteToNotion.ps1"
