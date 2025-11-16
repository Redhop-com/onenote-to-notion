# Complete UUID-Based Sync System

## Overview

The OneNote to Notion export/import system now has **complete UUID-based synchronization** that prevents duplicate records and allows for safe re-exports and re-imports.

## How It Works End-to-End

### 1. First Export

```
Run-Export.ps1
├─ Generates NEW UUIDs for all entities
│  ├─ Section Groups: [GUID]
│  ├─ Sections: [GUID]
│  └─ Pages: [GUID]
├─ Saves to index.json with UUIDs
└─ Creates PDFs
```

**Result:** `index.json` contains UUIDs for all notebooks, section groups, sections, and pages.

### 2. First Import

```
Run-Import.ps1
├─ Reads index.json
├─ Uses existing Notion database (created from template) with Sync-UUID property
├─ For each entity:
│  ├─ Checks if Sync-UUID exists in Notion
│  ├─ If NOT found → Create new page with Sync-UUID
│  └─ If found → Skip (return existing page ID)
└─ Links parent-child relationships
```

**Result:** Notion database populated with pages, all having Sync-UUID values.

### 3. Re-Export (Same Data)

```
Run-Export.ps1 (again)
├─ Loads existing index.json
├─ Builds UUID lookup tables:
│  ├─ Section Groups: "NotebookName|GroupName" → UUID
│  ├─ Sections: "NotebookName|GroupName|SectionName" → UUID
│  └─ Pages: PageID → UUID
├─ For each entity:
│  ├─ Checks if key exists in lookup
│  ├─ If found → REUSE existing UUID
│  └─ If NOT found → Generate NEW UUID (new item)
└─ Overwrites index.json with preserved UUIDs
```

**Result:**
- Existing items keep their UUIDs
- New items get new UUIDs
- Console shows "Reusing existing UUID" or "Generated new UUID"

### 4. Re-Import (Same Data)

```
Run-Import.ps1 (again)
├─ Reads index.json (with preserved UUIDs)
├─ For each entity:
│  ├─ Checks if Sync-UUID exists in Notion
│  ├─ If found → Returns existing page ID (NO duplicate created)
│  └─ If NOT found → Creates new page
└─ Console shows "Already exists (UUID: ...)"
```

**Result:** NO duplicates created. Existing pages reused.

## Key Components

### Export Script (Export-OneNoteToPDF.ps1)

**UUID Preservation Logic:**

1. **Load Existing JSON** (lines 800-838)
   ```powershell
   if (Test-Path $jsonIndexPath) {
       $existingJson = Get-Content $jsonIndexPath -Raw | ConvertFrom-Json
       # Build lookup tables for UUIDs
   }
   ```

2. **Reuse or Generate UUIDs**
   - **Section Groups** (lines 858-867): Key = "NotebookName|SectionGroupName"
   - **Sections** (lines 892-899, 920-927): Key = "NotebookName|SectionGroupName|SectionName"
   - **Pages** (lines 338-343): Key = PageID (OneNote's internal ID)

**Why PageID for Pages?**
- PageID is stable even if page is renamed
- Page names can change, but PageID stays the same
- Ensures we track the same page across exports even after renames

### Import Script (Import-OneNoteToNotion.ps1)

**Duplicate Detection Logic:**

1. **Find-ExistingNotionPage** (lines 539-577)
   ```powershell
   function Find-ExistingNotionPage {
       param([string]$SyncUUID, ...)

       # Query Notion for pages with matching Sync-UUID
       $queryBody = @{
           "filter" = @{
               "property" = "Sync-UUID"
               "rich_text" = @{ "equals" = $SyncUUID }
           }
       }

       # Returns existing page or $null
   }
   ```

2. **New-NotionContainerPage** (lines 726-737)
   - Checks for existing sections/section groups
   - Returns existing page ID if found
   - Creates new page if not found

3. **New-NotionPage** (lines 838-851, 949-960)
   - Checks for existing pages
   - Returns existing page ID if found
   - Adds Sync-UUID property when creating
   - Creates new page if not found

### Database Schema (Notion Template)

**Sync-UUID Property**:
```
Property Name: Sync-UUID
Property Type: Rich Text
```

This property stores the UUID from the JSON export, enabling duplicate detection.

## UUID Key Formats

| Entity | Key Format | Example | Why This Key? |
|--------|-----------|---------|---------------|
| Section Group | `Notebook\|GroupName` | `"Work\|Projects"` | Name-based (assumes no renames) |
| Section | `Notebook\|GroupName\|SectionName` | `"Work\|Projects\|Alpha"` | Name-based (assumes no renames) |
| Section (no group) | `Notebook\|\|SectionName` | `"Work\|\|Notes"` | Empty string for missing group |
| Page | `PageID` | `"{ABC...}...{123}"` | Stable OneNote ID (handles renames) |

## Workflow Scenarios

### Scenario 1: No Changes

```
Export → UUIDs: {a, b, c}
Import → Creates 3 pages
Re-Export → UUIDs: {a, b, c} (same)
Re-Import → Finds existing pages, NO duplicates
```

✅ **Result:** Same 3 pages in Notion, no duplicates.

### Scenario 2: New Page Added

```
Export → UUIDs: {a, b, c}
Import → Creates 3 pages
[User adds page D in OneNote]
Re-Export → UUIDs: {a, b, c, d} (d is new)
Re-Import → a, b, c: Found (skipped)
           → d: Not found (created)
```

✅ **Result:** 4 pages total (3 existing + 1 new).

### Scenario 3: Page Renamed

```
Export → Page "Meeting" → UUID: a, PageID: {123}
Import → Creates page with Sync-UUID: a
[User renames "Meeting" to "Meeting Notes" in OneNote]
Re-Export → Page "Meeting Notes" → UUID: a (same, uses PageID)
Re-Import → Finds existing page with UUID: a
```

✅ **Result:** NO duplicate. Existing page reused despite name change.

### Scenario 4: Section Renamed (⚠️ Limitation)

```
Export → Section "Old Name" → UUID: a
Import → Creates page with Sync-UUID: a
[User renames section to "New Name" in OneNote]
Re-Export → Section "New Name" → UUID: b (NEW! Key doesn't match)
Re-Import → UUID b not found → Creates NEW page
```

❌ **Result:** Duplicate section created. Section renames not tracked.

**Workaround:** Manually delete old section in Notion, or don't rename sections.

## Benefits

✅ **Idempotent Exports** - Can re-export safely without losing sync
✅ **Idempotent Imports** - Can re-import safely without creating duplicates
✅ **Handles Page Renames** - Uses stable PageID for tracking
✅ **Clean Workflow** - Just export and import, no manual cleanup needed
✅ **Incremental Updates** - Only new items are created

## Limitations

❌ **Section/Section Group Renames** - Creates duplicates because key is name-based
❌ **No Update Logic** - Existing pages are skipped, not updated with new content
❌ **No Deletion Tracking** - Deleted OneNote pages remain in Notion

## Future Enhancements

1. **Section ID Tracking**: Use OneNote's internal section IDs instead of names
2. **Update Existing Pages**: Add logic to update content of existing pages
3. **Deletion Sync**: Track deleted pages and remove from Notion
4. **Conflict Resolution**: Handle cases where content has changed on both sides

## Testing the System

### Test 1: Basic Re-Import

```powershell
# First run
.\Run-Export.ps1
.\Run-Import.ps1
# Note the page count in Notion

# Second run (no changes)
.\Run-Import.ps1
# Check: Same page count (no duplicates)
```

### Test 2: New Page

```powershell
# Initial export/import
.\Run-Export.ps1
.\Run-Import.ps1

# Add a new page in OneNote
# Re-export/import
.\Run-Export.ps1
.\Run-Import.ps1

# Check: Only 1 new page added, no duplicates of existing pages
```

### Test 3: Page Rename

```powershell
# Initial export/import
.\Run-Export.ps1
.\Run-Import.ps1

# Rename a page in OneNote
# Re-export/import
.\Run-Export.ps1
.\Run-Import.ps1

# Check: No duplicate page created (old name should still exist with old name)
```

### Test 4: UUID Preservation

```powershell
# Export twice without changing OneNote
.\Run-Export.ps1
# Save the index.json somewhere
.\Run-Export.ps1
# Compare the two index.json files

# Check: UUIDs should be identical for all existing items
```

## Visual Output

When everything is working correctly:

**During Re-Export:**
```
Processing Notebook: Work
  Section Group: Projects
    Reusing existing UUID  ← Existing section group
    UUID: d54b3bbd-e640-...

    Section: Alpha
      Reusing existing UUID  ← Existing section

    Section: Beta
      Generated new UUID  ← New section added
```

**During Re-Import:**
```
Processing Notebook: Work
  Section Group: Projects
    Already exists (updating)  ← Found by UUID

    Section: Alpha
      Already exists (updating)  ← Found by UUID

    Processing pages...
      Page: Meeting Notes
        Already exists (UUID: abc-123)  ← Found by UUID

      Page: New Topic
        Creating page...  ← Not found, creating new
```

## Related Documentation

- **UUID_PRESERVATION.md** - Detailed explanation of UUID preservation in export
- **SYNC_UUID_IMPLEMENTATION.md** - Technical implementation details
- **UUID_MAPPING_SYSTEM.md** - Overall UUID strategy for parent linking
- **JSON_STRUCTURE_OPTIMIZATION.md** - JSON structure without redundant data
