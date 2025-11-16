# Parent Page UUID-Based Linking

## Overview

The export/import system now uses **UUIDs instead of page names** for parent-child page relationships. This eliminates lookup conflicts with deeply nested pages and provides more reliable parent linking.

## The Problem (Before)

### Export Script Behavior:
For deeply nested pages, the export script built hierarchical parent names:
- Level 1: Page "Notes" → ParentPage = ""
- Level 2: Page "Subpage 1" → ParentPage = "Notes"
- Level 3: Page "Subpage 2" → ParentPage = "**Notes_Subpage 1**" (hierarchical!)
- Level 4: Page "Subpage 3" → ParentPage = "**Notes_Subpage 1_Subpage 2**"

### Import Script Behavior:
The import script cached pages by their simple `PageName`:
```powershell
$pageNameToIdCache["Notes"] = page-id-1
$pageNameToIdCache["Subpage 1"] = page-id-2
$pageNameToIdCache["Subpage 2"] = page-id-3
```

### The Conflict:
- Level 3 page looks for parent "Notes_Subpage 1"
- Cache only has "Subpage 1"
- ❌ **Lookup fails** → Parent link not created → Import fails

## The Solution (Now)

### Use UUIDs for Parent References

Instead of using hierarchical names, we now use the **immediate parent's UUID**.

### Export Script Changes:

1. **Track Parent UUIDs by Level** (Export-OneNoteToPDF.ps1:694-695):
```powershell
$parentUUIDsByLevel = @{}
```

2. **Get Immediate Parent's UUID** (lines 722-726):
```powershell
# Get the immediate parent's UUID (one level up)
$immediateParentLevel = $currentPageLevel - 1
if ($parentUUIDsByLevel.ContainsKey($immediateParentLevel)) {
    $parentPageUUID = $parentUUIDsByLevel[$immediateParentLevel]
}
```

3. **Store UUID After Export** (lines 745-746):
```powershell
$parentUUIDsByLevel[$currentPageLevel] = $result.PageUUID
```

4. **Pass to Export Function** (line 735):
```powershell
-ParentPageUUID $parentPageUUID
```

5. **Return in Metadata** (line 357):
```powershell
ParentPageUUID = $ParentPageUUID  # UUID of immediate parent page
```

6. **Include in JSON** (line 985):
```powershell
ParentPageUUID = $_.ParentPageUUID  # UUID of immediate parent (reliable linking)
```

### Import Script Changes:

1. **Cache by UUID** (Import-OneNoteToNotion.ps1:1375):
```powershell
$pageUUIDToIdCache = @{}  # Was: $pageNameToIdCache
```

2. **Lookup by ParentPageUUID** (lines 1625-1631):
```powershell
if (![string]::IsNullOrWhiteSpace($pageData.ParentPageUUID)) {
    if ($pageUUIDToIdCache.ContainsKey($pageData.ParentPageUUID)) {
        $parentNotionPageId = $pageUUIDToIdCache[$pageData.ParentPageUUID]
        Write-Host "    Parent: Page UUID '$($pageData.ParentPageUUID.Substring(0,8))...'"
    }
}
```

3. **Cache After Create** (lines 1654-1656):
```powershell
if (![string]::IsNullOrWhiteSpace($pageData.PageUUID)) {
    $pageUUIDToIdCache[$pageData.PageUUID] = $result.PageId
}
```

4. **Update Required Parents Check** (line 1525):
```powershell
$requiredParents[$pageData.ParentPageUUID] = $true  # Was: ParentPage
```

## JSON Structure

### Before:
```json
{
  "PageName": "Subpage 2",
  "PageUUID": "222dca40-4530-4f1a-a233-4d3c3e773e6d",
  "ParentPage": "Notes_Subpage 1",  ← Hierarchical name (problematic)
  "PageLevel": 3
}
```

### After:
```json
{
  "PageName": "Subpage 2",
  "PageUUID": "222dca40-4530-4f1a-a233-4d3c3e773e6d",
  "ParentPageUUID": "a75c8ce6-0a53-4453-a745-6c1a69e31538",  ← NEW! Reliable UUID
  "PageLevel": 3
}
```

Note: `ParentPage` has been removed from the JSON since `ParentPageUUID` provides reliable linking. The parent name hierarchy is still used internally for filename construction (e.g., "Parent_Child.pdf") but is not exported to JSON.

## Example Hierarchy

### OneNote Structure:
```
Notes (Level 1)
└─ Subpage 1 (Level 2)
   ├─ Subpage 2 (Level 3)
   └─ Subpage 3 (Level 3)
```

### Export Data:
```json
[
  {
    "PageName": "Notes",
    "PageUUID": "f01e924b-e4c9-48bd-9323-3fe4fc570410",
    "ParentPageUUID": "",
    "PageLevel": 1
  },
  {
    "PageName": "Subpage 1",
    "PageUUID": "a75c8ce6-0a53-4453-a745-6c1a69e31538",
    "ParentPageUUID": "f01e924b-e4c9-48bd-9323-3fe4fc570410",  ← Points to Notes
    "PageLevel": 2
  },
  {
    "PageName": "Subpage 2",
    "PageUUID": "222dca40-4530-4f1a-a233-4d3c3e773e6d",
    "ParentPageUUID": "a75c8ce6-0a53-4453-a745-6c1a69e31538",  ← Points to Subpage 1
    "PageLevel": 3
  },
  {
    "PageName": "Subpage 3",
    "PageUUID": "c60d6a41-a6ff-49bd-814d-97f1f917fdb9",
    "ParentPageUUID": "a75c8ce6-0a53-4453-a745-6c1a69e31538",  ← Points to Subpage 1
    "PageLevel": 3
  }
]
```

### Import Cache (After Level 2 Imported):
```powershell
$pageUUIDToIdCache = @{
    "f01e924b-e4c9-48bd-9323-3fe4fc570410" = "notion-page-id-1"  # Notes
    "a75c8ce6-0a53-4453-a745-6c1a69e31538" = "notion-page-id-2"  # Subpage 1
}
```

### Parent Lookup for Level 3:
```powershell
# Importing Subpage 2
$parentPageUUID = "a75c8ce6-0a53-4453-a745-6c1a69e31538"
$parentNotionPageId = $pageUUIDToIdCache[$parentPageUUID]  # ✅ Found!
# Result: "notion-page-id-2" (Subpage 1's Notion ID)
```

## Benefits

✅ **Reliable Parent Linking** - UUIDs never change, unlike names
✅ **No Hierarchical Name Conflicts** - Direct UUID lookup instead of complex name building
✅ **Simpler Cache Logic** - Just one UUID → ID mapping, no hierarchical names needed
✅ **Handles Deep Nesting** - Works for any depth of page hierarchy
✅ **Consistent with Sync System** - Uses same UUID strategy as duplicate prevention

## Simplified JSON Structure

The `ParentPage` field has been **removed from the JSON export** because:
- ✅ `ParentPageUUID` provides reliable, unambiguous parent linking
- ✅ Eliminates redundancy (no need for both name and UUID)
- ✅ Reduces JSON file size
- ✅ Prevents confusion about which field is authoritative

**Note:** The parent name hierarchy is still used **internally** during export for filename construction. Files are still named with hierarchical patterns like:
- `Notes.pdf` (Level 1)
- `Notes_Subpage 1.pdf` (Level 2)
- `Notes_Subpage 1_Subpage 2.pdf` (Level 3)

This filename pattern helps identify the page hierarchy when browsing the exported files, but the JSON only stores the UUID reference.

## Backward Compatibility

- **New exports** include only `ParentPageUUID`
- **Old exports** (with `ParentPage` field) will still work but the field is ignored during import

## Testing

To verify the fix works:

1. **Export a deeply nested structure:**
   ```
   Page A (Level 1)
   └─ Page B (Level 2)
      └─ Page C (Level 3)
         └─ Page D (Level 4)
   ```

2. **Check JSON has ParentPageUUID:**
   ```powershell
   Get-Content ".\OneNoteExports\Notebook\index.json" | ConvertFrom-Json |
     Select -ExpandProperty Pages |
     Where PageLevel -eq 3 |
     Select PageName, ParentPageUUID
   ```

3. **Import and verify:**
   - No "Warning: Parent page UUID not found" messages
   - All pages create successfully
   - Parent relations show correctly in Notion

## Files Modified

### Export Script: Export-OneNoteToPDF.ps1
- Line 108: Add `ParentPageUUID` parameter to Export-PageToPDF function
- Lines 694-695: Track parent UUIDs by level
- Lines 722-726: Get immediate parent's UUID
- Line 735: Pass ParentPageUUID to Export-PageToPDF
- Lines 745-746: Store current page's UUID for children
- Line 356: Return ParentPageUUID in metadata (removed ParentPage from export)
- Line 983: Include ParentPageUUID in JSON (removed ParentPage from JSON)

### Import Script: Import-OneNoteToNotion.ps1
- Line 1375: Rename cache to $pageUUIDToIdCache
- Lines 1625-1634: Lookup parent by ParentPageUUID
- Lines 1654-1656: Cache by PageUUID
- Line 1525: Use ParentPageUUID in required parents check
- Line 1549: Check required parents by UUID
- Line 1591: Removed ParentPage from $pageInfo hashtable

## Related Documentation

- **UUID_PRESERVATION.md** - How UUIDs are preserved across exports
- **SYNC_UUID_IMPLEMENTATION.md** - Overall UUID-based sync system
- **COMPLETE_SYNC_SYSTEM.md** - End-to-end sync workflow
