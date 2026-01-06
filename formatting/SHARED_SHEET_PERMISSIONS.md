# Shared Sheet Permissions & Triggers

This guide explains how automatic updates work in shared Google Sheets and what permissions are needed.

## How Triggers Work in Shared Sheets

### Installable Triggers (Recommended)

**Key Point:** Installable triggers run with the **permissions of the person who installed the trigger**, not the person editing the sheet.

**What this means:**
- ✅ **Anyone can edit the sheet** - The trigger will work for all editors
- ✅ **No authentication needed** - Editors don't need to run the script themselves
- ✅ **Consistent behavior** - All edits trigger updates the same way
- ⚠️ **Setup required** - One person (sheet owner/admin) must set up the trigger once

### Simple Triggers (onEdit)

**Key Point:** Simple triggers run with the **permissions of the person editing the sheet**.

**What this means:**
- ⚠️ **Each user needs permissions** - Each editor needs authorization to run the script
- ⚠️ **First-time authorization** - Each new editor will be prompted to authorize the script
- ⚠️ **May not work reliably** - Some users may not authorize, breaking the trigger
- ✅ **No setup needed** - Works automatically (if authorized)

## Recommended Setup for Shared Sheets

### Step 1: Sheet Owner Sets Up Installable Trigger

**Who should do this:**
- Sheet owner or administrator
- Someone with full edit access to the sheet

**Steps:**
1. Open Apps Script editor (Extensions → Apps Script)
2. Run `setupAutoUpdateTrigger()` function
3. Authorize permissions when prompted
4. Done! ✅

**Result:**
- Trigger is now active for ALL users
- Any editor can update source cells and concatenation will run automatically
- No additional setup needed for other users

### Step 2: Share Sheet with Editors

**Permissions needed:**
- **Editors** need "Editor" access to edit cells
- **Viewers** can view but won't trigger updates (they can't edit)

**How to share:**
1. Click "Share" button in Google Sheets
2. Add users with "Editor" role
3. They can now edit and trigger automatic updates

## What Happens When Someone Edits

### Scenario: User A (Editor) edits cell B8

1. **User A edits** `B8` from "124" to "125"
2. **Trigger fires** (runs with sheet owner's permissions)
3. **Script runs** `processAllConcatenations()`
4. **Cell D8 updates** automatically to "125M" with formatting
5. **User A sees** the update immediately

**User A does NOT need to:**
- ❌ Run any scripts
- ❌ Authorize anything
- ❌ Have special permissions
- ❌ Do anything except edit the cell

## Permission Requirements

### For Sheet Owner/Admin (Sets Up Trigger)

**Needs:**
- ✅ Owner or Editor access to the sheet
- ✅ Permission to create Apps Script triggers
- ✅ Authorization to run the script (one-time)

**To set up:**
```javascript
// Run this ONCE as sheet owner/admin
setupAutoUpdateTrigger()
```

### For Sheet Editors (Regular Users)

**Needs:**
- ✅ Editor access to the sheet (to edit cells)
- ❌ NO script authorization needed
- ❌ NO special permissions needed
- ❌ NO setup required

**They just:**
- Edit cells normally
- See automatic updates happen

## Troubleshooting

### "Trigger not working for some users"

**If using simple trigger (`onEdit`):**
- Each user must authorize the script
- Some users may not authorize (privacy concerns)
- **Solution:** Use installable trigger instead

**If using installable trigger:**
- Check that trigger was set up by sheet owner
- Verify trigger exists: Apps Script → Triggers (clock icon)
- Check that users have Editor access (not just Viewer)

### "Permission denied errors"

**Possible causes:**
1. **Trigger not set up** - Run `setupAutoUpdateTrigger()`
2. **Wrong user set up trigger** - Should be sheet owner/admin
3. **Script needs re-authorization** - Owner may need to re-authorize

**Solution:**
- Sheet owner should run `setupAutoUpdateTrigger()` again
- Re-authorize if prompted

### "Users see authorization prompt"

**If using simple trigger:**
- This is normal - each user must authorize
- **Solution:** Switch to installable trigger

**If using installable trigger:**
- This shouldn't happen
- Check that trigger is properly installed
- Verify it's an installable trigger (not simple trigger)

## Best Practices

1. **Use installable triggers** for shared sheets
2. **Sheet owner sets up trigger** (not individual users)
3. **Document who set it up** for troubleshooting
4. **Test with multiple users** before going live
5. **Monitor logs** to ensure triggers are working

## Security Considerations

### What Permissions Does the Script Need?

The script needs:
- ✅ **Read access** to the spreadsheet (to read source cells)
- ✅ **Write access** to the spreadsheet (to update destination cells)
- ✅ **Access to Format_Config sheet** (to read configurations)

### Who Can See the Script?

- **Sheet owners/admins** can see and edit the script
- **Editors** typically cannot see the script (unless explicitly shared)
- **Script runs** with owner's permissions (not visible to editors)

### Data Privacy

- Script only accesses the specific spreadsheet
- No external data is sent
- All processing happens within Google's servers
- Editors don't need to grant permissions

## Example: Multi-User Workflow

```
Day 1: Setup
├── Sheet Owner: Sets up installable trigger
├── Sheet Owner: Shares sheet with Team (Editor access)
└── ✅ Setup complete

Day 2: Team Uses Sheet
├── User A: Edits B8 → D8 updates automatically ✅
├── User B: Edits C9 → D9 updates automatically ✅
├── User C: Edits E6 → G6 updates automatically ✅
└── All working without any user setup! ✅
```

## Summary

**For shared sheets:**
- ✅ **Installable trigger** = Works for everyone, no user setup needed
- ⚠️ **Simple trigger** = Each user must authorize (not recommended)

**Setup:**
- One person (owner/admin) runs `setupAutoUpdateTrigger()` once
- All other users can edit and see automatic updates
- No additional authentication needed for editors

---

**Bottom line:** With installable triggers, anyone with Editor access can edit the sheet and trigger automatic updates, without needing to authenticate or set up anything themselves.
