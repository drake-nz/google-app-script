# Automatic Update Setup

This guide explains how to set up automatic concatenation updates when source cells are edited in your Google Sheet.

## Overview

When someone edits a source cell (e.g., `B8`, `C8`), the concatenation will automatically run to update the destination cell (e.g., `D8`) with the new values and formatting.

## Two Types of Triggers

### 1. Simple Trigger (`onEdit`)
- **Automatically enabled** - No setup needed
- **Limitations:**
  - 6-minute execution time limit
  - Cannot use external APIs
  - May not work in all scenarios
- **Best for:** Simple use cases with quick operations

### 2. Installable Trigger (`onEditTrigger`)
- **More reliable** - Recommended for production
- **Advantages:**
  - 30-minute execution time limit
  - Can use external APIs
  - More reliable execution
- **Requires setup:** Run `setupAutoUpdateTrigger()` once

## Setup Instructions

### Option 1: Simple Trigger (Already Active)

The `onEdit()` function is already in the code and will work automatically. No setup needed!

**How it works:**
- When any cell is edited, it checks if that cell is a source cell in your `Format_Config`
- If yes, it runs `processAllConcatenations()` to update all concatenated cells
- If no, it does nothing (no performance impact)

### Option 2: Installable Trigger (Recommended)

For more reliability, especially in shared sheets:

1. **Open Apps Script Editor:**
   ```bash
   clasp open-script
   ```
   Or go to: Extensions → Apps Script

2. **Run the setup function:**
   - Select `setupAutoUpdateTrigger` from the function dropdown
   - Click Run (▶️)
   - Authorize permissions if prompted

3. **Verify it's set up:**
   - Go to: Triggers (clock icon) in the left sidebar
   - You should see: `onEditTrigger` → `On edit` → `Head`

4. **Test it:**
   - Edit a source cell (e.g., `B8` in "Performance and Progress")
   - The destination cell (e.g., `D8`) should update automatically
   - Check View → Logs to see the trigger running

## How It Works

1. **User edits a cell** (e.g., changes `B8` from "124" to "125")
2. **Trigger detects the edit** and checks if it's a source cell
3. **If it's a source cell:**
   - Loads all configurations from `Format_Config` sheet
   - Processes all concatenations (updates all destination cells)
   - Preserves formatting from source cells
4. **If it's not a source cell:** Does nothing (no performance impact)

## Which Cells Trigger Updates?

Only cells that are listed in the `SourceRanges` column of your `Format_Config` sheet will trigger updates.

**Example:**
- If `Format_Config` has: `B8, C8` → `D8`
- Editing `B8` or `C8` will trigger update
- Editing `D8` (destination) will NOT trigger update (prevents loops)
- Editing `A8` or `E8` will NOT trigger update

## Disabling Automatic Updates

If you want to disable automatic updates:

1. **Remove triggers:**
   - Run `removeAutoUpdateTriggers()` function
   - Or manually delete triggers in: Triggers → (trash icon)

2. **Manual updates only:**
   - Run `processAllConcatenations()` manually when needed
   - Or use `concatenateCellsWithFormatting('configKey')` for specific configs

## Troubleshooting

### "Trigger not running"
- **Check triggers:** Go to Triggers (clock icon) in Apps Script editor
- **Check permissions:** Make sure you authorized the script
- **Check logs:** View → Logs to see if trigger is firing

### "Updates too slow"
- Simple triggers have a 6-minute limit
- Use installable triggers for better performance
- Consider processing only affected configs (future enhancement)

### "Infinite loop / Too many updates"
- The script skips edits to `Format_Config` sheet
- The script only processes when source cells are edited
- If issues persist, disable triggers and use manual updates

### "Permission errors"
- Installable triggers require authorization
- Run `setupAutoUpdateTrigger()` and authorize when prompted
- Check that the script has edit access to the spreadsheet

## Performance Considerations

- **Simple trigger:** Processes ALL configurations on ANY source cell edit
- **Optimization:** Only processes when source cells are edited (not destination cells)
- **Future enhancement:** Could process only affected configurations (not all)

## Best Practices

1. **Use installable triggers** for shared sheets (more reliable)
2. **Test first** on a copy of your sheet
3. **Monitor logs** to ensure triggers are working correctly
4. **Keep Format_Config** organized and up-to-date
5. **Document** which cells trigger updates for your team

## Example Workflow

```
User edits B8: "124" → "125"
  ↓
onEdit() or onEditTrigger() detects edit
  ↓
Checks: Is B8 in any SourceRanges? YES
  ↓
Runs: processAllConcatenations()
  ↓
Updates: D8 = "125" + "M" = "125M" (with formatting)
  ↓
Done! ✅
```

## Advanced: Custom Trigger Logic

If you need custom logic (e.g., only update specific configs), you can modify `onEditTrigger()` to:

```javascript
function onEditTrigger(e) {
  const editedCell = e.range.getA1Notation();
  
  // Only update if editing specific cells
  if (editedCell === 'B8' || editedCell === 'C8') {
    concatenateCellsWithFormatting('perfD8');
  }
  // ... etc
}
```

---

**Note:** For shared sheets with multiple users, installable triggers are highly recommended for reliability.
