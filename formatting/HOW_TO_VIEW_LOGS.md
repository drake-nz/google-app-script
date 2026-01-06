# How to View Logs in Google Apps Script

There are two different places to see logs - here's the difference:

## Option 1: View → Logs (Detailed Logger.log() Output) ✅

This is where you see the detailed logs from `Logger.log()` statements:

1. **In the Apps Script editor**, go to **View** → **Logs**
   - Or press `Cmd+Enter` (Mac) / `Ctrl+Enter` (Windows)
   - Or click the "View" menu at the top, then "Logs"

2. **A log viewer window will open** showing:
   - All `Logger.log()` messages
   - Timestamps
   - Detailed execution information
   - Error messages and stack traces

**This is what you want to see** - it shows all the diagnostic info and formatting progress!

## Option 2: Execution Log (Just Execution History)

The **Execution log** (sometimes shown in a sidebar) only shows:
- When functions ran
- Success/failure status
- Execution time
- But NOT the detailed `Logger.log()` messages

**This is NOT what you want** - it's just a history of runs.

## Quick Way to See Logs

**Method 1: Menu**
- Click **View** → **Logs** in the Apps Script editor

**Method 2: Keyboard Shortcut**
- Press `Cmd+Enter` (Mac) or `Ctrl+Enter` (Windows) after running a function

**Method 3: Terminal (Alternative)**
```bash
cd /Users/dot/Scripts/apps-script-projects/formatting
clasp logs
```

## What You Should See

When you run `testFormatMainSheet()`, you should see logs like:

```
=== Starting Format Test ===
Sheet ID: YOUR_SHEET_ID_HERE
Sheet Name: Test1
Running diagnostic...
✅ Spreadsheet opened: [Your Sheet Name]
Available sheets: Test1, ...
Active sheet: Test1
Last row: X
Last column: Y
Starting format for sheet ID: ...
Opened sheet: Test1 (ID: ...)
Sheet has X rows and Y columns
Formatted header row 1 (Y columns)
Formatted data rows 2 to X
Auto-resized Y columns
Added borders to data range
✅ Formatting complete!
```

## Troubleshooting

**If you don't see logs:**
1. Make sure you ran a function (clicked Run ▶️)
2. Check that the function actually executed (no errors before it started)
3. Try View → Logs again
4. Or use `clasp logs` in terminal

**If logs are empty:**
- The function might not have run
- Check for errors in the execution log
- Make sure you selected the function before clicking Run

## Pro Tip

Keep the Logs window open while testing - it updates in real-time as your script runs!
