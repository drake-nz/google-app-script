# Troubleshooting Formatting Issues

## Error: "Sheet is undefined or null"

This error means the script can't access your sheet. Here's how to fix it:

### Step 1: Run Diagnostic First

Before formatting, always run the diagnostic function:

1. In Apps Script editor, select `diagnoseSheet` from the function dropdown
2. Click Run (▶️)
3. Check the logs (View → Logs)

This will tell you:
- ✅ If the sheet opens successfully
- 📋 What sheet names are available
- 📊 Current row/column counts
- ❌ Any access errors

### Step 2: Common Issues and Fixes

#### Issue: Sheet is completely empty (no data at all)

**Solution:** Add at least a header row (row 1) with column names.

Example:
```
Column A: Name
Column B: Email
Column C: Date
```

Even if there's no data below, having headers allows formatting to work.

#### Issue: Sheet name doesn't match

**Symptoms:** Diagnostic shows available sheets, but formatting fails

**Solution:** Use the exact sheet name:
```javascript
formatMainSheet('Sheet1');  // Use exact name from diagnostic
```

#### Issue: Permissions error

**Symptoms:** Diagnostic shows "Could not open spreadsheet"

**Solution:**
1. Make sure you've authorized the script (click "Review Permissions")
2. Check that your Google account has access to the sheet
3. Verify the Sheet ID is correct in `Code.gs`

#### Issue: Sheet ID is wrong

**Symptoms:** "Could not open spreadsheet with ID: ..."

**Solution:**
1. Verify the Sheet ID in the URL: `https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit`
2. Update `CONFIG.SPREADSHEET_IDS.main` in `Code.gs`
3. Push the updated code: `clasp push`

### Step 3: Test with Minimal Data

If your sheet is empty, add this test data:

1. Open your Google Sheet
2. In row 1, add headers (e.g., "Name", "Email", "Date")
3. Optionally add 1-2 rows of test data
4. Run `testFormatMainSheet()` again

### Quick Test Function

Use this to test with a specific sheet name:

```javascript
// Format a specific sheet tab
formatMainSheet('Sheet1');  // Replace with your actual sheet name

// Or format the active sheet
formatSheet(CONFIG.SPREADSHEET_IDS.main);
```

## Still Having Issues?

1. **Check the logs** - They contain detailed error messages
2. **Run `diagnoseSheet()`** - This shows exactly what's accessible
3. **Verify Sheet ID** - Make sure it's correct in the CONFIG
4. **Check permissions** - Ensure you've authorized the script

## Expected Behavior

- ✅ Empty sheet with headers → Formats header row only
- ✅ Sheet with data → Formats headers + alternating row colors
- ✅ Multiple sheets → Specify sheet name to format specific one
