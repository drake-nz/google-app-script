# How to Find Your Google Sheet ID

The Sheet ID is a unique identifier for your Google Sheet. Here are several ways to find it:

## Method 1: From the URL (Easiest)

1. Open your Google Sheet in a browser
2. Look at the URL in the address bar

The URL format is:
```
https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit
```

**Example:**
If your URL is:
```
https://docs.google.com/spreadsheets/d/1abc123def456ghi789jkl012mno345pqr/edit
```

Then your Sheet ID is: `1abc123def456ghi789jkl012mno345pqr`

### Visual Guide:
```
https://docs.google.com/spreadsheets/d/[THIS_IS_YOUR_SHEET_ID]/edit
                                    ^^^^^^^^^^^^^^^^^^^^^^^^
```

## Method 2: Copy Link Method

1. Open your Google Sheet
2. Click **File** → **Share** → **Copy link** (or click the "Share" button)
3. The copied link will contain the Sheet ID in the same format as above

## Method 3: From Apps Script Editor

1. Open your Google Sheet
2. Go to **Extensions** → **Apps Script**
3. In the Apps Script editor, the URL will show:
   ```
   https://script.google.com/home/projects/<SCRIPT_ID>/edit
   ```
4. But to get the Sheet ID, look at the original sheet URL (Method 1)

## Method 4: Using Apps Script Code

If you're already in an Apps Script bound to a sheet:

```javascript
function getSheetId() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetId = sheet.getId();
  Logger.log('Sheet ID: ' + sheetId);
  return sheetId;
}
```

Run this function and check the logs (`clasp logs` or View → Logs).

## Method 5: From Share Settings

1. Open your Google Sheet
2. Click the **Share** button (top right)
3. In the share dialog, the URL shown contains the Sheet ID

## Quick Reference

| Location | Format |
|---------|--------|
| URL | `https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit` |
| Share link | Same as URL |
| Apps Script | Use `SpreadsheetApp.getActiveSpreadsheet().getId()` |

## Adding to Your Config

Once you have your Sheet ID, add it to `shared-utilities/Config.gs`:

```javascript
const CONFIG = {
  SPREADSHEET_IDS: {
    'main': '1abc123def456ghi789jkl012mno345pqr',  // Your Sheet ID here
    'backup': '1xyz789abc123def456ghi789jkl012mno', // Another sheet if needed
  },
  // ... rest of config
};
```

## Example

**URL:**
```
https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit#gid=0
```

**Sheet ID:**
```
1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms
```

**Note:** The `#gid=0` part at the end is the **sheet tab ID** (which tab is active), not the Sheet ID. The Sheet ID is the long string before `/edit`.

## Troubleshooting

**"Invalid Sheet ID" error:**
- Make sure you copied the entire ID (it's usually 44 characters long)
- Don't include `/edit` or anything after it
- Make sure the sheet is accessible (not deleted or access revoked)

**"Permission denied" error:**
- Make sure your Google account has access to the sheet
- Check that the Apps Script project has the correct OAuth scopes

## Quick Test

To verify you have the correct Sheet ID, try this in Apps Script:

```javascript
function testSheetId() {
  const SHEET_ID = 'your-sheet-id-here';
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('✅ Sheet found: ' + sheet.getName());
    return true;
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    return false;
  }
}
```

Run this function - if it logs the sheet name, you have the correct ID!
