# Formatting Project - Setup Steps

Follow these steps to set up and test your formatting project:

## Step 1: Navigate to the Formatting Directory

```bash
cd /Users/dot/Scripts/apps-script-projects/formatting
```

## Step 2: Create the Apps Script Project

```bash
clasp create --title "Sheet Formatting" --type standalone
```

This will:
- Create a new Apps Script project in your Google account
- Generate a `.clasp.json` file with your project ID

## Step 3: Push Your Code to Google

```bash
clasp push
```

This uploads your `Code.gs` file to Google Apps Script.

## Step 4: Open in Browser to Test

```bash
clasp open-script
```

Or alternatively:
```bash
clasp open-container
```

This opens your project in the Google Apps Script editor in your browser.

## Step 5: Test the Formatting

In the Apps Script editor:

1. **Select a function** from the dropdown (top right):
   - Choose `testFormatMainSheet` for a quick test
   - Or choose `formatMainSheet` if you want to specify a sheet name

2. **Click the Run button** (▶️) or press `Cmd+Enter` (Mac) / `Ctrl+Enter` (Windows)

3. **Authorize the script** (first time only):
   - Click "Review Permissions"
   - Choose your Google account
   - Click "Advanced" → "Go to [Project Name] (unsafe)"
   - Click "Allow"

4. **Check the logs**:
   - Click "View" → "Logs" or press `Cmd+Enter` / `Ctrl+Enter`
   - You should see: `✅ Formatting complete!`

5. **Check your sheet**:
   - Go back to your Google Sheet
   - You should see formatted headers and alternating row colors!

## Available Functions

### Quick Test Functions

- **`testFormatMainSheet()`** - Formats your main sheet (uses active sheet if no name specified)
- **`formatMainSheet(sheetName)`** - Formats a specific sheet tab by name

### Core Formatting Functions

- **`formatSheet(sheetId, sheetName)`** - Format entire sheet
- **`formatHeaderRow(sheet, rowNumber)`** - Format just the header row
- **`formatDataRows(sheet, startRow)`** - Format data rows with alternating colors
- **`formatColumn(sheet, columnIndex, formatOptions)`** - Format a specific column
- **`autoResizeColumns(sheet)`** - Auto-resize columns to fit content
- **`addBorders(sheet)`** - Add borders to data range

## Example Usage

### Format the active sheet in your main spreadsheet:
```javascript
testFormatMainSheet();
```

### Format a specific sheet tab:
```javascript
formatMainSheet('Data');  // Replace 'Data' with your sheet tab name
```

### Format a different spreadsheet:
```javascript
const otherSheetId = 'another-sheet-id-here';
formatSheet(otherSheetId, 'Sheet1');
```

### Format just headers:
```javascript
const sheetId = CONFIG.SPREADSHEET_IDS.main;
const spreadsheet = SpreadsheetApp.openById(sheetId);
const sheet = spreadsheet.getActiveSheet();
formatHeaderRow(sheet);
```

## Troubleshooting

**"Script function not found"**
- Make sure you've pushed your code: `clasp push`
- Refresh the Apps Script editor page

**"Permission denied"**
- Make sure you've authorized the script (see Step 5)
- Check that your Google account has access to the sheet

**"Sheet not found"**
- Verify your Sheet ID is correct in `Code.gs`
- Make sure the sheet name matches exactly (case-sensitive)

**"No data rows to format"**
- Your sheet might be empty
- Add some data to test the formatting

## Next Steps

Once formatting is working:
1. Customize colors in the `CONFIG` object
2. Add more formatting functions as needed
3. Set up triggers to auto-format on sheet changes
4. Move on to the slides-export project!

## Viewing Logs

To see what's happening:
```bash
clasp logs
```

Or in the Apps Script editor: View → Logs
