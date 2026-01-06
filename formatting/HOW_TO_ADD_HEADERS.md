# How to Add Headers to Your Sheet

## Step 1: Open Your Google Sheet

You have two ways to open your sheet:

### Option A: Using the Sheet ID (Direct Link)

Open this URL in your browser:
```
https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID_HERE/edit
```

### Option B: Find It in Google Drive

1. Go to [Google Drive](https://drive.google.com)
2. Search for the sheet name (or look for recently opened sheets)
3. Click to open it

## Step 2: Add Headers to Row 1

Once the sheet is open:

1. **Click on cell A1** (top-left cell)
2. **Type your first column header** (e.g., "Name", "Date", "Status", etc.)
3. **Press Tab** to move to B1
4. **Type your second column header**
5. **Continue adding headers** across row 1 for all your columns

### Example Headers:

```
Row 1:  Name    |  Email           |  Date       |  Status
```

Or if you're not sure what headers to use, start with something simple:

```
Row 1:  Column A  |  Column B  |  Column C
```

**Important:** 
- Headers go in **Row 1** (the very first row)
- You need at least **one column** with a header
- Headers can be any text you want

## Step 3: Verify Headers Are Added

After adding headers:
- You should see text in row 1
- The sheet should have at least one column with data
- Row 1 should not be completely empty

## Step 4: Test the Formatting

After adding headers:

1. Go back to Apps Script editor (`clasp open-script`)
2. Run `diagnoseSheet()` first to verify it can see your data
3. Then run `testFormatMainSheet()` to format the sheet

## What Happens When You Format?

Once you have headers:
- ✅ Row 1 (headers) will get blue background with white text
- ✅ Data rows (if any) will get alternating colors
- ✅ Columns will auto-resize to fit content
- ✅ Borders will be added around the data

## If You Have Multiple Sheet Tabs

If your spreadsheet has multiple tabs (like "Sheet1", "Sheet2", etc.):

1. **Add headers to the tab you want to format**
2. **Note the exact tab name** (case-sensitive!)
3. **Use that name when formatting:**
   ```javascript
   formatMainSheet('Sheet1');  // Use exact tab name
   ```

## Still Can't Find Your Sheet?

Run this in Apps Script to get the sheet URL:

```javascript
function getSheetUrl() {
  const sheetId = CONFIG.SPREADSHEET_IDS.main;
  const url = `https://docs.google.com/spreadsheets/d/${sheetId}/edit`;
  Logger.log(`Open your sheet here: ${url}`);
  return url;
}
```

Run `getSheetUrl()` and check the logs - it will show you the direct link to your sheet!
