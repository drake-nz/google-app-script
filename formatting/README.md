# Formatting Project

Handles all formatting operations for Google Sheets.

## Functions

- `formatSheet(sheetId, sheetName)` - Format entire sheet (headers + data)
- `formatHeaderRow(sheet, rowNumber)` - Format header row
- `formatDataRows(sheet, startRow)` - Format data rows with alternating colors
- `formatColumn(sheet, columnIndex, formatOptions)` - Format specific column
- `autoResizeColumns(sheet)` - Auto-resize columns to fit content
- `addBorders(sheet)` - Add borders to data range

## Usage Example

```javascript
// Format a specific sheet
const SHEET_ID = 'your-sheet-id-here';
formatSheet(SHEET_ID, 'Data');

// Or format active sheet
const sheet = SpreadsheetApp.getActiveSheet();
formatHeaderRow(sheet);
formatDataRows(sheet);
autoResizeColumns(sheet);
```

## Setup

```bash
cd formatting
clasp create --title "Sheet Formatting" --type standalone
clasp push
```
