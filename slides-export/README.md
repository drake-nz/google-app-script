# Slides Export Project

Exports data from Google Sheets to Google Slides presentations.

## Functions

- `exportSheetToSlides(sheetId, sheetName, presentationTitle)` - Export entire sheet to new presentation
- `exportRangeToSlides(sheetId, sheetName, rangeA1, presentationTitle)` - Export specific range to slides
- `addSheetDataToPresentation(presentationId, sheetId, sheetName)` - Add sheet data to existing presentation

## Usage Examples

```javascript
// Export entire sheet to new presentation
const SHEET_ID = 'your-sheet-id-here';
const presentationId = exportSheetToSlides(
  SHEET_ID, 
  'Data', 
  'Monthly Report'
);

// Export specific range
const rangeId = exportRangeToSlides(
  SHEET_ID,
  'Data',
  'A1:D20',
  'Q4 Summary'
);

// Add data to existing presentation
addSheetDataToPresentation('existing-presentation-id', SHEET_ID, 'Data');
```

## Setup

```bash
cd slides-export
clasp create --title "Slides Export" --type standalone
clasp push
```
