# Shared Utilities Library

Common functions and configuration shared across all Apps Script projects.

## Setup as Library

1. Deploy this project:
   ```bash
   cd shared-utilities
   clasp create --title "Shared Utilities" --type standalone
   clasp push
   ```

2. Get the Script ID:
   - Run `clasp status` or check `.clasp.json`
   - Or find it in the Apps Script editor URL

3. Add as library in other projects:
   - Open other project in Apps Script editor
   - Go to **Resources > Libraries**
   - Add library using the Script ID
   - Use identifier: `SharedUtils`

## Usage in Other Projects

After adding as library:

```javascript
// Access config
const sheetId = SharedUtils.CONFIG.SPREADSHEET_IDS.main;

// Use utility functions
const sheet = SharedUtils.getSheet(sheetId, 'Data');
const headers = SharedUtils.getHeaders(sheet);
SharedUtils.log('Processing complete');
```

## Files

- `Config.gs` - Configuration constants and helper functions
- `Utils.gs` - Common utility functions
- `appsscript.json` - Project manifest
