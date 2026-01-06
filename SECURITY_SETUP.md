# Security Setup Guide

This guide explains how to protect sensitive data (like Spreadsheet IDs) in your public GitHub repository.

## Problem

Spreadsheet IDs and other sensitive configuration should not be committed to public repositories.

## Solution: Use PropertiesService

Google Apps Script provides `PropertiesService` to store sensitive data securely in script properties (not in code).

## Quick Setup

### Step 1: Store Your Sheet IDs Securely

1. **Open Apps Script Editor:**
   ```bash
   clasp open-script
   ```
   Or: Extensions → Apps Script

2. **Run the setup function:**
   - Select `setupSheetIds` from the function dropdown
   - Click Run (▶️)
   - Authorize permissions if prompted

3. **Or set IDs programmatically:**
   ```javascript
   // In Apps Script editor, run this once:
   setupSheetIds({
     main: '1MmoIyz2t9WlzBtcuywhUOExzejHNbx2Vs5cSwXgszeg'
   });
   ```

### Step 2: Remove IDs from Code

After running `setupSheetIds()`, your IDs are stored securely. You can now:

1. **Update Config.gs** to use empty placeholders (already done)
2. **Commit the changes** - IDs are no longer in code

### Step 3: Verify It Works

1. **Test the functions:**
   ```javascript
   // Should work without IDs in code
   getMainSheetId();  // Returns ID from PropertiesService
   processAllConcatenations();  // Uses stored ID
   ```

2. **Check PropertiesService:**
   - In Apps Script editor: Project Settings (gear icon)
   - Scroll to "Script properties"
   - You should see `SPREADSHEET_ID_MAIN` with your ID

## How It Works

### Before (Insecure - IDs in code)
```javascript
const CONFIG = {
  SPREADSHEET_IDS: {
    'main': '1MmoIyz2t9WlzBtcuywhUOExzejHNbx2Vs5cSwXgszeg', // ❌ Visible in GitHub
  }
};
```

### After (Secure - IDs in PropertiesService)
```javascript
// Code (public)
const CONFIG = {
  SPREADSHEET_IDS: {
    'main': '', // ✅ Empty - stored in PropertiesService
  }
};

// PropertiesService (private, not in code)
SPREADSHEET_ID_MAIN = '1MmoIyz2t9WlzBtcuywhUOExzejHNbx2Vs5cSwXgszeg'
```

### Code Reads from PropertiesService
```javascript
function getMainSheetId() {
  // Try PropertiesService first (secure)
  const properties = PropertiesService.getScriptProperties();
  const storedId = properties.getProperty('SPREADSHEET_ID_MAIN');
  if (storedId) {
    return storedId; // ✅ Returns from PropertiesService
  }
  
  // Fallback to CONFIG (for local development)
  return CONFIG.SPREADSHEET_IDS.main;
}
```

## Alternative: Local Config File (For Development)

If you prefer to keep IDs in a file for local development:

1. **Create `Config.local.gs`** (gitignored):
   ```javascript
   // This file is gitignored - add your real IDs here
   const CONFIG = {
     SPREADSHEET_IDS: {
       'main': '1MmoIyz2t9WlzBtcuywhUOExzejHNbx2Vs5cSwXgszeg',
     }
   };
   ```

2. **Use `Config.example.gs`** as template (committed to GitHub)

3. **Update `.gitignore`** (already done):
   ```
   **/Config.local.gs
   ```

## For Each New Project

When setting up a new Apps Script project:

1. **Copy `Config.example.gs`** to your project
2. **Run `setupSheetIds()`** with your IDs
3. **Remove IDs from code** before committing

## Security Best Practices

1. ✅ **Use PropertiesService** for production/shared projects
2. ✅ **Never commit** real IDs to public repositories
3. ✅ **Use placeholders** in example/template files
4. ✅ **Document** where IDs are stored (PropertiesService)
5. ✅ **Rotate IDs** if accidentally committed (change sharing permissions)

## Troubleshooting

### "Spreadsheet ID not found" Error

**Solution:**
1. Run `setupSheetIds()` to store your IDs
2. Or temporarily add ID to `CONFIG.SPREADSHEET_IDS.main` for testing

### "PropertiesService not available"

**Solution:**
- PropertiesService is always available in Apps Script
- Make sure you're running in Apps Script environment (not local)

### "How do I see stored properties?"

**Solution:**
- Apps Script Editor → Project Settings (gear icon)
- Scroll to "Script properties"
- View/edit stored properties there

## Migration Checklist

- [x] Update `Config.gs` to use PropertiesService
- [x] Create `Config.example.gs` template
- [x] Update `.gitignore` to exclude local configs
- [ ] Run `setupSheetIds()` in Apps Script
- [ ] Remove any remaining IDs from code
- [ ] Test that functions still work
- [ ] Commit changes (IDs no longer in code)

---

**Note:** PropertiesService stores data per script/project. Each Apps Script project has its own properties.
