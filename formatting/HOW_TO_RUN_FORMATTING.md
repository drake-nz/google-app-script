# How to Run the Formatting Function - Step by Step

## ⚠️ IMPORTANT: Which Function to Select

**DO NOT select `formatHeaderRow` from the dropdown!**

That function is a helper function and will give you an error.

## ✅ Correct Way to Run Formatting

### Step 1: Open Apps Script Editor

```bash
cd /Users/dot/Scripts/apps-script-projects/formatting
clasp open-script
```

### Step 2: Select the RIGHT Function

In the Apps Script editor, look at the function dropdown (top right, next to the Run button).

**Select ONE of these:**
- ✅ `simpleFormatTest` (EASIEST - recommended for first time)
- ✅ `testFormatMainSheet` (Full test with diagnostics)

**DO NOT select:**
- ❌ `formatHeaderRow` (This will give an error!)
- ❌ `formatDataRows` (Helper function)
- ❌ `formatSheet` (Helper function)
- ❌ `formatColumn` (Helper function)
- ❌ Any other function that doesn't start with "test" or "simple"

### Step 3: Click Run

1. Make sure `simpleFormatTest` (or `testFormatMainSheet`) is selected in the dropdown
2. Click the **Run** button (▶️) or press `Cmd+Enter` (Mac) / `Ctrl+Enter` (Windows)
3. If prompted, authorize the script (first time only)

### Step 4: Check the Logs

1. Click **View** → **Logs** (or press `Cmd+Enter` / `Ctrl+Enter`)
2. You should see:
   - "=== Simple Format Test ==="
   - "✅ Got sheet: TEST1"
   - "✅ Simple format test complete!"

### Step 5: Check Your Sheet

Open your Google Sheet and you should see:
- Blue header row (row 1) with white text
- Alternating row colors (if you have data)
- Auto-resized columns
- Borders around data

## Visual Guide

```
Apps Script Editor
┌─────────────────────────────────────┐
│ [Function Dropdown ▼] [Run ▶️]    │
│                                     │
│  Select: simpleFormatTest          │
│  NOT: formatHeaderRow ❌           │
└─────────────────────────────────────┘
```

## Troubleshooting

**If you see "Sheet is undefined":**
- You selected the wrong function
- Make sure you selected `simpleFormatTest` or `testFormatMainSheet`

**If you see "TEST1 sheet not found":**
- Run `diagnoseSheet()` first to see available sheet names
- Make sure the sheet name is exactly "TEST1" (all caps)

**If nothing happens:**
- Check View → Logs for error messages
- Make sure you authorized the script

## Quick Reference

| Function | When to Use | Status |
|----------|-------------|--------|
| `simpleFormatTest` | First time / Quick test | ✅ Use this |
| `testFormatMainSheet` | Full test with diagnostics | ✅ Use this |
| `diagnoseSheet` | Debug sheet access issues | ✅ Use for debugging |
| `formatHeaderRow` | Helper function - DO NOT USE | ❌ Don't select |
| `formatSheet` | Helper function - DO NOT USE | ❌ Don't select |

## Remember

**Always select a function that starts with "test" or "simple" when running from the dropdown!**
