# Format_Config Sheet Setup Guide

This guide explains how to set up the `Format_Config` sheet to configure concatenation instead of editing Code.gs.

## Sheet Structure

Create a sheet named **`Format_Config`** in your spreadsheet with the following structure:

### Headers (Row 1)
| Column A | Column B | Column C | Column D |
|----------|----------|----------|----------|
| ConfigKey | SourceSheet | SourceRanges | DestinationCell |

### Data Rows (Row 2+)
| ConfigKey | SourceSheet | SourceRanges | DestinationCell |
|-----------|-------------|--------------|-----------------|
| row8 | Performance and Progress | B8, C8 | D8 |
| row9 | Performance and Progress | B9, C9 | D9 |
| row10 | Performance and Progress | B10, C10, E10 | F10 |

## Column Descriptions

### Column A: ConfigKey
- **Purpose**: Unique identifier for this configuration
- **Format**: Any text (e.g., "row8", "row9", "summary")
- **Required**: Yes

### Column B: SourceSheet
- **Purpose**: Name of the sheet containing source cells
- **Format**: Exact sheet name (case-sensitive)
- **Required**: Yes
- **Example**: `Performance and Progress`

### Column C: SourceRanges
- **Purpose**: Cell ranges to concatenate
- **Format**: Comma-separated cell references (A1 notation)
- **Required**: Yes
- **Examples**: 
  - `B8, C8` (two cells)
  - `B9, C9, E9` (three cells)
  - `A1, B1` (two cells)

### Column D: DestinationCell
- **Purpose**: Where to put the concatenated result
- **Format**: Single cell reference (A1 notation)
- **Required**: Yes
- **Example**: `D8`

## Setup Steps

### Step 1: Create the Sheet
1. Open your Google Sheet
2. Click the **+** button at the bottom to add a new sheet
3. Rename it to **`Format_Config`**

### Step 2: Add Headers
In Row 1, add these headers:
```
A1: ConfigKey
B1: SourceSheet
C1: SourceRanges
D1: DestinationCell
```

### Step 3: Add Your Configurations
Starting from Row 2, add your configurations:

**Example Row 2:**
```
A2: row8
B2: Performance and Progress
C2: B8, C8
D2: D8
```

**Example Row 3:**
```
A3: row9
B3: Performance and Progress
C3: B9, C9
D3: D9
```

## Usage

### Run with Config Key
```javascript
concatenateCellsWithFormatting('row8');  // Uses row8 from Format_Config sheet
concatenateCellsWithFormatting('row9');  // Uses row9 from Format_Config sheet
```

### Default (if no key provided)
```javascript
concatenateCellsWithFormatting();  // Uses 'row8' by default
```

## Example Sheet Layout

```
┌──────────┬──────────────────────────┬──────────────┬─────────────────┐
│ ConfigKey│ SourceSheet              │ SourceRanges │ DestinationCell │
├──────────┼──────────────────────────┼──────────────┼─────────────────┤
│ row8     │ Performance and Progress │ B8, C8       │ D8              │
│ row9     │ Performance and Progress │ B9, C9       │ D9              │
│ row10    │ Performance and Progress │ B10, C10, E10│ F10             │
│ summary  │ Summary Sheet            │ A1, B1       │ C1              │
└──────────┴──────────────────────────┴──────────────┴─────────────────┘
```

## Notes

- **Case-sensitive**: Sheet names must match exactly
- **Comma-separated**: Use commas to separate multiple source ranges
- **No spaces required**: Spaces around commas are optional (`B8,C8` or `B8, C8` both work)
- **Empty rows**: Empty rows are automatically skipped
- **Fallback**: If Format_Config sheet doesn't exist, falls back to CONFIG.CONCATENATION in Code.gs

## Benefits

✅ **No code changes needed** - Just edit the sheet  
✅ **Easy to manage** - Add/remove configurations without touching code  
✅ **Visual** - See all configurations at a glance  
✅ **Non-technical friendly** - Anyone can update the sheet  

## Troubleshooting

**"Configuration not found"**
- Check that ConfigKey matches exactly (case-sensitive)
- Verify the Format_Config sheet exists
- Check that the row isn't empty

**"Sheet not found"**
- Verify SourceSheet name matches exactly (case-sensitive)
- Check for extra spaces in the sheet name

**"No source ranges"**
- Make sure SourceRanges column has values
- Check that ranges are comma-separated

## Quick Start Template

Copy this into your Format_Config sheet:

```
ConfigKey	SourceSheet	SourceRanges	DestinationCell
row8	Performance and Progress	B8, C8	D8
```

Then run:
```javascript
concatenateCellsWithFormatting('row8');
```
