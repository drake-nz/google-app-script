# How to Use Concatenation Functions

## Main Function: Process All Configurations

**Use this function to process ALL configurations from Format_Config sheet:**

```javascript
processAllConcatenations()
```

This function:
- ✅ Reads ALL configurations from Format_Config sheet
- ✅ Processes them in the order they appear (row9 first, then row8, etc.)
- ✅ No hardcoded values - everything comes from the sheet

## Single Configuration Function

**To process just one configuration:**

```javascript
concatenateCellsWithFormatting('row9');  // Process row9 only
concatenateCellsWithFormatting('row8');  // Process row8 only
```

## Format_Config Sheet Structure

Your Format_Config sheet should have:

| ConfigKey | SourceSheet | SourceRanges | DestinationCell |
|-----------|-------------|--------------|-----------------|
| row9 | Performance and Progress | B9, C9 | D9 |
| row8 | Performance and Progress | B8, C8 | D8 |

**Order matters!** Configurations are processed in the order they appear in the sheet.

## What Was Removed

All hardcoded configurations have been removed:
- ❌ No hardcoded 'row8' default
- ❌ No hardcoded 'Performance and Progress' sheet name
- ❌ No hardcoded B8, C8, D8 cell references
- ✅ Everything reads from Format_Config sheet

## Quick Start

1. **Set up Format_Config sheet** (if not already done):
   - Run `setupFormatConfigSheet()` to create it automatically
   - Or create manually with headers: ConfigKey | SourceSheet | SourceRanges | DestinationCell

2. **Add your configurations** to the sheet (one per row)

3. **Run the main function**:
   ```javascript
   processAllConcatenations()
   ```

4. **Check the logs** to see which configurations were processed

## Example Workflow

```
Format_Config Sheet:
Row 2: row9 | Performance and Progress | B9, C9 | D9
Row 3: row8 | Performance and Progress | B8, C8 | D8

Run: processAllConcatenations()

Result:
- Processes row9 first → populates D9
- Processes row8 second → populates D8
```

## Troubleshooting

**"No configurations found"**
- Check that Format_Config sheet exists
- Verify sheet has data starting from row 2
- Check that all columns are filled (ConfigKey, SourceSheet, SourceRanges, DestinationCell)

**"Configuration not found"**
- Verify ConfigKey matches exactly (case-sensitive)
- Check for extra spaces in the ConfigKey

**"Sheet not found"**
- Verify SourceSheet name matches exactly (case-sensitive)
- Check for extra spaces in sheet name
