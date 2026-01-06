# Concatenation Configuration Guide

This guide explains how to configure the concatenation function to join cells with formatting preservation using the `Format_Config` sheet.

## Configuration Structure

Add your configurations to the **`Format_Config`** sheet in your Google Spreadsheet. The sheet should have these columns:

| Column | Header | Description | Example |
|--------|--------|-------------|---------|
| A | `ConfigKey` | **Unique identifier** for this configuration (can be any name) | `row8`, `performance_total`, `site_metrics` |
| B | `SourceSheet` | Name of the sheet containing source cells | `Performance and Progress` |
| C | `SourceRanges` | Comma-separated cell ranges to concatenate | `B8, C8` |
| D | `DestinationCell` | Cell where concatenated result will be placed | `D8` |

## ConfigKey Naming

**The `ConfigKey` does NOT need to be related to row numbers!** It's just a unique identifier for logging and tracking. You can use:

- **Row-based names**: `row8`, `row9`, `row10` (simple but can cause duplicates)
- **Descriptive names**: `performance_total`, `site_origin_response`, `cache_size`
- **Sheet-based names**: `perf_progress_row8`, `site_perf_row6`
- **Functional names**: `concatenate_metrics_row8`, `join_units_row6`

**Important**: Each `ConfigKey` must be unique. If duplicates are found, the script will automatically append the sheet name to make it unique (e.g., `row8` becomes `row8_Site_Performance`).

## Example Configurations

### Basic Example (Row-based naming)
```
ConfigKey: row8
SourceSheet: Performance and Progress
SourceRanges: B8, C8
DestinationCell: D8
```

### Descriptive naming (recommended)
```
ConfigKey: performance_total_memory
SourceSheet: Performance and Progress
SourceRanges: B8, C8
DestinationCell: D8
```

### Multiple source cells
```
ConfigKey: site_metrics
SourceSheet: Site Performance
SourceRanges: E6, F6, G6
DestinationCell: H6
```

### Different sheets
```
ConfigKey: summary_header
SourceSheet: Summary Sheet
SourceRanges: A1, B1
DestinationCell: C1
```

## How to Use

### Process All Configurations (Recommended)
```javascript
processAllConcatenations();  // Processes all entries from Format_Config sheet in order
```

### Process Single Configuration
```javascript
concatenateCellsWithFormatting('row9');  // Uses 'row9' configuration
concatenateCellsWithFormatting('performance_total');  // Uses descriptive key
```

## Adding New Configurations

1. **Open the `Format_Config` sheet** in your Google Spreadsheet

2. **Add a new row** with your configuration:
   - **Column A (ConfigKey)**: Choose a unique identifier (e.g., `performance_total`, `site_metrics`)
   - **Column B (SourceSheet)**: Sheet name (must match exactly, case-sensitive)
   - **Column C (SourceRanges)**: Comma-separated cell ranges (e.g., `B8, C8`)
   - **Column D (DestinationCell)**: Destination cell (e.g., `D8`)

3. **Run the script:**
   ```javascript
   processAllConcatenations();  // Processes all configurations including the new one
   ```

**No code changes needed!** The script reads directly from the `Format_Config` sheet.

## Configuration Parameters

| Parameter | Type | Description | Example |
|-----------|------|-------------|---------|
| `ConfigKey` | string | **Unique identifier** (any name you want) | `row8`, `performance_total`, `site_metrics` |
| `SourceSheet` | string | Name of the sheet containing source cells | `Performance and Progress` |
| `SourceRanges` | string | Comma-separated cell ranges (A1 notation) | `B8, C8` or `E6, F6, G6` |
| `DestinationCell` | string | Cell where concatenated result will be placed | `D8` |

## Notes

- **ConfigKey flexibility**: Use any identifier you want - it doesn't need to match row numbers
- **Order matters**: Configurations are processed in the order they appear in the `Format_Config` sheet
- **Cell order**: Within `SourceRanges`, cells are concatenated in the order listed
- **Formatting preserved**: Each source cell's formatting (color, font, size) is preserved in the destination
- **Empty cells**: Empty cells are handled gracefully (converted to empty strings)
- **Sheet name**: Must match exactly (case-sensitive)
- **Duplicate keys**: If duplicate `ConfigKey` values are found, the script automatically makes them unique by appending the sheet name

## Quick Reference

**To process all configurations:**
```javascript
processAllConcatenations();  // Processes all entries from Format_Config sheet
```

**To process a specific configuration:**
```javascript
concatenateCellsWithFormatting('your_config_key');  // Uses the ConfigKey from Format_Config
```

**To add a new configuration:**
1. Add a row to the `Format_Config` sheet
2. Run `processAllConcatenations()` - no code changes needed!
