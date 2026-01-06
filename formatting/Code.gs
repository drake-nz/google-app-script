/**
 * Sheet Formatting Project
 * Handles all formatting operations for Google Sheets
 */

// Configuration - Sheet IDs stored in PropertiesService (secure)
// Run setupSheetIds() in shared-utilities/Config.gs to configure
const CONFIG = {
  SPREADSHEET_IDS: {
    'main': '', // Stored in PropertiesService - run setupSheetIds() to configure
  },
  DEFAULT_SHEET_NAME: 'TEST1',
  FORMATTING: {
    HEADER_BACKGROUND: '#4285f4',
    HEADER_TEXT_COLOR: '#ffffff',
    HEADER_FONT_SIZE: 12,
    HEADER_FONT_WEIGHT: 'bold',
    DATA_ROW_BACKGROUND: '#f8f9fa',
    ALTERNATE_ROW_BACKGROUND: '#ffffff'
  },
  // Concatenation configuration (DEPRECATED - use Format_Config sheet instead)
  // This is kept only as a fallback if Format_Config sheet doesn't exist
  CONCATENATION: {
    // All configurations should be in Format_Config sheet
    // This fallback is empty - add configs to Format_Config sheet instead
  }
};

/**
 * Format header row
 * @param {Sheet} sheet - The sheet to format
 * @param {number} rowNumber - Row number (default: 1)
 * 
 * NOTE: This function should NOT be called directly from the function selector.
 * Use testFormatMainSheet() or simpleFormatTest() instead.
 */
function formatHeaderRow(sheet, rowNumber = 1) {
  Logger.log(`=== formatHeaderRow called ===`);
  Logger.log(`rowNumber: ${rowNumber}`);
  Logger.log(`Sheet parameter type: ${typeof sheet}`);
  Logger.log(`Sheet is null: ${sheet === null}`);
  Logger.log(`Sheet is undefined: ${sheet === undefined}`);
  
  // Check if called directly without sheet parameter
  if (sheet === undefined || sheet === null) {
    Logger.log('ERROR: formatHeaderRow was called without a sheet parameter!');
    Logger.log('This function should not be called directly.');
    Logger.log('Use testFormatMainSheet() or simpleFormatTest() instead.');
    throw new Error('formatHeaderRow requires a sheet parameter. Use testFormatMainSheet() or simpleFormatTest() instead.');
  }
  
  // Verify sheet has required methods
  if (typeof sheet.getLastColumn !== 'function') {
    Logger.log('ERROR: sheet.getLastColumn is not a function');
    throw new Error('Sheet object is invalid - missing getLastColumn method');
  }
  
  if (typeof sheet.getRange !== 'function') {
    Logger.log('ERROR: sheet.getRange is not a function');
    throw new Error('Sheet object is invalid - missing getRange method');
  }
  
  Logger.log('Sheet object is valid, proceeding...');
  const config = getConfig();
  const lastCol = sheet.getLastColumn();
  Logger.log(`Last column: ${lastCol}`);
  
  // Handle empty sheet - format at least column A
  const numCols = lastCol > 0 ? lastCol : 1;
  Logger.log(`Formatting ${numCols} columns`);
  const headerRange = sheet.getRange(rowNumber, 1, 1, numCols);
  Logger.log(`Got range: ${headerRange.getA1Notation()}`);
  
  headerRange.setBackground(config.FORMATTING.HEADER_BACKGROUND);
  headerRange.setFontColor(config.FORMATTING.HEADER_TEXT_COLOR);
  headerRange.setFontSize(config.FORMATTING.HEADER_FONT_SIZE);
  headerRange.setFontWeight(config.FORMATTING.HEADER_FONT_WEIGHT);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  
  Logger.log(`Formatted header row ${rowNumber} (${numCols} columns)`);
}

/**
 * Format data rows with alternating colors
 * @param {Sheet} sheet - The sheet to format
 * @param {number} startRow - Starting row (default: 2, after header)
 */
function formatDataRows(sheet, startRow = 2) {
  const config = getConfig();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < startRow) {
    Logger.log('No data rows to format');
    return;
  }
  
  for (let row = startRow; row <= lastRow; row++) {
    const rowRange = sheet.getRange(row, 1, 1, lastCol);
    const isEven = (row - startRow) % 2 === 0;
    
    rowRange.setBackground(
      isEven 
        ? config.FORMATTING.DATA_ROW_BACKGROUND 
        : config.FORMATTING.ALTERNATE_ROW_BACKGROUND
    );
  }
  
  Logger.log(`Formatted data rows ${startRow} to ${lastRow}`);
}

/**
 * Format entire sheet (headers + data)
 * @param {string} sheetId - The spreadsheet ID
 * @param {string} sheetName - The sheet name (optional)
 */
function formatSheet(sheetId, sheetName = null) {
  try {
    if (!sheetId) {
      throw new Error('Sheet ID is required');
    }
    
    Logger.log(`=== Starting formatSheet ===`);
    Logger.log(`Sheet ID: ${sheetId}`);
    Logger.log(`Sheet Name: ${sheetName || 'active (null)'}`);
    
    // Get the sheet with detailed logging
    Logger.log('Calling getSheetByName...');
    const sheet = getSheetByName(sheetId, sheetName);
    
    Logger.log(`Sheet retrieved: ${sheet ? 'SUCCESS' : 'FAILED'}`);
    
    if (!sheet) {
      Logger.log('ERROR: Sheet object is null after getSheetByName');
      throw new Error('Sheet object is null - check sheet name and permissions');
    }
    
    Logger.log(`Sheet name: ${sheet.getName()}`);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    Logger.log(`Sheet has ${lastRow} rows and ${lastCol} columns`);
    
    // Verify sheet object before formatting
    Logger.log(`Before formatHeaderRow - sheet type: ${typeof sheet}`);
    Logger.log(`Before formatHeaderRow - sheet is null: ${sheet === null}`);
    Logger.log(`Before formatHeaderRow - sheet is undefined: ${sheet === undefined}`);
    
    if (!sheet) {
      Logger.log('ERROR: Sheet is null/undefined before calling formatHeaderRow');
      throw new Error('Sheet object is null or undefined before formatting');
    }
    
    if (typeof sheet.getRange !== 'function') {
      Logger.log('ERROR: Sheet object missing getRange method');
      throw new Error('Sheet object is invalid - getRange method not found');
    }
    
    // Format header (always format header row even if sheet is empty)
    Logger.log('Calling formatHeaderRow with sheet object...');
    Logger.log(`Sheet name before call: ${sheet.getName()}`);
    formatHeaderRow(sheet, 1); // Explicitly pass rowNumber
    Logger.log('formatHeaderRow completed');
    
    // Format data rows only if there are data rows
    if (lastRow > 1) {
      Logger.log('Calling formatDataRows...');
      formatDataRows(sheet);
      Logger.log('formatDataRows completed');
    } else {
      Logger.log('No data rows to format (only header row exists)');
    }
    
    // Auto-resize columns only if there are columns
    if (lastCol > 0) {
      Logger.log('Calling autoResizeColumns...');
      autoResizeColumns(sheet);
      Logger.log('autoResizeColumns completed');
    }
    
    // Add borders only if there's data
    if (lastRow > 0 && lastCol > 0) {
      Logger.log('Calling addBorders...');
      addBorders(sheet);
      Logger.log('addBorders completed');
    }
    
    Logger.log(`✅ Successfully formatted sheet: ${sheetName || 'active sheet'}`);
  } catch (error) {
    Logger.log(`❌ Error formatting sheet: ${error.message}`);
    Logger.log(`Error type: ${error.constructor.name}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
  }
}

/**
 * Auto-resize columns to fit content
 * @param {Sheet} sheet - The sheet to format
 */
function autoResizeColumns(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol > 0) {
    sheet.autoResizeColumns(1, lastCol);
    Logger.log(`Auto-resized ${lastCol} columns`);
  }
}

/**
 * Add borders to data range
 * @param {Sheet} sheet - The sheet to format
 */
function addBorders(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow > 0 && lastCol > 0) {
    const range = sheet.getRange(1, 1, lastRow, lastCol);
    range.setBorder(true, true, true, true, true, true);
    Logger.log('Added borders to data range');
  }
}

/**
 * Format specific column
 * @param {Sheet} sheet - The sheet to format
 * @param {number} columnIndex - Column index (1-based)
 * @param {Object} formatOptions - Formatting options
 */
function formatColumn(sheet, columnIndex, formatOptions = {}) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return;
  
  const columnRange = sheet.getRange(1, columnIndex, lastRow, 1);
  
  if (formatOptions.backgroundColor) {
    columnRange.setBackground(formatOptions.backgroundColor);
  }
  if (formatOptions.fontColor) {
    columnRange.setFontColor(formatOptions.fontColor);
  }
  if (formatOptions.fontSize) {
    columnRange.setFontSize(formatOptions.fontSize);
  }
  if (formatOptions.numberFormat) {
    columnRange.setNumberFormat(formatOptions.numberFormat);
  }
  if (formatOptions.horizontalAlignment) {
    columnRange.setHorizontalAlignment(formatOptions.horizontalAlignment);
  }
  
  Logger.log(`Formatted column ${columnIndex}`);
}

/**
 * Helper: Get sheet by name or active sheet
 */
function getSheetByName(sheetId, sheetName) {
  try {
    Logger.log(`getSheetByName called with sheetId: ${sheetId}, sheetName: ${sheetName || 'null'}`);
    
    Logger.log('Opening spreadsheet...');
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    
    if (!spreadsheet) {
      Logger.log('ERROR: Spreadsheet object is null');
      throw new Error(`Could not open spreadsheet with ID: ${sheetId}`);
    }
    
    Logger.log(`✅ Spreadsheet opened: ${spreadsheet.getName()}`);
    
    // List all available sheets for debugging
    const allSheets = spreadsheet.getSheets();
    const sheetNames = allSheets.map(s => s.getName());
    Logger.log(`Available sheets: ${sheetNames.join(', ')}`);
    
    let sheet;
    if (sheetName) {
      Logger.log(`Looking for sheet named: "${sheetName}"`);
      sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log(`ERROR: Sheet "${sheetName}" not found`);
        Logger.log(`Available sheet names: ${sheetNames.join(', ')}`);
        throw new Error(`Sheet "${sheetName}" not found. Available sheets: ${sheetNames.join(', ')}`);
      }
      Logger.log(`✅ Found sheet: ${sheet.getName()}`);
    } else {
      Logger.log('Getting active sheet...');
      sheet = spreadsheet.getActiveSheet();
      if (!sheet) {
        Logger.log('ERROR: Active sheet is null');
        throw new Error(`No active sheet found in spreadsheet ${sheetId}`);
      }
      Logger.log(`✅ Active sheet: ${sheet.getName()}`);
    }
    
    // Verify sheet object
    if (typeof sheet.getName !== 'function') {
      Logger.log('ERROR: Sheet object is invalid');
      throw new Error('Sheet object is invalid');
    }
    
    Logger.log(`✅ Successfully retrieved sheet: ${sheet.getName()} (ID: ${sheetId})`);
    return sheet;
  } catch (error) {
    Logger.log(`❌ Error in getSheetByName: ${error.message}`);
    Logger.log(`Error stack: ${error.stack}`);
    throw error;
  }
}

/**
 * Helper: Get config
 */
function getConfig() {
  return CONFIG;
}

/**
 * Quick test function - formats your main sheet
 * Run this function to test the formatting on your sheet
 * 
 * NOTE: Your sheet needs to have at least a header row (row 1) with some data
 * Even if it's just one column with headers, that's enough to format
 */
function testFormatMainSheet() {
  const sheetId = CONFIG.SPREADSHEET_IDS.main;
  const sheetName = 'TEST1'; // Your sheet tab name (case-sensitive!)
  Logger.log(`=== Starting Format Test ===`);
  Logger.log(`Sheet ID: ${sheetId}`);
  Logger.log(`Sheet Name: ${sheetName}`);
  
  try {
    // First, run diagnostic
    Logger.log('Running diagnostic...');
    const diagnostic = diagnoseSheet();
    
    if (!diagnostic.success) {
      throw new Error(`Diagnostic failed: ${diagnostic.error}`);
    }
    
    Logger.log(`Diagnostic passed. Sheet: ${diagnostic.sheetName}, Rows: ${diagnostic.lastRow}, Cols: ${diagnostic.lastColumn}`);
    
    // Now format the TEST1 sheet
    formatSheet(sheetId, sheetName);
    Logger.log('✅ Formatting complete!');
  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
    Logger.log(`Full error: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    throw error;
  }
}

/**
 * Simple direct test - formats TEST1 sheet directly
 * Use this if testFormatMainSheet is having issues
 */
function simpleFormatTest() {
  try {
    Logger.log('=== Simple Format Test ===');
    const sheetId = getMainSheetId();
    Logger.log(`Sheet ID: ${sheetId}`);
    
    // Open spreadsheet
    Logger.log('Opening spreadsheet...');
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    Logger.log(`Spreadsheet opened: ${spreadsheet.getName()}`);
    
    // Get TEST1 sheet
    Logger.log('Getting TEST1 sheet...');
    const sheet = spreadsheet.getSheetByName('TEST1');
    
    if (!sheet) {
      // List available sheets
      const allSheets = spreadsheet.getSheets();
      const names = allSheets.map(s => s.getName());
      throw new Error(`TEST1 sheet not found. Available sheets: ${names.join(', ')}`);
    }
    
    Logger.log(`✅ Got sheet: ${sheet.getName()}`);
    Logger.log(`Sheet type: ${typeof sheet}`);
    Logger.log(`Has getRange method: ${typeof sheet.getRange === 'function'}`);
    
    // Verify sheet is valid before formatting
    const testRange = sheet.getRange(1, 1);
    Logger.log(`✅ Can access range A1: ${testRange.getA1Notation()}`);
    
    // Now format
    Logger.log('Calling formatHeaderRow...');
    formatHeaderRow(sheet);
    Logger.log('✅ formatHeaderRow completed');
    
    Logger.log('✅ Simple format test complete!');
  } catch (error) {
    Logger.log(`❌ Error in simpleFormatTest: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    throw error;
  }
}

/**
 * Get the main spreadsheet ID (secure - from PropertiesService)
 * @return {string} The main spreadsheet ID
 */
function getMainSheetId() {
  // Try PropertiesService first (most secure - stored in script properties)
  const properties = PropertiesService.getScriptProperties();
  const storedId = properties.getProperty('SPREADSHEET_ID_MAIN');
  if (storedId) {
    return storedId;
  }
  
  // Fallback to CONFIG (for local development)
  const configId = CONFIG.SPREADSHEET_IDS.main;
  if (configId && configId.trim() !== '') {
    return configId;
  }
  
  throw new Error('Spreadsheet ID not found. Run setupSheetIds() in shared-utilities/Config.gs to configure.');
}

/**
 * Format main sheet with specific sheet name
 * @param {string} sheetName - Name of the sheet tab to format
 */
function formatMainSheet(sheetName = null) {
  const sheetId = getMainSheetId();
  formatSheet(sheetId, sheetName);
}

/**
 * Get the URL to open your sheet in the browser
 */
function getSheetUrl() {
  const sheetId = getMainSheetId();
  const url = `https://docs.google.com/spreadsheets/d/${sheetId}/edit`;
  Logger.log(`📋 Open your sheet here: ${url}`);
  return url;
}

/**
 * DEPRECATED: Use processAllConcatenations() or concatenateCellsWithFormatting() instead
 * This function has been replaced by the configuration-based approach
 * Kept for backward compatibility only
 */
function concatenateB8C8ToD8() {
  Logger.log('⚠️  concatenateB8C8ToD8() is deprecated. Use processAllConcatenations() instead.');
  Logger.log('Processing all configurations from Format_Config sheet...');
  processAllConcatenations();
}

/**
 * Load concatenation configuration from Format_Config sheet
 * @return {Array} Array of config objects in sheet order, each with {key, sourceSheet, sourceRanges, destinationCell}
 */
function loadConcatenationConfigFromSheet() {
  try {
    const sheetId = getMainSheetId();
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const configSheet = spreadsheet.getSheetByName('Format_Config');
    
    if (!configSheet) {
      Logger.log('Format_Config sheet not found, using empty array');
      return [];
    }
    
    const lastRow = configSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('Format_Config sheet is empty');
      return [];
    }
    
    // Expected format:
    // Row 1: Headers (ConfigKey | SourceSheet | SourceRanges | DestinationCell)
    // Row 2+: Data
    const data = configSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const configs = [];
    
    // Track seen keys to detect duplicates
    const seenKeys = new Map();
    
    data.forEach((row, index) => {
      let configKey = String(row[0] || '').trim();
      const sourceSheet = String(row[1] || '').trim();
      const sourceRangesStr = String(row[2] || '').trim();
      const destinationCell = String(row[3] || '').trim();
      
      // Skip empty rows
      if (!configKey || !sourceSheet || !sourceRangesStr || !destinationCell) {
        return;
      }
      
      // Check for duplicate keys and make them unique
      if (seenKeys.has(configKey)) {
        // Duplicate found - append sheet name to make it unique
        const originalKey = configKey;
        configKey = `${configKey}_${sourceSheet.replace(/\s+/g, '_')}`;
        Logger.log(`Warning: Duplicate config key "${originalKey}" found. Using unique key "${configKey}"`);
      }
      seenKeys.set(configKey, true);
      
      // Parse source ranges (comma-separated)
      const sourceRanges = sourceRangesStr.split(',').map(r => r.trim()).filter(r => r);
      
      if (sourceRanges.length === 0) {
        Logger.log(`Warning: No source ranges for config "${configKey}"`);
        return;
      }
      
      configs.push({
        key: configKey,
        sourceSheet: sourceSheet,
        sourceRanges: sourceRanges,
        destinationCell: destinationCell
      });
      
      Logger.log(`Loaded config "${configKey}": ${sourceSheet} -> ${sourceRanges.join(', ')} -> ${destinationCell}`);
    });
    
    Logger.log(`Loaded ${configs.length} configurations from Format_Config sheet`);
    
    // Verify all configs were loaded and show order
    if (configs.length > 0) {
      const configKeys = configs.map(c => c.key);
      Logger.log(`Configuration order: ${configKeys.join(' -> ')}`);
    }
    
    return configs;
    
  } catch (error) {
    Logger.log(`Error loading config from sheet: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    return [];
  }
}

/**
 * Process ALL concatenation configurations from Format_Config sheet
 * Processes configurations in the order they appear in the sheet (row9, then row8, etc.)
 * This is the main function to run - it processes everything automatically
 */
function processAllConcatenations() {
  try {
    Logger.log(`=== Processing ALL concatenations from Format_Config sheet ===`);
    
    // Load all configurations from sheet (returns array in sheet order)
    const configs = loadConcatenationConfigFromSheet();
    
    if (configs.length === 0) {
      Logger.log('No configurations found in Format_Config sheet');
      return;
    }
    
    const configKeys = configs.map(c => c.key);
    Logger.log(`Found ${configs.length} configurations to process in order: ${configKeys.join(', ')}`);
    
    // Process each configuration in the order they appear in the sheet
    // Use for loop instead of forEach to ensure all are processed
    let successCount = 0;
    let errorCount = 0;
    
    for (let index = 0; index < configs.length; index++) {
      const config = configs[index];
      Logger.log(`\n--- Processing ${index + 1}/${configs.length}: ${config.key} ---`);
      Logger.log(`Config details: ${config.sourceSheet} -> ${config.sourceRanges.join(', ')} -> ${config.destinationCell}`);
      
      try {
        concatenateCellsWithFormatting(config);
        successCount++;
        Logger.log(`✅ Successfully completed ${config.key}`);
      } catch (error) {
        errorCount++;
        Logger.log(`❌ Error processing ${config.key}: ${error.message}`);
        Logger.log(`Stack: ${error.stack}`);
        // Continue with next configuration - don't stop on error
      }
    }
    
    Logger.log(`\n=== Summary ===`);
    Logger.log(`Total configurations: ${configs.length}`);
    Logger.log(`Successful: ${successCount}`);
    Logger.log(`Errors: ${errorCount}`);
    Logger.log(`✅ Finished processing all configurations`);
    
  } catch (error) {
    Logger.log(`❌ Error in processAllConcatenations: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    throw error;
  }
}

/**
 * Concatenate cells with formatting using configuration object
 * @param {Object|string} configOrKey - Either a config object or a config key string
 * If string provided, looks up config from Format_Config sheet
 */
function concatenateCellsWithFormatting(configOrKey = null) {
  try {
    const sheetId = getMainSheetId();
    let config;
    
    // Handle both config object and config key string
    if (typeof configOrKey === 'string') {
      // Look up config by key
      const configs = loadConcatenationConfigFromSheet();
      const foundConfig = configs.find(c => c.key === configOrKey);
      
      if (!foundConfig) {
        const availableKeys = configs.map(c => c.key).join(', ');
        throw new Error(`Configuration "${configOrKey}" not found. Available keys: ${availableKeys || 'none'}`);
      }
      
      config = foundConfig;
      Logger.log(`Found config for key "${configOrKey}"`);
    } else if (configOrKey && typeof configOrKey === 'object') {
      // Use provided config object directly
      config = configOrKey;
    } else {
      // No config provided, use first available
      const configs = loadConcatenationConfigFromSheet();
      if (configs.length === 0) {
        throw new Error('No configurations found in Format_Config sheet');
      }
      config = configs[0];
      Logger.log(`No config provided, using first available: ${config.key}`);
    }
    
    const sourceSheetName = config.sourceSheet;
    const sourceRanges = config.sourceRanges;
    const destinationCell = config.destinationCell;
    const configKey = config.key || 'unknown';
    
    Logger.log(`=== Concatenating cells with formatting ===`);
    Logger.log(`Config: ${configKey}`);
    Logger.log(`Source Sheet: ${sourceSheetName}`);
    Logger.log(`Source Ranges: ${sourceRanges.join(', ')}`);
    Logger.log(`Destination: ${destinationCell}`);
    
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const sheet = spreadsheet.getSheetByName(sourceSheetName);
    
    if (!sheet) {
      throw new Error(`Sheet "${sourceSheetName}" not found`);
    }
    
    // Get all source cells
    const sourceCells = sourceRanges.map(range => sheet.getRange(range));
    const destinationCellObj = sheet.getRange(destinationCell);
    
    // Get cell values and rich text values for all source cells
    const sourceData = sourceCells.map((cell, index) => {
      // Use getDisplayValue() to get the formatted/rounded value as shown in the cell
      // This respects the cell's number formatting (e.g., 1.79 displayed as 1.8)
      const displayValue = cell.getDisplayValue();
      const rawValue = cell.getValue(); // Keep for reference/logging
      
      let richText = null;
      try {
        richText = cell.getRichTextValue();
      } catch (e) {
        // Rich text not available
      }
      
      Logger.log(`Source ${index + 1} (${sourceRanges[index]}): display="${displayValue}", raw="${rawValue}" (type: ${typeof rawValue}), rich text: ${richText ? 'found' : 'null'}`);
      
      return {
        cell: cell,
        range: sourceRanges[index],
        value: displayValue, // Use display value (respects cell formatting/rounding)
        rawValue: rawValue,   // Keep raw value for reference
        richText: richText
      };
    });
    
    // Build the concatenated rich text
    const builder = SpreadsheetApp.newRichTextValue();
    let combinedText = '';
    
    // Build the complete text string by concatenating all source cells
    // Use display values directly (already strings) to preserve exact formatting
    combinedText = sourceData.map(item => item.value || '').join('');
    
    Logger.log(`Complete text to format: "${combinedText}" (length: ${combinedText.length})`);
    
    // Calculate ranges for each source cell in the final combined text
    let currentOffset = 0;
    const ranges = sourceData.map((item, index) => {
      // value is already a display string, no conversion needed
      const text = item.value || '';
      const startIdx = currentOffset;
      const endIdx = currentOffset + text.length;
      currentOffset = endIdx;
      
      Logger.log(`Source ${index + 1} (${item.range}): "${text}" (range: ${startIdx}-${endIdx})`);
      
      return {
        ...item,
        startIdx: startIdx,
        endIdx: endIdx
      };
    });
    
    // Set the text ONCE before applying any styles
    builder.setText(combinedText);
    
    // Apply formatting for each source cell
    ranges.forEach((rangeData, index) => {
      applyFormattingToRange(
        rangeData.richText,
        rangeData.value,
        rangeData.range,
        rangeData.cell,
        rangeData.startIdx,
        rangeData.endIdx
      );
    });
    
    // Helper function to create TextStyle from cell formatting
    function createTextStyleFromCell(cell) {
      const textStyleBuilder = SpreadsheetApp.newTextStyle();
      
      try {
        // Get font color - try hex string first (more reliable)
        let fontColorSet = false;
        try {
          const colorHex = cell.getFontColor();
          if (colorHex) {
            textStyleBuilder.setForegroundColor(colorHex);
            Logger.log(`Setting font color hex: ${colorHex}`);
            fontColorSet = true;
          }
        } catch (e1) {
          Logger.log(`Could not get font color hex: ${e1.message}`);
        }
        
        // If hex didn't work, try Color object
        if (!fontColorSet) {
          try {
            const fontColorObj = cell.getFontColorObject();
            if (fontColorObj) {
              textStyleBuilder.setForegroundColor(fontColorObj);
              Logger.log(`Setting font color object: ${fontColorObj}`);
              fontColorSet = true;
            }
          } catch (e2) {
            Logger.log(`Could not get font color object: ${e2.message}`);
          }
        }
        
        if (!fontColorSet) {
          Logger.log(`Warning: Could not set font color for cell`);
        }
        
        const fontFamily = cell.getFontFamily();
        if (fontFamily) {
          textStyleBuilder.setFontFamily(fontFamily);
        }
        
        const fontSize = cell.getFontSize();
        if (fontSize) {
          textStyleBuilder.setFontSize(fontSize);
          Logger.log(`Setting font size: ${fontSize}`);
        }
        
        const fontWeight = cell.getFontWeight();
        if (fontWeight === 'bold') {
          textStyleBuilder.setBold(true);
        }
        
        // Check italic using getFontStyle or try-catch
        try {
          const fontStyle = cell.getFontStyle();
          if (fontStyle === 'italic') {
            textStyleBuilder.setItalic(true);
          }
        } catch (e) {
          // Font style method might not be available, skip
        }
        
        // Check underline - use try-catch as method might not exist
        try {
          if (typeof cell.isUnderline === 'function' && cell.isUnderline()) {
            textStyleBuilder.setUnderline(true);
          }
        } catch (e) {
          // Method not available, skip
        }
        
        // Check strikethrough - use try-catch as method might not exist
        try {
          if (typeof cell.isStrikethrough === 'function' && cell.isStrikethrough()) {
            textStyleBuilder.setStrikethrough(true);
          }
        } catch (e) {
          // Method not available, skip
        }
      } catch (error) {
        Logger.log(`Error creating text style: ${error.message}`);
        Logger.log(`Stack: ${error.stack}`);
      }
      
      const builtStyle = textStyleBuilder.build();
      
      // Verify the color was set
      try {
        const verifyColor = builtStyle.getForegroundColorObject();
        Logger.log(`Verified TextStyle color: ${verifyColor ? verifyColor : 'null'}`);
      } catch (e) {
        Logger.log(`Could not verify TextStyle color: ${e.message}`);
      }
      
      return builtStyle;
    }
    
    // Helper function to apply cell formatting to builder using setTextStyle
    function applyCellFormatting(cell, startIdx, endIdx) {
      try {
        // Text should already be set - don't reset it here
        const textStyle = createTextStyleFromCell(cell);
        
        // Verify the style before applying
        try {
          const verifyColor = textStyle.getForegroundColorObject();
          Logger.log(`TextStyle color before applying: ${verifyColor ? verifyColor : 'null'}`);
        } catch (e) {
          Logger.log(`Could not verify TextStyle color: ${e.message}`);
        }
        
        builder.setTextStyle(startIdx, endIdx, textStyle);
        Logger.log(`Applied cell formatting to range ${startIdx}-${endIdx} (text length: ${combinedText.length})`);
        
        // Verify after applying by checking the builder's current state
        try {
          const tempRichText = builder.build();
          const tempRuns = tempRichText.getRuns();
          Logger.log(`After applying, builder has ${tempRuns.length} runs`);
          tempRuns.forEach((run, idx) => {
            const runStyle = run.getTextStyle();
            const color = runStyle.getForegroundColorObject();
            Logger.log(`  Builder run ${idx}: "${run.getText()}" (${run.getStartIndex()}-${run.getEndIndex()}) - color: ${color ? color : 'null'}`);
          });
        } catch (e) {
          Logger.log(`Could not verify builder state: ${e.message}`);
        }
      } catch (error) {
        Logger.log(`Error applying cell formatting: ${error.message}`);
        Logger.log(`Stack: ${error.stack}`);
      }
    }
    
    // Helper function to apply formatting to a specific range
    function applyFormattingToRange(richText, cellValue, cellName, cell, startIdx, endIdx) {
      if (startIdx >= endIdx) return; // Skip empty text
      
      if (!richText) {
        // No rich text, get formatting directly from cell
        Logger.log(`${cellName}: No rich text, using cell formatting (${startIdx}-${endIdx})`);
        applyCellFormatting(cell, startIdx, endIdx);
        return;
      }
      
      const runs = richText.getRuns();
      if (runs.length === 0) {
        // No runs, use cell formatting
        Logger.log(`${cellName}: No runs, using cell formatting (${startIdx}-${endIdx})`);
        applyCellFormatting(cell, startIdx, endIdx);
        return;
      }
      
      try {
        Logger.log(`${cellName}: Found ${runs.length} formatting runs`);
        
        // Process each run using setTextStyle API
        let currentOffset = startIdx;
        runs.forEach((run, index) => {
          const runText = run.getText();
          if (!runText) return; // Skip empty runs
          
          const runStartIdx = currentOffset;
          const runEndIdx = currentOffset + runText.length; // endIdx is exclusive
          currentOffset += runText.length;
          
          // Get the TextStyle from the run and apply it using setTextStyle
          const runStyle = run.getTextStyle();
          
          // Create a new TextStyle builder and copy all attributes
          const textStyleBuilder = SpreadsheetApp.newTextStyle();
          
          try {
            // Try to get color - prioritize hex string for reliability
            let colorSet = false;
            let colorToUse = null;
            
            // First try hex string method (more reliable)
            try {
              const colorHex = runStyle.getForegroundColor();
              if (colorHex) {
                colorToUse = colorHex;
                colorSet = true;
                Logger.log(`${cellName} run ${index}: Got color from runStyle (hex: ${colorHex})`);
              }
            } catch (e1) {
              Logger.log(`${cellName} run ${index}: Could not get hex color from runStyle: ${e1.message}`);
            }
            
            // Fallback to color object if hex not available
            if (!colorSet) {
              try {
                const colorObj = runStyle.getForegroundColorObject();
                if (colorObj) {
                  colorToUse = colorObj;
                  colorSet = true;
                  Logger.log(`${cellName} run ${index}: Got color from runStyle (object)`);
                }
              } catch (e2) {
                Logger.log(`${cellName} run ${index}: Could not get color object from runStyle: ${e2.message}`);
              }
            }
            
            // Final fallback: get from cell itself
            if (!colorSet) {
              try {
                const cellColorHex = cell.getFontColor();
                if (cellColorHex) {
                  colorToUse = cellColorHex;
                  colorSet = true;
                  Logger.log(`${cellName} run ${index}: Using cell color as fallback (hex: ${cellColorHex})`);
                }
              } catch (e3) {
                Logger.log(`${cellName} run ${index}: Could not get cell color: ${e3.message}`);
              }
            }
            
            // Apply the color if we found one
            if (colorSet && colorToUse) {
              textStyleBuilder.setForegroundColor(colorToUse);
              Logger.log(`${cellName} run ${index}: Applied color to TextStyle builder: ${typeof colorToUse === 'string' ? colorToUse : 'Color object'}`);
            } else {
              Logger.log(`${cellName} run ${index}: WARNING - No color found, TextStyle will use default`);
            }
            
            const fontFamily = runStyle.getFontFamily();
            if (fontFamily) {
              textStyleBuilder.setFontFamily(fontFamily);
            }
            
            const fontSize = runStyle.getFontSize();
            if (fontSize) {
              textStyleBuilder.setFontSize(fontSize);
              Logger.log(`${cellName} run ${index}: Setting font size ${fontSize}`);
            }
            
            if (runStyle.getBold !== undefined) {
              textStyleBuilder.setBold(runStyle.getBold());
            }
            
            if (runStyle.getItalic !== undefined) {
              textStyleBuilder.setItalic(runStyle.getItalic());
            }
            
            if (runStyle.getUnderline !== undefined) {
              textStyleBuilder.setUnderline(runStyle.getUnderline());
            }
            
            if (runStyle.getStrikethrough !== undefined) {
              textStyleBuilder.setStrikethrough(runStyle.getStrikethrough());
            }
            
            // Build the TextStyle and apply it
            const textStyle = textStyleBuilder.build();
            
            // Verify the TextStyle has the color before applying
            try {
              const builtColorHex = textStyle.getForegroundColor();
              const builtColorObj = textStyle.getForegroundColorObject();
              Logger.log(`${cellName} run ${index}: TextStyle built - hex: ${builtColorHex}, obj: ${builtColorObj ? 'set' : 'null'}`);
            } catch (e) {
              Logger.log(`${cellName} run ${index}: Could not verify built TextStyle color: ${e.message}`);
            }
            
            builder.setTextStyle(runStartIdx, runEndIdx, textStyle);
            Logger.log(`${cellName} run ${index}: Applied style to ${runStartIdx}-${runEndIdx}`);
            
            // Verify immediately after applying
            try {
              const tempRichText = builder.build();
              const tempRuns = tempRichText.getRuns();
              const relevantRun = tempRuns.find(r => r.getStartIndex() === runStartIdx);
              if (relevantRun) {
                const appliedStyle = relevantRun.getTextStyle();
                const appliedColorHex = appliedStyle.getForegroundColor();
                Logger.log(`${cellName} run ${index}: Verified after applying - hex: ${appliedColorHex}`);
              }
            } catch (e) {
              Logger.log(`${cellName} run ${index}: Could not verify applied style: ${e.message}`);
            }
            
          } catch (e) {
            Logger.log(`Error copying run style: ${e.message}`);
            Logger.log(`Stack: ${e.stack}`);
            // Fallback: apply cell formatting for this range
            applyCellFormatting(cell, runStartIdx, runEndIdx);
          }
        });
      } catch (error) {
        Logger.log(`Error copying runs from ${cellName}: ${error.message}`);
        Logger.log(`Stack: ${error.stack}`);
        // Fallback to cell formatting
        applyCellFormatting(cell, startIdx, endIdx);
      }
    }
    
    // Build the final rich text value
    Logger.log(`Final text length: ${combinedText.length}`);
    Logger.log(`Building rich text value...`);
    const finalRichText = builder.build();
    
    // Verify before setting
    const verifyRunsBefore = finalRichText.getRuns();
    Logger.log(`Rich text has ${verifyRunsBefore.length} runs before setting`);
    verifyRunsBefore.forEach((run, idx) => {
      try {
        const runStyle = run.getTextStyle();
        const color = runStyle.getForegroundColorObject();
        const colorHex = runStyle.getForegroundColor();
        Logger.log(`Run ${idx}: "${run.getText()}" - color object: ${color}, color hex: ${colorHex}, size: ${runStyle.getFontSize()}`);
      } catch (e) {
        Logger.log(`Run ${idx}: "${run.getText()}" - error getting style: ${e.message}`);
      }
    });
    
    Logger.log(`Setting rich text value to ${destinationCell}...`);
    destinationCellObj.setRichTextValue(finalRichText);
    
    // Explicitly set cell format to Plain text to prevent percentage auto-detection
    // This helps preserve rich text formatting, especially for cells ending with %
    try {
      destinationCellObj.setNumberFormat('@'); // '@' is the format code for Plain text
      Logger.log(`Set cell format to Plain text for ${destinationCell}`);
    } catch (e) {
      Logger.log(`Could not set cell format: ${e.message}`);
    }
    
    // Force a flush by reading the cell value (helps ensure rich text is persisted)
    try {
      destinationCellObj.getValue();
    } catch (e) {
      // Ignore errors, just trying to force a flush
    }
    
    const sourceValues = sourceData.map(item => `"${item.value}"`).join(' + ');
    Logger.log(`✅ Concatenated ${sourceValues} = "${combinedText}"`);
    Logger.log(`✅ Value set to ${destinationCell} with formatting preserved`);
    
    // Verify the result (with null check and retry)
    let verificationSuccess = false;
    for (let retry = 0; retry < 2; retry++) {
      try {
        // Small delay on retry to allow Google Sheets to process
        if (retry > 0) {
          Utilities.sleep(100); // 100ms delay
        }
        
        const verifyRichText = destinationCellObj.getRichTextValue();
        if (verifyRichText) {
          const verifyRuns = verifyRichText.getRuns();
          Logger.log(`Verification: ${destinationCell} has ${verifyRuns.length} formatting runs`);
          
          // Log the actual colors being used
          verifyRuns.forEach((run, idx) => {
            try {
              const runStyle = run.getTextStyle();
              const colorObj = runStyle.getForegroundColorObject();
              const colorHex = runStyle.getForegroundColor();
              Logger.log(`  Verified run ${idx}: "${run.getText()}" - color obj: ${colorObj ? 'set' : 'null'}, hex: ${colorHex}`);
            } catch (e) {
              Logger.log(`  Verified run ${idx}: "${run.getText()}" - error reading color`);
            }
          });
          
          verificationSuccess = true;
          break;
        } else {
          if (retry === 0) {
            Logger.log(`Verification attempt ${retry + 1}: ${destinationCell} rich text is null, retrying...`);
          } else {
            Logger.log(`Verification: ${destinationCell} rich text is null after retry (formatting may still be applied)`);
          }
        }
      } catch (verifyError) {
        if (retry === 0) {
          Logger.log(`Verification attempt ${retry + 1} error: ${verifyError.message}, retrying...`);
        } else {
          Logger.log(`Verification warning: Could not verify rich text for ${destinationCell}: ${verifyError.message}`);
        }
      }
    }
    
    if (!verificationSuccess) {
      Logger.log(`Note: Verification failed for ${destinationCell}, but value was set. Check the cell visually to confirm formatting.`);
    }
    
  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    throw error;
  }
}

/**
 * Diagnostic function - list all configurations from Format_Config sheet
 * Use this to verify configurations are being read correctly
 */
function listConcatenationConfigs() {
  try {
    Logger.log(`=== Listing all configurations from Format_Config sheet ===`);
    
    const configs = loadConcatenationConfigFromSheet();
    
    if (configs.length === 0) {
      Logger.log('No configurations found');
      return [];
    }
    
    Logger.log(`\nFound ${configs.length} configuration(s):\n`);
    
    configs.forEach((config, index) => {
      Logger.log(`${index + 1}. Config Key: "${config.key}"`);
      Logger.log(`   Source Sheet: "${config.sourceSheet}"`);
      Logger.log(`   Source Ranges: ${config.sourceRanges.join(', ')}`);
      Logger.log(`   Destination: ${config.destinationCell}`);
      Logger.log('');
    });
    
    Logger.log(`\nTo process all, run: processAllConcatenations()`);
    if (configs.length > 0) {
      Logger.log(`To process one, run: concatenateCellsWithFormatting('${configs[0].key}')`);
    }
    
    return configs;
    
  } catch (error) {
    Logger.log(`❌ Error listing configs: ${error.message}`);
    throw error;
  }
}

/**
 * Diagnostic function - check sheet access and properties
 * Run this to debug sheet connection issues
 */
function diagnoseSheet() {
  const sheetId = CONFIG.SPREADSHEET_IDS.main;
  Logger.log('=== Sheet Diagnostic ===');
  Logger.log(`Sheet ID: ${sheetId}`);
  
  try {
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    Logger.log(`✅ Spreadsheet opened: ${spreadsheet.getName()}`);
    
    const sheetNames = spreadsheet.getSheets().map(s => s.getName());
    Logger.log(`Available sheets: ${sheetNames.join(', ')}`);
    
    const activeSheet = spreadsheet.getActiveSheet();
    Logger.log(`Active sheet: ${activeSheet.getName()}`);
    Logger.log(`Last row: ${activeSheet.getLastRow()}`);
    Logger.log(`Last column: ${activeSheet.getLastColumn()}`);
    
    if (activeSheet.getLastRow() === 0 || activeSheet.getLastColumn() === 0) {
      Logger.log('⚠️  Sheet appears to be empty');
    }
    
    Logger.log('=== End Diagnostic ===');
    return {
      success: true,
      sheetName: activeSheet.getName(),
      lastRow: activeSheet.getLastRow(),
      lastColumn: activeSheet.getLastColumn()
    };
  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Helper function to check if a cell range matches any source range in configurations
 * @param {string} editedSheetName - Name of the sheet that was edited
 * @param {string} editedCellA1 - A1 notation of the edited cell (e.g., "B8")
 * @returns {boolean} - True if the edited cell matches any source range
 */
function isCellInSourceRanges(editedSheetName, editedCellA1) {
  try {
    const configs = loadConcatenationConfigFromSheet();
    
    for (const config of configs) {
      // Check if edited sheet matches this config's source sheet
      if (config.sourceSheet !== editedSheetName) {
        continue;
      }
      
      // Check if edited cell is in any of the source ranges
      for (const sourceRange of config.sourceRanges) {
        // Handle both single cells (B8) and ranges (B8:C8)
        const rangeParts = sourceRange.split(':');
        const startCell = rangeParts[0];
        
        // If edited cell matches the start cell or is in the range
        if (editedCellA1 === startCell) {
          return true;
        }
        
        // If it's a range, check if edited cell is within it
        if (rangeParts.length === 2) {
          const endCell = rangeParts[1];
          // Simple check: if edited cell matches start or end, it's in range
          // For more complex range checking, you'd need to parse A1 notation
          if (editedCellA1 === startCell || editedCellA1 === endCell) {
            return true;
          }
        }
      }
    }
    
    return false;
  } catch (error) {
    Logger.log(`Error checking source ranges: ${error.message}`);
    return false;
  }
}

/**
 * Simple trigger: Runs automatically when a cell is edited
 * NOTE: Simple triggers have limitations (6 min execution time, no external APIs)
 * For more reliability, use installable triggers (see setupAutoUpdateTrigger)
 * 
 * @param {Event} e - The edit event
 */
function onEdit(e) {
  try {
    // Get edit information
    const editedSheet = e.source.getActiveSheet();
    const editedSheetName = editedSheet.getName();
    const editedRange = e.range;
    const editedCellA1 = editedRange.getA1Notation();
    
    // Skip if editing Format_Config sheet (to avoid infinite loops)
    if (editedSheetName === 'Format_Config') {
      return;
    }
    
    // Check if the edited cell is in any source range
    const isSourceCell = isCellInSourceRanges(editedSheetName, editedCellA1);
    
    if (isSourceCell) {
      Logger.log(`📝 Cell ${editedCellA1} in sheet "${editedSheetName}" was edited - triggering concatenation update`);
      
      // Process all concatenations (will update all affected destination cells)
      // Use try-catch to prevent errors from breaking the sheet
      try {
        processAllConcatenations();
        Logger.log(`✅ Concatenation update completed after edit to ${editedCellA1}`);
      } catch (error) {
        Logger.log(`❌ Error updating concatenation: ${error.message}`);
        // Don't throw - we don't want to break the edit operation
      }
    }
  } catch (error) {
    Logger.log(`❌ Error in onEdit trigger: ${error.message}`);
    // Don't throw - simple triggers should not throw errors
  }
}

/**
 * Set up an installable trigger for automatic updates
 * Installable triggers are more reliable than simple triggers
 * 
 * Run this function ONCE to set up the trigger
 * 
 * To run: In Apps Script editor, select setupAutoUpdateTrigger and click Run
 */
function setupAutoUpdateTrigger() {
  try {
    const sheetId = getMainSheetId();
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    
    // Delete existing triggers with the same name (if any)
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onEditTrigger') {
        ScriptApp.deleteTrigger(trigger);
        Logger.log('Deleted existing trigger');
      }
    });
    
    // Create new installable trigger
    const trigger = ScriptApp.newTrigger('onEditTrigger')
      .onEdit()
      .create();
    
    Logger.log('✅ Installable trigger created successfully!');
    Logger.log(`Trigger ID: ${trigger.getUniqueId()}`);
    Logger.log('The trigger will run automatically when cells are edited.');
    
    return trigger;
  } catch (error) {
    Logger.log(`❌ Error setting up trigger: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    throw error;
  }
}

/**
 * Installable trigger handler (more reliable than simple onEdit)
 * This function is called by the installable trigger
 * 
 * @param {Event} e - The edit event
 */
function onEditTrigger(e) {
  try {
    // Get edit information
    const editedSheet = e.source.getActiveSheet();
    const editedSheetName = editedSheet.getName();
    const editedRange = e.range;
    const editedCellA1 = editedRange.getA1Notation();
    
    // Skip if editing Format_Config sheet
    if (editedSheetName === 'Format_Config') {
      return;
    }
    
    // Check if the edited cell is in any source range
    const isSourceCell = isCellInSourceRanges(editedSheetName, editedCellA1);
    
    if (isSourceCell) {
      Logger.log(`📝 Cell ${editedCellA1} in sheet "${editedSheetName}" was edited - triggering concatenation update`);
      
      // Process all concatenations
      try {
        processAllConcatenations();
        Logger.log(`✅ Concatenation update completed after edit to ${editedCellA1}`);
      } catch (error) {
        Logger.log(`❌ Error updating concatenation: ${error.message}`);
        // Installable triggers can throw, but we'll log and continue
      }
    }
  } catch (error) {
    Logger.log(`❌ Error in onEditTrigger: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    // Installable triggers can throw errors, but we'll log them
  }
}

/**
 * Remove all auto-update triggers
 * Run this if you want to disable automatic updates
 */
function removeAutoUpdateTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removedCount = 0;
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onEditTrigger' || 
          trigger.getHandlerFunction() === 'onEdit') {
        ScriptApp.deleteTrigger(trigger);
        removedCount++;
        Logger.log(`Deleted trigger: ${trigger.getHandlerFunction()}`);
      }
    });
    
    Logger.log(`✅ Removed ${removedCount} trigger(s)`);
    return removedCount;
  } catch (error) {
    Logger.log(`❌ Error removing triggers: ${error.message}`);
    throw error;
  }
}
