/**
 * Shared Configuration - EXAMPLE/TEMPLATE
 * 
 * Copy this file to Config.local.gs and add your actual Sheet IDs
 * Config.local.gs is gitignored and will not be committed
 * 
 * OR use PropertiesService (recommended for production):
 * - Run setupSheetIds() once to store IDs securely
 * - IDs will be stored in script properties (not in code)
 */

const CONFIG = {
  // Add your Sheet IDs here
  // Get from URL: https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit
  SPREADSHEET_IDS: {
    'main': 'YOUR_MAIN_SHEET_ID_HERE',
    // Add more sheet IDs here as needed
    // Example: 'backup': '1abc123def456...',
  },
  
  // Default sheet name
  DEFAULT_SHEET_NAME: 'Sheet1',
  
  // Formatting defaults
  FORMATTING: {
    HEADER_BACKGROUND: '#4285f4',
    HEADER_TEXT_COLOR: '#ffffff',
    HEADER_FONT_SIZE: 12,
    HEADER_FONT_WEIGHT: 'bold',
    DATA_ROW_BACKGROUND: '#f8f9fa',
    ALTERNATE_ROW_BACKGROUND: '#ffffff'
  },
  
  // Slides export defaults
  SLIDES: {
    TEMPLATE_ID: '', // Optional: Template presentation ID
    TITLE_SLIDE_LAYOUT: 'TITLE',
    CONTENT_SLIDE_LAYOUT: 'TITLE_AND_BODY',
    DEFAULT_THEME: 'GOOGLE'
  }
};

/**
 * Get spreadsheet by ID
 * @param {string} sheetId - The spreadsheet ID
 * @return {Spreadsheet} The spreadsheet object
 */
function getSpreadsheet(sheetId) {
  try {
    return SpreadsheetApp.openById(sheetId);
  } catch (error) {
    Logger.log(`Error opening spreadsheet ${sheetId}: ${error.message}`);
    throw error;
  }
}

/**
 * Get sheet by name from spreadsheet
 * @param {string} sheetId - The spreadsheet ID
 * @param {string} sheetName - The sheet name
 * @return {Sheet} The sheet object
 */
function getSheet(sheetId, sheetName = CONFIG.DEFAULT_SHEET_NAME) {
  const spreadsheet = getSpreadsheet(sheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found in spreadsheet ${sheetId}`);
  }
  return sheet;
}

/**
 * Get the main spreadsheet ID
 * Uses PropertiesService (secure) if available, falls back to CONFIG
 * @return {string} The main spreadsheet ID
 */
function getMainSheetId() {
  // Try PropertiesService first (most secure)
  const properties = PropertiesService.getScriptProperties();
  const storedId = properties.getProperty('SPREADSHEET_ID_MAIN');
  if (storedId) {
    return storedId;
  }
  
  // Fallback to CONFIG (for local development)
  return CONFIG.SPREADSHEET_IDS.main;
}

/**
 * Get the main spreadsheet object
 * @return {Spreadsheet} The main spreadsheet
 */
function getMainSpreadsheet() {
  return getSpreadsheet(getMainSheetId());
}

/**
 * Setup Sheet IDs in PropertiesService (recommended for production)
 * Run this ONCE to store IDs securely in script properties
 * 
 * @param {Object} sheetIds - Object with sheet ID mappings
 * Example: { main: '1abc123...', backup: '1def456...' }
 */
function setupSheetIds(sheetIds = null) {
  const properties = PropertiesService.getScriptProperties();
  
  if (sheetIds) {
    // Store provided IDs
    Object.keys(sheetIds).forEach(key => {
      properties.setProperty(`SPREADSHEET_ID_${key.toUpperCase()}`, sheetIds[key]);
      Logger.log(`Stored SPREADSHEET_ID_${key.toUpperCase()}`);
    });
  } else {
    // Store from CONFIG (for initial setup)
    Object.keys(CONFIG.SPREADSHEET_IDS).forEach(key => {
      const id = CONFIG.SPREADSHEET_IDS[key];
      if (id && id !== 'YOUR_MAIN_SHEET_ID_HERE') {
        properties.setProperty(`SPREADSHEET_ID_${key.toUpperCase()}`, id);
        Logger.log(`Stored SPREADSHEET_ID_${key.toUpperCase()}`);
      }
    });
  }
  
  Logger.log('✅ Sheet IDs stored in PropertiesService');
  Logger.log('⚠️  You can now remove IDs from Config.gs for security');
}
