/**
 * Shared Configuration
 * Store your Sheet IDs and other config here
 * This file can be copied to each project or used as a library
 */

const CONFIG = {
  // Add your Sheet IDs here
  // Get from URL: https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit
  SPREADSHEET_IDS: {
    'main': '1MmoIyz2t9WlzBtcuywhUOExzejHNbx2Vs5cSwXgszeg',
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
 * @return {string} The main spreadsheet ID
 */
function getMainSheetId() {
  return CONFIG.SPREADSHEET_IDS.main;
}

/**
 * Get the main spreadsheet object
 * @return {Spreadsheet} The main spreadsheet
 */
function getMainSpreadsheet() {
  return getSpreadsheet(getMainSheetId());
}
