/**
 * Shared Utility Functions
 * Common functions used across multiple projects
 */

/**
 * Log with timestamp
 * @param {string} message - Message to log
 */
function log(message) {
  const timestamp = new Date().toISOString();
  Logger.log(`[${timestamp}] ${message}`);
  console.log(`[${timestamp}] ${message}`);
}

/**
 * Get data range from sheet
 * @param {Sheet} sheet - The sheet object
 * @return {Range} The data range
 */
function getDataRange(sheet) {
  return sheet.getDataRange();
}

/**
 * Get all values from sheet
 * @param {Sheet} sheet - The sheet object
 * @return {Array<Array>} 2D array of values
 */
function getAllValues(sheet) {
  return getDataRange(sheet).getValues();
}

/**
 * Get headers from sheet (first row)
 * @param {Sheet} sheet - The sheet object
 * @return {Array<string>} Array of header names
 */
function getHeaders(sheet) {
  const values = getAllValues(sheet);
  return values.length > 0 ? values[0] : [];
}

/**
 * Get data rows (excluding header)
 * @param {Sheet} sheet - The sheet object
 * @return {Array<Array>} 2D array of data rows
 */
function getDataRows(sheet) {
  const values = getAllValues(sheet);
  return values.length > 1 ? values.slice(1) : [];
}

/**
 * Find column index by header name
 * @param {Sheet} sheet - The sheet object
 * @param {string} headerName - The header name to find
 * @return {number} Column index (1-based) or -1 if not found
 */
function findColumnIndex(sheet, headerName) {
  const headers = getHeaders(sheet);
  const index = headers.indexOf(headerName);
  return index >= 0 ? index + 1 : -1;
}

/**
 * Validate sheet ID format
 * @param {string} sheetId - The sheet ID to validate
 * @return {boolean} True if valid format
 */
function isValidSheetId(sheetId) {
  return sheetId && typeof sheetId === 'string' && sheetId.length > 0;
}
