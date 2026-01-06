/**
 * Helper function to create and populate the Format_Config sheet
 * Run this once to set up the Format_Config sheet with headers and example data
 */
function setupFormatConfigSheet() {
  try {
    const sheetId = CONFIG.SPREADSHEET_IDS.main;
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    
    // Check if sheet already exists
    let configSheet = spreadsheet.getSheetByName('Format_Config');
    
    if (configSheet) {
      Logger.log('Format_Config sheet already exists');
      // Clear existing data (optional - comment out if you want to keep existing data)
      // configSheet.clear();
    } else {
      // Create new sheet
      configSheet = spreadsheet.insertSheet('Format_Config');
      Logger.log('Created Format_Config sheet');
    }
    
    // Set headers
    const headers = [['ConfigKey', 'SourceSheet', 'SourceRanges', 'DestinationCell']];
    configSheet.getRange(1, 1, 1, 4).setValues(headers);
    
    // Format header row
    const headerRange = configSheet.getRange(1, 1, 1, 4);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // Add example data
    const exampleData = [
      ['row8', 'Performance and Progress', 'B8, C8', 'D8'],
      ['row9', 'Performance and Progress', 'B9, C9', 'D9']
    ];
    
    if (configSheet.getLastRow() === 1) {
      // Only add examples if sheet is empty (just headers)
      configSheet.getRange(2, 1, exampleData.length, 4).setValues(exampleData);
      Logger.log('Added example configurations');
    }
    
    // Auto-resize columns
    configSheet.autoResizeColumns(1, 4);
    
    Logger.log('✅ Format_Config sheet is ready!');
    Logger.log('Edit the sheet to add your configurations, then run concatenateCellsWithFormatting()');
    
    // Return the sheet URL
    const url = `https://docs.google.com/spreadsheets/d/${sheetId}/edit#gid=${configSheet.getSheetId()}`;
    Logger.log(`Sheet URL: ${url}`);
    
    return configSheet;
    
  } catch (error) {
    Logger.log(`❌ Error setting up Format_Config sheet: ${error.message}`);
    throw error;
  }
}
