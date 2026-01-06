/**
 * Google Slides Export Project
 * Exports data from Google Sheets to Google Slides presentations
 */

/**
 * Export sheet data to a new Google Slides presentation
 * @param {string} sheetId - The spreadsheet ID
 * @param {string} sheetName - The sheet name (optional, uses active sheet if not provided)
 * @param {string} presentationTitle - Title for the new presentation
 * @return {string} The ID of the created presentation
 */
function exportSheetToSlides(sheetId, sheetName = null, presentationTitle = 'Exported Presentation') {
  try {
    // Get sheet data
    const sheet = getSheetByName(sheetId, sheetName);
    const data = getAllSheetData(sheet);
    
    // Create new presentation
    const presentation = SlidesApp.create(presentationTitle);
    const presentationId = presentation.getId();
    
    // Add title slide
    addTitleSlide(presentation, presentationTitle, sheetName || 'Sheet Data');
    
    // Add data slides
    addDataSlides(presentation, data);
    
    Logger.log(`Created presentation: ${presentationId}`);
    Logger.log(`View at: https://docs.google.com/presentation/d/${presentationId}/edit`);
    
    return presentationId;
  } catch (error) {
    Logger.log(`Error exporting to slides: ${error.message}`);
    throw error;
  }
}

/**
 * Export specific range to slides
 * @param {string} sheetId - The spreadsheet ID
 * @param {string} sheetName - The sheet name
 * @param {string} rangeA1 - A1 notation range (e.g., 'A1:D10')
 * @param {string} presentationTitle - Title for the presentation
 * @return {string} The ID of the created presentation
 */
function exportRangeToSlides(sheetId, sheetName, rangeA1, presentationTitle = 'Exported Range') {
  try {
    const sheet = getSheetByName(sheetId, sheetName);
    const range = sheet.getRange(rangeA1);
    const values = range.getValues();
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();
    
    // Create presentation
    const presentation = SlidesApp.create(presentationTitle);
    const presentationId = presentation.getId();
    
    // Add title slide
    addTitleSlide(presentation, presentationTitle, `Range: ${rangeA1}`);
    
    // Add table slide
    addTableSlide(presentation, values, numRows, numCols);
    
    Logger.log(`Created presentation: ${presentationId}`);
    return presentationId;
  } catch (error) {
    Logger.log(`Error exporting range: ${error.message}`);
    throw error;
  }
}

/**
 * Add data to existing presentation
 * @param {string} presentationId - The presentation ID
 * @param {string} sheetId - The spreadsheet ID
 * @param {string} sheetName - The sheet name
 */
function addSheetDataToPresentation(presentationId, sheetId, sheetName = null) {
  try {
    const presentation = SlidesApp.openById(presentationId);
    const sheet = getSheetByName(sheetId, sheetName);
    const data = getAllSheetData(sheet);
    
    addDataSlides(presentation, data);
    
    Logger.log(`Added data to presentation: ${presentationId}`);
  } catch (error) {
    Logger.log(`Error adding data to presentation: ${error.message}`);
    throw error;
  }
}

/**
 * Helper: Get all sheet data
 */
function getAllSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow === 0 || lastCol === 0) {
    return { headers: [], rows: [] };
  }
  
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  const values = range.getValues();
  
  return {
    headers: values[0] || [],
    rows: values.slice(1) || []
  };
}

/**
 * Helper: Add title slide
 */
function addTitleSlide(presentation, title, subtitle) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE);
  slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE).asShape().getText().setText(title);
  
  if (subtitle) {
    const subtitleShape = slide.getPlaceholder(SlidesApp.PlaceholderType.SUBTITLE);
    if (subtitleShape) {
      subtitleShape.asShape().getText().setText(subtitle);
    }
  }
}

/**
 * Helper: Add data slides
 */
function addDataSlides(presentation, data) {
  const { headers, rows } = data;
  
  if (rows.length === 0) {
    Logger.log('No data rows to add');
    return;
  }
  
  // Create table slide for headers and first few rows
  const rowsPerSlide = 10; // Adjust as needed
  const totalSlides = Math.ceil(rows.length / rowsPerSlide);
  
  for (let i = 0; i < totalSlides; i++) {
    const startRow = i * rowsPerSlide;
    const endRow = Math.min(startRow + rowsPerSlide, rows.length);
    const slideRows = rows.slice(startRow, endRow);
    
    // Create slide with title and body layout
    const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
    
    // Set title
    const titleShape = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
    if (titleShape) {
      titleShape.asShape().getText().setText(
        `Data (${startRow + 1}-${endRow} of ${rows.length})`
      );
    }
    
    // Add table
    addTableToSlide(slide, headers, slideRows);
  }
}

/**
 * Helper: Add table to slide
 */
function addTableSlide(presentation, values, numRows, numCols) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  
  // Create table
  const table = slide.insertTable(numRows, numCols);
  
  // Populate table
  for (let row = 0; row < numRows; row++) {
    for (let col = 0; col < numCols; col++) {
      const cell = table.getCell(row, col);
      cell.getText().setText(String(values[row][col] || ''));
      
      // Format header row
      if (row === 0) {
        cell.getText().getTextStyle().setBold(true);
        cell.getFill().setSolidFill('#4285f4');
        cell.getText().getTextStyle().setForegroundColor('#ffffff');
      }
    }
  }
  
  // Position table (center it)
  const pageWidth = presentation.getPageWidth();
  const pageHeight = presentation.getPageHeight();
  const tableWidth = table.getWidth();
  const tableHeight = table.getHeight();
  
  table.setLeft((pageWidth - tableWidth) / 2);
  table.setTop((pageHeight - tableHeight) / 2);
}

/**
 * Helper: Add table to existing slide
 */
function addTableToSlide(slide, headers, rows) {
  const numRows = rows.length + 1; // +1 for header
  const numCols = headers.length;
  
  if (numCols === 0) return;
  
  // Create table
  const table = slide.insertTable(numRows, numCols);
  
  // Add headers
  for (let col = 0; col < numCols; col++) {
    const cell = table.getCell(0, col);
    cell.getText().setText(String(headers[col] || ''));
    cell.getText().getTextStyle().setBold(true);
    cell.getFill().setSolidFill('#4285f4');
    cell.getText().getTextStyle().setForegroundColor('#ffffff');
  }
  
  // Add data rows
  for (let row = 0; row < rows.length; row++) {
    for (let col = 0; col < numCols; col++) {
      const cell = table.getCell(row + 1, col);
      cell.getText().setText(String(rows[row][col] || ''));
    }
  }
  
  // Position table in body area
  const bodyShape = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
  if (bodyShape) {
    const bodyLeft = bodyShape.getLeft();
    const bodyTop = bodyShape.getTop();
    table.setLeft(bodyLeft);
    table.setTop(bodyTop);
  }
}

/**
 * Helper: Get sheet by name or active sheet
 */
function getSheetByName(sheetId, sheetName) {
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  return sheetName 
    ? spreadsheet.getSheetByName(sheetName) 
    : spreadsheet.getActiveSheet();
}
