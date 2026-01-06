/**
 * Migration Helper: Move Sheet IDs to PropertiesService
 * 
 * Run this ONCE to migrate your existing Sheet ID from Config.gs to PropertiesService
 * After running, you can safely remove the ID from Config.gs
 * 
 * INSTRUCTIONS:
 * 1. Open Apps Script editor
 * 2. Select this function: migrateSheetIdToPropertiesService
 * 3. Click Run (▶️)
 * 4. Check the logs to confirm the ID was stored
 * 5. Update Config.gs to use empty string
 * 6. Commit changes to GitHub
 */

/**
 * Migrate existing Sheet ID from Config.gs to PropertiesService
 * Run this function ONCE in Apps Script editor
 */
function migrateSheetIdToPropertiesService() {
  try {
    const properties = PropertiesService.getScriptProperties();
    
    // Your existing Sheet ID (from Config.gs)
    const existingId = 'YOUR_SHEET_ID_HERE';
    
    // Store in PropertiesService
    properties.setProperty('SPREADSHEET_ID_MAIN', existingId);
    
    Logger.log('✅ Sheet ID migrated to PropertiesService');
    Logger.log(`Stored: SPREADSHEET_ID_MAIN = ${existingId}`);
    Logger.log('');
    Logger.log('Next steps:');
    Logger.log('1. Verify the ID is stored: Project Settings → Script properties');
    Logger.log('2. Update Config.gs to use empty string: main: ""');
    Logger.log('3. Test that getMainSheetId() still works');
    Logger.log('4. Commit changes to GitHub (ID no longer in code)');
    
    return true;
  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    return false;
  }
}
