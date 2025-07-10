/**
 * Custom functions for Google Sheets
 * 
 * These functions will be available in Google Sheets after deployment
 */

/**
 * Returns a greeting message
 * @param {string} name - The name to greet
 * @return {string} The greeting message
 * @customfunction
 */
function HELLO(name: string): string {
  return `Hello, ${name}!`;
}

/**
 * Multiplies a number by 2
 * @param {number} number - The number to multiply
 * @return {number} The result
 * @customfunction
 */
function MULTIPLY_BY_TWO(number: number): number {
  return number * 2;
}

/**
 * Concatenates text with a prefix
 * @param {string} prefix - The prefix to add
 * @param {string} text - The text to concatenate
 * @return {string} The concatenated result
 * @customfunction
 */
function CONCATENATE_WITH_PREFIX(prefix: string, text: string): string {
  return `${prefix}: ${text}`;
}

/**
 * Calculates the percentage of a value
 * @param {number} value - The value
 * @param {number} total - The total
 * @return {number} The percentage
 * @customfunction
 */
function CALCULATE_PERCENTAGE(value: number, total: number): number {
  if (total === 0) {
    return 0;
  }
  return (value / total) * 100;
}

/**
 * Formats a date as a readable string
 * @param {Date} date - The date to format
 * @return {string} The formatted date string
 * @customfunction
 */
function FORMAT_DATE(date: Date): string {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM dd, yyyy');
}

/**
 * Generates a random number between min and max
 * @param {number} min - The minimum value
 * @param {number} max - The maximum value
 * @return {number} A random number
 * @customfunction
 */
function RANDOM_BETWEEN(min: number, max: number): number {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

/**
 * Updates a specific cell with QuickBooks account data
 * @param {string} accountNumber - The QuickBooks account number
 * @param {number} accountValue - The account balance/value
 * @param {string} cellAddress - The cell address (e.g., "A1", "B2")
 * @param {string} spreadsheetId - The spreadsheet ID (optional, uses active if not provided)
 * @param {string} sheetName - The name of the sheet (optional)
 * @return {string} Success message
 * @customfunction
 */
function UPDATE_QB_ACCOUNT(accountNumber: string, accountValue: number, cellAddress: string, spreadsheetId?: string, sheetName?: string): string {
  try {
    // Use specific spreadsheet if ID provided, otherwise use active spreadsheet
    const spreadsheet = spreadsheetId ? 
      SpreadsheetApp.openById(spreadsheetId) : 
      SpreadsheetApp.getActiveSpreadsheet();
    
    const sheet = sheetName ? spreadsheet.getSheetByName(sheetName) : spreadsheet.getActiveSheet();
    
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }
    
    // Update the cell with the account value
    const range = sheet.getRange(cellAddress);
    range.setValue(accountValue);
    
    // Add a timestamp to track when this was last updated
    const timestampCell = sheet.getRange(cellAddress).offset(0, 1);
    timestampCell.setValue(new Date());
    
    return `Account ${accountNumber} updated: ${accountValue} at ${new Date().toLocaleString()}`;
  } catch (error) {
    console.error('Error updating QuickBooks account:', error);
    // Re-throw the error instead of returning an error message
    throw error;
  }
}

/**
 * Web App endpoint to receive QuickBooks data from external service
 * This function handles POST requests from the Rust service
 */
function doPost(e: GoogleAppsScript.Events.DoPost): GoogleAppsScript.Content.TextOutput {
  try {
    const data = JSON.parse(e.postData.contents);
    
    // Validate required fields
    if (!data.accountNumber || data.accountValue === undefined || !data.cellAddress) {
      throw new Error('Missing required fields: accountNumber, accountValue, cellAddress');
    }
    
    // Validate API key for security
    if (!data.apiKey || data.apiKey !== PropertiesService.getScriptProperties().getProperty('QB_API_KEY')) {
      throw new Error('Invalid API key');
    }
    
    // Update the QuickBooks account data
    const result = UPDATE_QB_ACCOUNT(
      data.accountNumber,
      data.accountValue,
      data.cellAddress,
      data.spreadsheetId,  // Pass spreadsheet ID if provided
      data.sheetName
    );
    
    // Log the update
    console.log(`QuickBooks sync: ${result}`);
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, message: result }))
      .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    console.error('Error in doPost:', error);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error instanceof Error ? error.message : String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Test function to verify Web App deployment
 */
function testWebAppEndpoint(): void {
  try {
    // Create a test POST request payload
    const testPayload = {
      accountNumber: 'TEST',
      accountValue: 999.99,
      spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId(),
      cellAddress: 'A1',
      sheetName: 'Sheet1',
      apiKey: PropertiesService.getScriptProperties().getProperty('QB_API_KEY')
    };
    
    console.log('Test payload:', JSON.stringify(testPayload, null, 2));
    
    // Simulate the doPost function
    const mockEvent = {
      postData: {
        contents: JSON.stringify(testPayload)
      }
    } as GoogleAppsScript.Events.DoPost;
    
    const response = doPost(mockEvent);
    console.log('Response:', response.getContent());
    
  } catch (error) {
    console.error('Test failed:', error);
  }
}

/**
 * Set up the API key for QuickBooks integration
 * Run this once to set up authentication and test the setup
 */
function setupQuickBooksIntegration(): string {
  try {
    // Generate a unique API key
    const apiKey = Utilities.getUuid();
    
    // Store the API key in script properties
    PropertiesService.getScriptProperties().setProperty('QB_API_KEY', apiKey);
    
    // Log setup completion
    console.log('QuickBooks integration setup completed!');
    console.log('API Key:', apiKey);
    console.log('Copy this API key to your Rust service configuration.');
    
    // Test the UPDATE_QB_ACCOUNT function
    const testResult = UPDATE_QB_ACCOUNT('TEST', 1234.56, 'A1');
    console.log('Test result:', testResult);
    
    return apiKey;
  } catch (error) {
    console.error('Setup failed:', error);
    throw error;
  }
}

/**
 * Get the current API key (for setup purposes)
 * @return {string} The current API key
 */
function getQuickBooksApiKey(): string {
  return PropertiesService.getScriptProperties().getProperty('QB_API_KEY') || 'Not set';
}

/**
 * Manually trigger QuickBooks account update (for testing)
 * @param {number} accountValue - The account balance to set
 * @param {string} cellAddress - The cell address to update
 * @return {string} Result message
 * @customfunction
 */
function TEST_QB_UPDATE(accountValue: number, cellAddress: string): string {
  return UPDATE_QB_ACCOUNT('9445', accountValue, cellAddress);
}

/**
 * Setup function to configure permissions and test access
 * Run this function first to set up all necessary permissions
 */
function setupPermissions(): void {
  try {
    console.log('🔧 Setting up permissions for QuickBooks integration...');
    
    // Test basic Google Sheets access
    console.log('📊 Testing Google Sheets access...');
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = activeSpreadsheet.getActiveSheet();
    
    console.log('✅ Active spreadsheet ID:', activeSpreadsheet.getId());
    console.log('✅ Active sheet name:', activeSheet.getName());
    
    // Test writing to a cell
    console.log('📝 Testing write access...');
    const testCell = activeSheet.getRange('A1');
    const originalValue = testCell.getValue();
    
    testCell.setValue('Permission Test - ' + new Date().toISOString());
    console.log('✅ Successfully wrote to cell A1');
    
    // Test reading from the cell
    const readValue = testCell.getValue();
    console.log('✅ Successfully read from cell A1:', readValue);
    
    // Restore original value if it wasn't empty
    if (originalValue !== '') {
      testCell.setValue(originalValue);
      console.log('✅ Restored original cell value');
    }
    
    // Test accessing script properties
    console.log('🔑 Testing script properties access...');
    const properties = PropertiesService.getScriptProperties();
    const testKey = 'PERMISSION_TEST';
    const testValue = 'SUCCESS_' + Date.now();
    
    properties.setProperty(testKey, testValue);
    const retrievedValue = properties.getProperty(testKey);
    
    if (retrievedValue === testValue) {
      console.log('✅ Script properties access working');
      properties.deleteProperty(testKey);
    } else {
      console.log('❌ Script properties access failed');
    }
    
    console.log('');
    console.log('🎉 Permission setup completed successfully!');
    console.log('📋 Next steps:');
    console.log('   1. Run setupQuickBooksIntegration() to generate API key');
    console.log('   2. Deploy as Web App with "Anyone" access');
    console.log('   3. Test with testWebAppEndpoint()');
    
  } catch (error) {
    console.error('❌ Permission setup failed:', error);
    console.log('');
    console.log('🔧 Troubleshooting:');
    console.log('   1. Make sure you have edit access to this spreadsheet');
    console.log('   2. Check that the Google Apps Script project has the necessary scopes');
    console.log('   3. Try running the script manually first');
    
    throw error;
  }
}

/**
 * Check current permissions and access levels
 */
function checkPermissions(): void {
  try {
    console.log('🔍 Checking current permissions...');
    
    // Check basic access
    const user = Session.getActiveUser();
    console.log('👤 Active user:', user.getEmail());
    
    // Check spreadsheet access
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    console.log('📊 Spreadsheet ID:', spreadsheet.getId());
    console.log('📊 Spreadsheet name:', spreadsheet.getName());
    console.log('📊 Spreadsheet URL:', spreadsheet.getUrl());
    
    // Check sheet access
    const sheets = spreadsheet.getSheets();
    console.log('📄 Available sheets:');
    sheets.forEach((sheet, index) => {
      console.log(`   ${index + 1}. ${sheet.getName()} (${sheet.getMaxRows()}x${sheet.getMaxColumns()})`);
    });
    
    // Check script properties
    const apiKey = PropertiesService.getScriptProperties().getProperty('QB_API_KEY');
    console.log('🔑 API Key configured:', apiKey ? 'YES' : 'NO');
    
    // Check if we can create a trigger (advanced permission)
    try {
      const triggers = ScriptApp.getProjectTriggers();
      console.log('⚡ Existing triggers:', triggers.length);
    } catch (e) {
      console.log('⚡ Trigger access: Limited (may need manual authorization)');
    }
    
    console.log('');
    console.log('✅ Permission check completed');
    
  } catch (error) {
    console.error('❌ Permission check failed:', error);
    throw error;
  }
}
