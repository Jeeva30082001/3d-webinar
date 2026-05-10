function doGet(e) {
  return handleFormSubmission(e);
}

function doPost(e) {
  return handleFormSubmission(e);
}

function handleFormSubmission(e) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const spreadsheetId = scriptProperties.getProperty('SPREADSHEET_ID');
    
    if (!spreadsheetId) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'Spreadsheet ID not configured'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // Get or create "Responses" sheet
    let sheet = spreadsheet.getSheetByName('Responses');
    if (!sheet) {
      sheet = spreadsheet.insertSheet('Responses');
      // Add headers with bold formatting
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'Phone']);
      const headerRange = sheet.getRange(1, 1, 1, 4);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e8e8e8');
    }
    
    // Get form data from query parameters
    const params = e.parameter;
    const name = params.name || '';
    const email = params.email || '';
    const phone = params.phone || '';
    
    // Validate that at least one field is filled
    if (!name && !email && !phone) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'No data received'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Append row with timestamp
    const timestamp = new Date().toLocaleString('en-IN', {
      timeZone: 'Asia/Kolkata',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    });
    
    sheet.appendRow([timestamp, name, email, phone]);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Data saved successfully'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Helper function to setup - run this once to store your spreadsheet ID
function setupSpreadsheetId() {
  // Replace with your actual Google Sheet ID
  const spreadsheetId = 'YOUR_SPREADSHEET_ID_HERE';
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
  Logger.log('Spreadsheet ID saved: ' + spreadsheetId);
}
