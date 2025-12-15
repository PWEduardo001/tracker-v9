/**
 * FORM COLUMN DEBUGGER
 * Use this to see the EXACT column names from your Google Form
 * This will help fix the parsing issue
 */

/**
 * Debug: Show all form response columns
 * Run this AFTER someone submits the form
 */
function debugFormColumns() {
  try {
    logSeparator('FORM COLUMNS DEBUGGER');
    
    var masterSpreadsheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
    var registrationSheet = masterSpreadsheet.getSheetByName('registration');
    
    if (!registrationSheet) {
      Logger.log('ERROR: "registration" sheet not found');
      Logger.log('Make sure your Google Form is linked to the master spreadsheet');
      return;
    }
    
    // Get headers (column names)
    var headers = registrationSheet.getRange(1, 1, 1, registrationSheet.getLastColumn()).getValues()[0];
    
    Logger.log('TOTAL COLUMNS: ' + headers.length);
    Logger.log('');
    Logger.log('COLUMN NAMES (exactly as they appear in the form):');
    Logger.log('═══════════════════════════════════════════════════');
    
    for (var i = 0; i < headers.length; i++) {
      var columnLetter = String.fromCharCode(65 + i); // A, B, C, etc.
      Logger.log('Column ' + columnLetter + ' (index ' + i + '): "' + headers[i] + '"');
    }
    
    Logger.log('');
    Logger.log('═══════════════════════════════════════════════════');
    Logger.log('');
    
    // Get last row of data as example
    var lastRow = registrationSheet.getLastRow();
    if (lastRow > 1) {
      Logger.log('EXAMPLE DATA (last submission):');
      Logger.log('═══════════════════════════════════════════════════');
      
      var exampleData = registrationSheet.getRange(lastRow, 1, 1, headers.length).getValues()[0];
      
      for (var i = 0; i < headers.length; i++) {
        var columnLetter = String.fromCharCode(65 + i);
        Logger.log('Column ' + columnLetter + ' - ' + headers[i] + ':');
        Logger.log('  Value: "' + exampleData[i] + '"');
        Logger.log('');
      }
    }
    
    logSeparator('DEBUG COMPLETE');
    Logger.log('');
    Logger.log('NEXT STEPS:');
    Logger.log('1. Copy the EXACT column names from above');
    Logger.log('2. Update parseGoogleFormSubmission() function');
    Logger.log('3. Match getValue() calls to exact column names');
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
  }
}

/**
 * Test parsing with actual form data
 * Run this to see how many people are being parsed
 */
function testFormParsing() {
  try {
    logSeparator('TESTING FORM PARSING');
    
    var masterSpreadsheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
    var registrationSheet = masterSpreadsheet.getSheetByName('registration');
    
    if (!registrationSheet) {
      Logger.log('ERROR: "registration" sheet not found');
      return;
    }
    
    var lastRow = registrationSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('No data in registration sheet');
      return;
    }
    
    Logger.log('Testing with last submission (row ' + lastRow + ')...');
    Logger.log('');
    
    // Get the data
    var headers = registrationSheet.getRange(1, 1, 1, registrationSheet.getLastColumn()).getValues()[0];
    var data = registrationSheet.getRange(lastRow, 1, 1, registrationSheet.getLastColumn()).getValues()[0];
    
    // Create namedValues format (like form submission)
    var namedValues = {};
    for (var i = 0; i < headers.length; i++) {
      namedValues[headers[i]] = [data[i]];
    }
    
    // Create mock event object
    var mockEvent = {
      namedValues: namedValues
    };
    
    // Parse it
    var formData = parseGoogleFormSubmission(mockEvent);
    
    Logger.log('PARSING RESULTS:');
    Logger.log('═══════════════════════════════════════════════════');
    Logger.log('Community: ' + formData.community);
    Logger.log('Email: ' + formData.email);
    Logger.log('Number of registrants: ' + formData.registrants.length);
    Logger.log('');
    Logger.log('Registrants:');
    
    for (var i = 0; i < formData.registrants.length; i++) {
      var person = formData.registrants[i];
      Logger.log('  ' + (i + 1) + '. ' + person.firstName + ' ' + person.lastName + ' (' + person.type + ')');
    }
    
    Logger.log('');
    Logger.log('═══════════════════════════════════════════════════');
    
    if (formData.registrants.length === 0) {
      Logger.log('');
      Logger.log('⚠️ WARNING: No registrants parsed!');
      Logger.log('This means the column names don\'t match.');
      Logger.log('Run debugFormColumns() to see exact column names.');
    }
    
    logSeparator('TEST COMPLETE');
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
  }
}

/**
 * Show what the form expects vs what it's getting
 */
function showExpectedColumns() {
  logSeparator('EXPECTED FORM COLUMNS');
  
  Logger.log('The parseGoogleFormSubmission() function expects these EXACT column names:');
  Logger.log('');
  
  Logger.log('REQUIRED COLUMNS:');
  Logger.log('  1. "Which community are you registering for?"');
  Logger.log('  2. "Email Address"');
  Logger.log('  3. "First Name"');
  Logger.log('  4. "Last Name"');
  Logger.log('  5. "I am registering as"');
  Logger.log('');
  
  Logger.log('OPTIONAL COLUMNS (for additional people):');
  Logger.log('  6. "Will anyone else be joining with you?" → Answer: "Yes" or "No"');
  Logger.log('  7. "#2 First Name"');
  Logger.log('  8. "#2 Last Name"');
  Logger.log('  9. "#2 I am registering as"');
  Logger.log(' 10. "#2 Add another person?" → Answer: "Yes" or "No"');
  Logger.log(' 11. "#3 First Name"');
  Logger.log(' 12. "#3 Last Name"');
  Logger.log(' 13. "#3 I am registering as"');
  Logger.log(' 14. "#3 Add another person?" → Answer: "Yes" or "No"');
  Logger.log(' 15. "#4 First Name"');
  Logger.log(' 16. "#4 Last Name"');
  Logger.log(' 17. "#4 I am registering as"');
  Logger.log(' 18. "#4 Add another person?" → Answer: "Yes" or "No"');
  Logger.log(' 19. "#5 First Name"');
  Logger.log(' 20. "#5 Last Name"');
  Logger.log(' 21. "#5 I am registering as"');
  Logger.log('');
  
  Logger.log('NOTE: Column names must match EXACTLY (including spaces, capitalization, punctuation)');
  Logger.log('');
  Logger.log('To see your actual column names:');
  Logger.log('  Run: debugFormColumns()');
  
  logSeparator();
}

function removeTypeColumnOnce() {
  var communities = getCommunityNames();
  var fixed = 0;
  
  // Fix community sheets
  for (var i = 0; i < communities.length; i++) {
    var sheetId = getCommunitySheetId(communities[i]);
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('Registrations');
    if (sheet) {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      for (var j = 0; j < headers.length; j++) {
        if (headers[j].toString().toLowerCase() === 'type') {
          sheet.deleteColumn(j + 1);
          fixed++;
          break;
        }
      }
    }
  }
  
  // Fix master sheet
  var master = SpreadsheetApp.openById(MASTER_SHEET_ID);
  for (var i = 0; i < communities.length; i++) {
    var sheet = master.getSheetByName(communities[i]);
    if (sheet) {
      var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      for (var j = 0; j < headers.length; j++) {
        if (headers[j].toString().toLowerCase() === 'type') {
          sheet.deleteColumn(j + 1);
          fixed++;
          break;
        }
      }
    }
  }
  
  Browser.msgBox('✓ Fixed ' + fixed + ' sheets! Type column removed from all existing data.');
}
