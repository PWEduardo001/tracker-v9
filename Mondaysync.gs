/**
 * MONDAY MORNING SYNC - UPDATED FOR YEAR-LONG SERVICE SETTINGS
 * Distributes Google Form registration data to master and community sheets
 * UPDATED: Type column completely removed - no more Member/Guest distinction
 * 
 * WHAT THIS DOES:
 * 1. Reads unsynced responses from "registration" sheet (Google Form responses)
 * 2. Parses 1-5 people per form submission
 * 3. Distributes to master sheet community tabs (7 columns)
 * 4. Distributes to individual community sheets (7 columns)
 * 5. Updates all dashboard tabs
 * 6. Marks rows as synced to prevent duplicates
 */

/**
 * Main Monday morning sync function
 */
function mondayMorningSync() {
  try {
    logSeparator('MONDAY MORNING SYNC STARTED');
    var startTime = new Date().getTime();
    
    var masterSpreadsheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
    Logger.log('✓ Connected to master spreadsheet');
    
    var registrationSheet = masterSpreadsheet.getSheetByName('registration');
    
    if (!registrationSheet) {
      Logger.log('ERROR: "registration" sheet not found');
      Logger.log('Make sure your Google Form is linked to this spreadsheet');
      logSeparator('SYNC FAILED');
      return;
    }
    
    Logger.log('✓ Found registration sheet');
    
    ensureSyncedColumn(registrationSheet);
    
    var allData = registrationSheet.getDataRange().getValues();
    
    if (allData.length <= 1) {
      Logger.log('No data in registration sheet (only headers)');
      logSeparator('SYNC COMPLETE - NO DATA');
      return;
    }
    
    Logger.log('Total rows in registration sheet: ' + (allData.length - 1));
    
    var registrations = parseFormResponses(registrationSheet);
    
    if (registrations.length === 0) {
      Logger.log('No unsynced registrations found');
      logSeparator('SYNC COMPLETE - ALL UP TO DATE');
      return;
    }
    
    Logger.log('Found ' + registrations.length + ' unsynced registration(s)');
    
    Logger.log('');
    Logger.log('STEP 1: Syncing to master sheet community tabs...');
    syncToMasterCommunityTabs(registrations);
    
    Logger.log('');
    Logger.log('STEP 2: Syncing to individual community sheets...');
    syncToCommunitySheets(registrations);
    
    Logger.log('');
    Logger.log('STEP 3: Marking rows as synced...');
    markRowsAsSynced(registrationSheet, registrations);
    
    var endTime = new Date().getTime();
    var duration = (endTime - startTime) / 1000;
    
    logSeparator('SYNC COMPLETE');
    Logger.log('Total registrations processed: ' + registrations.length);
    Logger.log('Time taken: ' + duration + ' seconds');
    logSeparator();
    
  } catch (error) {
    Logger.log('ERROR in mondayMorningSync: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    logSeparator('SYNC FAILED');
    sendSyncErrorNotification(error);
  }
}

/**
 * Ensure the "Synced" column exists
 */
function ensureSyncedColumn(registrationSheet) {
  var headers = registrationSheet.getRange(1, 1, 1, registrationSheet.getLastColumn()).getValues()[0];
  
  var syncedColumnIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === 'Synced') {
      syncedColumnIndex = i;
      break;
    }
  }
  
  if (syncedColumnIndex === -1) {
    var lastColumn = registrationSheet.getLastColumn();
    registrationSheet.getRange(1, lastColumn + 1).setValue('Synced');
    
    registrationSheet.getRange(1, lastColumn + 1)
                     .setFontWeight('bold')
                     .setBackground(SETTINGS.COLORS.PRIMARY)
                     .setFontColor('#ffffff')
                     .setHorizontalAlignment('center');
    
    Logger.log('✓ Created "Synced" column');
  } else {
    Logger.log('✓ "Synced" column already exists');
  }
}

/**
 * Parse form responses into structured registration data
 * UPDATED: Type field removed - no more Member/Guest parsing
 */
function parseFormResponses(registrationSheet) {
  var data = registrationSheet.getDataRange().getValues();
  var headers = data[0];
  
  var syncedColIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === 'Synced') {
      syncedColIndex = i;
      break;
    }
  }
  
  var registrations = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Check if already synced
    if (syncedColIndex !== -1 && row[syncedColIndex]) {
      var syncedValue = row[syncedColIndex].toString().trim();
      if (syncedValue !== '' && syncedValue.toLowerCase().indexOf('yes') !== -1) {
        continue;
      }
    }
    
    // Parse this unsynced row - NO TYPE FIELD
    var timestamp = row[0];
    var email = row[1];
    var community = row[2];
    
    if (!community) {
      Logger.log('Warning: Row ' + (i + 1) + ' has no community, skipping');
      continue;
    }
    
    var registrants = [];
    
    // Person 1 (columns D, E) - NO TYPE
    var person1FirstName = row[3];
    var person1LastName = row[4];
    
    if (person1FirstName && person1LastName) {
      registrants.push({
        firstName: person1FirstName,
        lastName: person1LastName,
        email: email
      });
    }
    
    // Person 2 - NO TYPE
    if (row[5] === 'Yes') {
      var person2FirstName = row[6];
      var person2LastName = row[7];
      
      if (person2FirstName && person2LastName) {
        registrants.push({
          firstName: person2FirstName,
          lastName: person2LastName,
          email: email
        });
      }
    }
    
    // Person 3 - NO TYPE
    if (row[8] === 'Yes') {
      var person3FirstName = row[9];
      var person3LastName = row[10];
      
      if (person3FirstName && person3LastName) {
        registrants.push({
          firstName: person3FirstName,
          lastName: person3LastName,
          email: email
        });
      }
    }
    
    // Person 4 - NO TYPE
    if (row[11] === 'Yes') {
      var person4FirstName = row[12];
      var person4LastName = row[13];
      
      if (person4FirstName && person4LastName) {
        registrants.push({
          firstName: person4FirstName,
          lastName: person4LastName,
          email: email
        });
      }
    }
    
    // Person 5 - NO TYPE
    if (row[14] === 'Yes') {
      var person5FirstName = row[15];
      var person5LastName = row[16];
      
      if (person5FirstName && person5LastName) {
        registrants.push({
          firstName: person5FirstName,
          lastName: person5LastName,
          email: email
        });
      }
    }
    
    if (registrants.length > 0) {
      var serviceSettings = getNextSundayServiceSettings(community);
      var sundayDate = serviceSettings ? serviceSettings.sundayDate : getNextSunday().formatted;
      
      registrations.push({
        rowNumber: i + 1,
        timestamp: timestamp,
        community: community,
        registrants: registrants,
        sundayDate: sundayDate
      });
      
      Logger.log('Parsed row ' + (i + 1) + ': ' + community + ' - ' + registrants.length + ' people - ' + sundayDate);
    }
  }
  
  return registrations;
}

/**
 * Sync registrations to master sheet community tabs
 * UPDATED: 7 columns (no Type)
 */
function syncToMasterCommunityTabs(registrations) {
  var masterSpreadsheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var processed = 0;
  
  for (var i = 0; i < registrations.length; i++) {
    var reg = registrations[i];
    
    try {
      var communityTab = getOrCreateMasterCommunityTab(masterSpreadsheet, reg.community);
      
      for (var j = 0; j < reg.registrants.length; j++) {
        var person = reg.registrants[j];
        
        // 7 columns (no Type)
        communityTab.appendRow([
          reg.timestamp,
          reg.community,
          person.firstName,
          person.lastName,
          person.email,
          'Sunday Service',
          reg.sundayDate
        ]);
        
        processed++;
      }
      
      var lastRow = communityTab.getLastRow();
      if (lastRow > 2) {
        var dataRange = communityTab.getRange(3, 1, lastRow - 2, 7);
        dataRange.sort({column: 1, ascending: false});
      }
      
    } catch (error) {
      Logger.log('Error syncing to master sheet for ' + reg.community + ': ' + error.toString());
    }
  }
  
  Logger.log('✓ Synced ' + processed + ' registrants to master sheet community tabs');
}

/**
 * Sync registrations to individual community sheets
 * UPDATED: 7 columns (no Type)
 */
function syncToCommunitySheets(registrations) {
  var totalProcessed = 0;
  
  var byCommunity = {};
  for (var i = 0; i < registrations.length; i++) {
    var reg = registrations[i];
    if (!byCommunity[reg.community]) {
      byCommunity[reg.community] = [];
    }
    byCommunity[reg.community].push(reg);
  }
  
  var communities = Object.keys(byCommunity);
  for (var i = 0; i < communities.length; i++) {
    var communityName = communities[i];
    var communityRegs = byCommunity[communityName];
    
    Logger.log('Processing ' + communityName + ' (' + communityRegs.length + ' registration(s))...');
    
    try {
      var sheetId = getCommunitySheetId(communityName);
      if (!sheetId) {
        Logger.log('Warning: No sheet ID found for ' + communityName);
        continue;
      }
      
      var spreadsheet = SpreadsheetApp.openById(sheetId);
      var registrationsTab = getOrCreateCommunityRegistrationsTab(spreadsheet);
      
      var count = 0;
      for (var j = 0; j < communityRegs.length; j++) {
        var reg = communityRegs[j];
        
        for (var k = 0; k < reg.registrants.length; k++) {
          var person = reg.registrants[k];
          
          // 7 columns (no Type)
          registrationsTab.appendRow([
            reg.timestamp,
            reg.community,
            person.firstName,
            person.lastName,
            person.email,
            'Sunday Service',
            reg.sundayDate
          ]);
          
          count++;
          totalProcessed++;
        }
      }
      
      var lastRow = registrationsTab.getLastRow();
      if (lastRow > 1) {
        var dataRange = registrationsTab.getRange(2, 1, lastRow - 1, 7);
        dataRange.sort({column: 1, ascending: false});
      }
      
      Logger.log('  ✓ Added ' + count + ' registrants to ' + communityName);
      
      Logger.log('  Updating dashboards for ' + communityName + '...');
      updateCommunityDashboards(sheetId);
      Logger.log('  ✓ Dashboards updated');
      
    } catch (error) {
      Logger.log('  Error processing ' + communityName + ': ' + error.toString());
    }
  }
  
  Logger.log('✓ Synced ' + totalProcessed + ' registrants to community sheets');
}

/**
 * Mark rows as synced in the registration sheet
 */
function markRowsAsSynced(registrationSheet, registrations) {
  var headers = registrationSheet.getRange(1, 1, 1, registrationSheet.getLastColumn()).getValues()[0];
  
  var syncedColIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === 'Synced') {
      syncedColIndex = i + 1;
      break;
    }
  }
  
  if (syncedColIndex === -1) {
    Logger.log('Warning: Could not find Synced column');
    return;
  }
  
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  for (var i = 0; i < registrations.length; i++) {
    var rowNumber = registrations[i].rowNumber;
    registrationSheet.getRange(rowNumber, syncedColIndex).setValue('Yes - ' + timestamp);
  }
  
  Logger.log('✓ Marked ' + registrations.length + ' row(s) as synced');
}

/**
 * Send error notification email if sync fails
 */
function sendSyncErrorNotification(error) {
  try {
    var recipient = Session.getActiveUser().getEmail();
    var subject = '⚠️ Monday Sync Failed - Sunday Service Registration';
    
    var body = 
      'The Monday morning sync process failed.\n\n' +
      'Error: ' + error.toString() + '\n\n' +
      'Stack trace:\n' + error.stack + '\n\n' +
      'Time: ' + new Date().toString() + '\n\n' +
      'Please check the execution logs and fix the issue.\n' +
      'You can manually run the sync from: Menu → Weekly Sync Management → Sync Now';
    
    MailApp.sendEmail(recipient, subject, body);
    Logger.log('✓ Error notification sent to ' + recipient);
    
  } catch (emailError) {
    Logger.log('Could not send error notification email: ' + emailError.toString());
  }
}

function setupMondayMorningSyncTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'mondayMorningSync') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger('mondayMorningSync')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .create();
  
  Logger.log('✓ Monday morning sync trigger installed (6 AM every Monday)');
  
  Browser.msgBox(
    'Trigger Installed',
    'Monday morning sync trigger has been installed.\n\n' +
    'The sync will run automatically every Monday at 6:00 AM.\n\n' +
    'You can also run it manually from:\n' +
    'Menu → Weekly Sync Management → Sync Now',
    Browser.Buttons.OK
  );
}

function removeMondayMorningSyncTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'mondayMorningSync') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  
  Logger.log('Removed ' + removed + ' Monday sync trigger(s)');
  
  Browser.msgBox(
    'Trigger Removed',
    'Monday morning sync trigger has been removed.\n\n' +
    'Automatic syncing is now disabled.\n' +
    'You can still run manual syncs from the menu.',
    Browser.Buttons.OK
  );
}

function testMondaySync() {
  Logger.log('===== TESTING MONDAY SYNC =====');
  
  Browser.msgBox(
    'Test Monday Sync',
    'This will process any unsynced registrations in the "registration" sheet.\n\n' +
    'Check the execution log for results.',
    Browser.Buttons.OK
  );
  
  mondayMorningSync();
  
  Logger.log('===== TEST COMPLETE =====');
}

function viewSyncStatus() {
  try {
    var masterSpreadsheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
    var registrationSheet = masterSpreadsheet.getSheetByName('registration');
    
    if (!registrationSheet) {
      Browser.msgBox(
        'Status Check Failed',
        '"registration" sheet not found.\n\n' +
        'Make sure your Google Form is linked to this spreadsheet.',
        Browser.Buttons.OK
      );
      return;
    }
    
    var data = registrationSheet.getDataRange().getValues();
    var headers = data[0];
    
    var syncedColIndex = -1;
    for (var i = 0; i < headers.length; i++) {
      if (headers[i] === 'Synced') {
        syncedColIndex = i;
        break;
      }
    }
    
    var totalRows = data.length - 1;
    var syncedCount = 0;
    var unsyncedCount = 0;
    
    if (syncedColIndex !== -1) {
      for (var i = 1; i < data.length; i++) {
        var syncedValue = data[i][syncedColIndex];
        if (syncedValue && syncedValue.toString().trim() !== '' && 
            syncedValue.toString().toLowerCase().indexOf('yes') !== -1) {
          syncedCount++;
        } else {
          unsyncedCount++;
        }
      }
    } else {
      unsyncedCount = totalRows;
    }
    
    Logger.log('===== SYNC STATUS =====');
    Logger.log('Total registrations: ' + totalRows);
    Logger.log('Synced: ' + syncedCount);
    Logger.log('Unsynced: ' + unsyncedCount);
    Logger.log('=====================');
    
    Browser.msgBox(
      'Sync Status',
      'Total registrations: ' + totalRows + '\n' +
      'Synced: ' + syncedCount + '\n' +
      'Unsynced: ' + unsyncedCount + '\n\n' +
      (unsyncedCount > 0 ? 
        'You can sync now from:\nMenu → Weekly Sync Management → Sync Now' :
        'All registrations are synced!'),
      Browser.Buttons.OK
    );
    
  } catch (error) {
    Logger.log('Error checking sync status: ' + error.toString());
    Browser.msgBox(
      'Status Check Failed',
      'Error: ' + error.toString(),
      Browser.Buttons.OK
    );
  }
}
