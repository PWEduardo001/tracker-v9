/**
 * MAIN CODE - CACHE REMOVED
 * System now reads Service Settings directly on every form submission
 */

/**
 * Creates a custom menu in Google Sheets for easy access
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Sunday Service System')
      .addSubMenu(ui.createMenu('⚙️ Setup & Configuration')
          .addItem('📋 Install All Triggers', 'installAllTriggers')
          .addItem('🔧 Install Form Trigger Only', 'installFormTrigger')
          .addItem('📅 Install Monday Sync Trigger', 'installMondaySyncTrigger')
          .addSeparator()
          .addItem('🏢 Setup All Communities', 'setupAllCommunities')
          .addItem('🏛️ Setup Individual Community', 'showSetupMenu')
          .addItem('📊 Setup Master Sheet', 'setupMasterSheet'))
      .addSubMenu(ui.createMenu('🔄 Weekly Sync Management')
          .addItem('▶️ Sync Now (Manual)', 'syncNow')
          .addItem('📊 View Sync Status', 'viewSyncStatus')
          .addItem('🧪 Test Sync System', 'testSyncSystem'))
      .addSubMenu(ui.createMenu('🧪 Testing')
          .addItem('✉️ Test Form Handler', 'testFormHandler')
          .addItem('🔧 Test Email System', 'testEmailSystem')
          .addItem('📝 View System Configuration', 'viewConfiguration'))
      .addSubMenu(ui.createMenu('📊 Data Management')
          .addItem('🎲 Add Test Data', 'addTestData')
          .addItem('🗑️ Clear All Test Data', 'clearAllTestData'))
      .addToUi();
}

// ==================== MASTER SHEET SETUP ====================

/**
 * Set up the master sheet
 */
function setupMasterSheet() {
  try {
    logSeparator('SETTING UP MASTER SHEET');
    
    var masterSpreadsheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
    Logger.log('Connected to master sheet: ' + masterSpreadsheet.getName());
    Logger.log('Sheet URL: ' + masterSpreadsheet.getUrl());
    Logger.log('');
    
    // Check if registration sheet exists (created by Google Form)
    var registrationSheet = masterSpreadsheet.getSheetByName('registration');
    if (!registrationSheet) {
      Logger.log('⚠️ WARNING: "registration" sheet not found');
      Logger.log('   Make sure your Google Form is linked to this spreadsheet.');
      Logger.log('');
    } else {
      Logger.log('✓ "registration" sheet found');
      Logger.log('');
    }
    
    // STEP 1: Create community tabs for all 8 communities
    Logger.log('STEP 1: Creating community tabs...');
    var communities = getCommunityNames();
    var tabsCreated = 0;
    
    for (var i = 0; i < communities.length; i++) {
      var communityName = communities[i];
      Logger.log('  Processing: ' + communityName);
      
      try {
        getOrCreateMasterCommunityTab(masterSpreadsheet, communityName);
        tabsCreated++;
        Logger.log('  ✓ ' + communityName + ' ready');
      } catch (error) {
        Logger.log('  ✗ Error: ' + error.toString());
      }
    }
    
    Logger.log('✓ Processed ' + tabsCreated + ' community tabs');
    Logger.log('');
    
    // STEP 2: Create summary tabs
    Logger.log('STEP 2: Creating summary tabs...');
    
    try {
      createOverallSummaryTab(masterSpreadsheet);
      Logger.log('✓ Overall Summary tab created');
    } catch (error) {
      Logger.log('✗ Error creating Overall Summary: ' + error.toString());
    }
    
    try {
      createCommunityComparisonTab(masterSpreadsheet);
      Logger.log('✓ Community Comparison tab created');
    } catch (error) {
      Logger.log('✗ Error creating Community Comparison: ' + error.toString());
    }
    
    try {
      createWeeklyTrendsTab(masterSpreadsheet);
      Logger.log('✓ Weekly Trends tab created');
    } catch (error) {
      Logger.log('✗ Error creating Weekly Trends: ' + error.toString());
    }
    
    Logger.log('');
    
    logSeparator('MASTER SHEET SETUP COMPLETE');
    Logger.log('✓ ' + tabsCreated + ' community tabs (7 columns)');
    Logger.log('✓ 3 summary tabs');
    Logger.log('');
    Logger.log('Sheet URL: ' + masterSpreadsheet.getUrl());
    logSeparator();
    
    return { success: true, url: masterSpreadsheet.getUrl() };
    
  } catch (error) {
    logSeparator('MASTER SHEET SETUP FAILED');
    Logger.log('Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    logSeparator();
    
    return { success: false, error: error.toString() };
  }
}

// ==================== TRIGGER INSTALLATION ====================

/**
 * Install all triggers at once (Form + Monday Sync only)
 */
function installAllTriggers() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Install All Triggers',
    'This will install:\n\n' +
    '1. Form Submit Trigger (instant emails)\n' +
    '2. Monday Sync Trigger (6 AM)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  logSeparator('INSTALLING ALL TRIGGERS');
  
  var results = {
    form: false,
    monday: false
  };
  
  try {
    installFormTrigger();
    results.form = true;
    Logger.log('✓ Form trigger installed');
  } catch (error) {
    Logger.log('✗ Form trigger failed: ' + error.toString());
  }
  
  try {
    installMondaySyncTrigger();
    results.monday = true;
    Logger.log('✓ Monday sync trigger installed');
  } catch (error) {
    Logger.log('✗ Monday sync trigger failed: ' + error.toString());
  }
  
  logSeparator('TRIGGER INSTALLATION COMPLETE');
  
  var successCount = (results.form ? 1 : 0) + (results.monday ? 1 : 0);
  
  ui.alert(
    'Trigger Installation Complete',
    'Successfully installed ' + successCount + ' of 2 triggers:\n\n' +
    '1. Form Submit: ' + (results.form ? '✓' : '✗') + '\n' +
    '2. Monday Sync: ' + (results.monday ? '✓' : '✗') + '\n\n' +
    'Check the Execution log for details.',
    ui.ButtonSet.OK
  );
}

/**
 * Install form submission trigger
 */
function installFormTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create();
  
  Logger.log('✓ Form submission trigger installed');
}

/**
 * Install Monday sync trigger (6 AM)
 */
function installMondaySyncTrigger() {
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
  
  Logger.log('✓ Monday sync trigger installed (6 AM)');
}

// ==================== ALIAS FUNCTIONS FOR MENU ====================

function setupMondayMorningSyncTrigger() {
  installMondaySyncTrigger();
}

function syncNow() {
  mondayMorningSync();
}

function testSyncSystem() {
  testMondaySync();
}

// ==================== CONFIGURATION VIEWING ====================

/**
 * View system configuration
 */
function viewConfiguration() {
  logSeparator('SYSTEM CONFIGURATION');
  
  Logger.log('Master Sheet ID: ' + MASTER_SHEET_ID);
  Logger.log('Data Structure: 7 columns (Type removed)');
  Logger.log('Service Settings: Direct lookup (no cache)');
  Logger.log('');
  
  var communities = getCommunityNames();
  Logger.log('Communities (' + communities.length + '):');
  for (var i = 0; i < communities.length; i++) {
    Logger.log('  ' + (i + 1) + '. ' + communities[i]);
  }
  
  Logger.log('');
  var sundayInfo = getNextSunday();
  Logger.log('Next Sunday: ' + sundayInfo.formatted);
  
  Logger.log('');
  Logger.log('Installed Triggers:');
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    Logger.log('  - ' + triggers[i].getHandlerFunction());
  }
  
  logSeparator();
}

/**
 * Show setup menu for individual communities
 */
function showSetupMenu() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Setup Individual Community',
    'Enter community name (e.g., "Belvedere Family Church"):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    var communityName = response.getResponseText().trim();
    
    if (isValidCommunity(communityName)) {
      var result = setupCommunity(communityName);
      
      if (result.success) {
        ui.alert(
          'Setup Complete',
          'Successfully set up ' + communityName + '\n\n' +
          '✓ All 5 tabs created\n' +
          '✓ Service Settings tab created\n\n' +
          'Check execution log for details.',
          ui.ButtonSet.OK
        );
      } else {
        ui.alert(
          'Setup Failed',
          'Error: ' + result.message + '\n\n' +
          'Check execution log for details.',
          ui.ButtonSet.OK
        );
      }
    } else {
      ui.alert(
        'Invalid Community',
        'Community "' + communityName + '" not found.\n\n' +
        'Valid communities:\n' +
        getCommunityNames().join('\n'),
        ui.ButtonSet.OK
      );
    }
  }
}
