/**
 * MASTER SHEET OPERATIONS
 * Handles all master sheet updates and aggregation
 * UPDATED: Type column completely removed - no more Member/Guest distinction
 */

/**
 * Write registration to master sheet
 * This writes to the community-specific tab in the master sheet
 */
function writeToMasterSheet(communityName, registrationData) {
  try {
    Logger.log('Writing to master sheet for: ' + communityName);
    
    var masterSpreadsheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
    var communityTab = getOrCreateMasterCommunityTab(masterSpreadsheet, communityName);
    
    // Add each registrant to the community tab - 7 columns (no Type)
    var timestamp = new Date();
    for (var i = 0; i < registrationData.registrants.length; i++) {
      var person = registrationData.registrants[i];
      
      communityTab.appendRow([
        timestamp,
        communityName,
        person.firstName,
        person.lastName,
        person.email,
        registrationData.sessionInfo || 'Sunday Service',
        registrationData.sundayDate || ''
      ]);
    }
    
    // Sort by timestamp (newest first)
    var lastRow = communityTab.getLastRow();
    if (lastRow > 1) {
      var range = communityTab.getRange(2, 1, lastRow - 1, 7);
      range.sort({column: 1, ascending: false});
    }
    
    Logger.log('✓ Written to master sheet: ' + communityName);
    return true;
    
  } catch (error) {
    Logger.log('Error writing to master sheet: ' + error.toString());
    return false;
  }
}

/**
 * Get or create a community tab in the master sheet
 */
function getOrCreateMasterCommunityTab(masterSpreadsheet, communityName) {
  var tabName = communityName; // Use full community name
  var sheet = masterSpreadsheet.getSheetByName(tabName);
  
  if (!sheet) {
    Logger.log('Creating new tab in master sheet: ' + tabName);
    sheet = masterSpreadsheet.insertSheet(tabName);
    formatMasterCommunityTab(sheet, communityName);
  }
  
  return sheet;
}

/**
 * Format a community tab in the master sheet
 * UPDATED: Type column removed - 7 columns instead of 8
 */
function formatMasterCommunityTab(sheet, communityName) {
  // Headers - 7 columns (no Type)
  var headers = [
    COLUMNS.TIMESTAMP,
    COLUMNS.COMMUNITY,
    COLUMNS.FIRST_NAME,
    COLUMNS.LAST_NAME,
    COLUMNS.EMAIL,
    COLUMNS.SESSION,
    COLUMNS.SUNDAY_DATE
  ];
  
  sheet.appendRow(headers);
  
  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold')
            .setFontSize(12)
            .setBackground(SETTINGS.COLORS.PRIMARY)
            .setFontColor('#ffffff')
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');
  
  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1);
  
  // Set column widths
  sheet.setColumnWidth(1, 180);  // Timestamp
  sheet.setColumnWidth(2, 180);  // Community
  sheet.setColumnWidth(3, 120);  // First Name
  sheet.setColumnWidth(4, 120);  // Last Name
  sheet.setColumnWidth(5, 250);  // Email
  sheet.setColumnWidth(6, 180);  // Session
  sheet.setColumnWidth(7, 180);  // Sunday Date
  
  // Add title
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, 7).merge()
       .setValue(communityName.toUpperCase() + ' - REGISTRATIONS')
       .setFontWeight('bold')
       .setFontSize(14)
       .setBackground(SETTINGS.COLORS.SECONDARY)
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 50);
  sheet.setFrozenRows(2);
  
  Logger.log('✓ Master community tab formatted: ' + communityName);
}

/**
 * Setup all summary tabs in master sheet
 */
function setupMasterSummaryTabs() {
  try {
    Logger.log('Setting up master sheet summary tabs...');
    
    var masterSpreadsheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
    
    // Create Overall Summary tab
    createOverallSummaryTab(masterSpreadsheet);
    
    // Create Community Comparison tab
    createCommunityComparisonTab(masterSpreadsheet);
    
    // Create Weekly Trends tab
    createWeeklyTrendsTab(masterSpreadsheet);
    
    Logger.log('✓ All master summary tabs created');
    return true;
    
  } catch (error) {
    Logger.log('Error setting up master summary tabs: ' + error.toString());
    return false;
  }
}

/**
 * Create Overall Summary tab
 * UPDATED: Members/Guests columns removed
 */
function createOverallSummaryTab(masterSpreadsheet) {
  var tabName = MASTER_TAB_NAMES.OVERALL_SUMMARY;
  var sheet = masterSpreadsheet.getSheetByName(tabName);
  
  if (sheet) {
    masterSpreadsheet.deleteSheet(sheet);
  }
  
  sheet = masterSpreadsheet.insertSheet(tabName);
  
  // Title
  sheet.getRange(1, 1, 1, 4).merge()
       .setValue('📊 OVERALL REGISTRATION SUMMARY - ALL COMMUNITIES')
       .setFontWeight('bold')
       .setFontSize(14)
       .setBackground(SETTINGS.COLORS.PRIMARY)
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 50);
  
  // Last updated
  sheet.getRange(2, 1, 1, 4).merge()
       .setValue('Last Updated: ' + new Date().toLocaleString())
       .setFontSize(10)
       .setFontColor('#666666')
       .setHorizontalAlignment('center');
  
  // Headers - SIMPLIFIED (no Members/Guests)
  var headers = ['Community', 'Total Registrations', 'Unique Emails', 'Last Registration'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers])
       .setFontWeight('bold')
       .setBackground(SETTINGS.COLORS.SECONDARY)
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center');
  
  sheet.setRowHeight(4, 40);
  sheet.setFrozenRows(4);
  
  // Set column widths
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 180);
  
  Logger.log('✓ Overall Summary tab created');
}

/**
 * Create Community Comparison tab
 */
function createCommunityComparisonTab(masterSpreadsheet) {
  var tabName = MASTER_TAB_NAMES.COMMUNITY_COMPARISON;
  var sheet = masterSpreadsheet.getSheetByName(tabName);
  
  if (sheet) {
    masterSpreadsheet.deleteSheet(sheet);
  }
  
  sheet = masterSpreadsheet.insertSheet(tabName);
  
  // Title
  sheet.getRange(1, 1, 1, 5).merge()
       .setValue('📈 COMMUNITY COMPARISON - MONTHLY BREAKDOWN')
       .setFontWeight('bold')
       .setFontSize(14)
       .setBackground(SETTINGS.COLORS.PRIMARY)
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 50);
  
  Logger.log('✓ Community Comparison tab created');
}

/**
 * Create Weekly Trends tab
 */
function createWeeklyTrendsTab(masterSpreadsheet) {
  var tabName = MASTER_TAB_NAMES.WEEKLY_TRENDS;
  var sheet = masterSpreadsheet.getSheetByName(tabName);
  
  if (sheet) {
    masterSpreadsheet.deleteSheet(sheet);
  }
  
  sheet = masterSpreadsheet.insertSheet(tabName);
  
  // Title
  sheet.getRange(1, 1, 1, 5).merge()
       .setValue('📅 WEEKLY TRENDS - SUNDAY BY SUNDAY')
       .setFontWeight('bold')
       .setFontSize(14)
       .setBackground(SETTINGS.COLORS.PRIMARY)
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center');
  
  sheet.setRowHeight(1, 50);
  
  Logger.log('✓ Weekly Trends tab created');
}

/**
 * Update all master sheet summary tabs
 */
function updateMasterSummaryTabs() {
  try {
    Logger.log('Updating master sheet summary tabs...');
    
    var masterSpreadsheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
    
    updateOverallSummary(masterSpreadsheet);
    updateCommunityComparison(masterSpreadsheet);
    updateWeeklyTrends(masterSpreadsheet);
    
    Logger.log('✓ Master summary tabs updated');
    
  } catch (error) {
    Logger.log('Error updating master summary tabs: ' + error.toString());
  }
}

/**
 * Update Overall Summary tab with current data
 * UPDATED: Members/Guests statistics removed
 */
function updateOverallSummary(masterSpreadsheet) {
  var sheet = masterSpreadsheet.getSheetByName(MASTER_TAB_NAMES.OVERALL_SUMMARY);
  if (!sheet) return;
  
  // Clear existing data (keep headers)
  var lastRow = sheet.getLastRow();
  if (lastRow > 4) {
    sheet.getRange(5, 1, lastRow - 4, 4).clear();
  }
  
  var communities = getCommunityNames();
  var summaryData = [];
  
  for (var i = 0; i < communities.length; i++) {
    var communityName = communities[i];
    var communityTab = masterSpreadsheet.getSheetByName(communityName);
    
    if (communityTab) {
      var data = communityTab.getDataRange().getValues();
      
      if (data.length > 2) { // Has data beyond title and headers
        var registrations = data.slice(2); // Skip title and header rows
        var emails = [];
        var lastTimestamp = null;
        
        for (var j = 0; j < registrations.length; j++) {
          var email = registrations[j][4];
          if (email && emails.indexOf(email) === -1) {
            emails.push(email);
          }
          
          if (!lastTimestamp || registrations[j][0] > lastTimestamp) {
            lastTimestamp = registrations[j][0];
          }
        }
        
        summaryData.push([
          communityName,
          registrations.length,
          emails.length,
          lastTimestamp ? Utilities.formatDate(new Date(lastTimestamp), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : 'N/A'
        ]);
      } else {
        summaryData.push([
          communityName,
          0,
          0,
          'No data'
        ]);
      }
    }
  }
  
  if (summaryData.length > 0) {
    sheet.getRange(5, 1, summaryData.length, 4).setValues(summaryData);
    
    // Format rows
    for (var i = 0; i < summaryData.length; i++) {
      var bgColor = i % 2 === 0 ? SETTINGS.COLORS.LIGHT_BG : SETTINGS.COLORS.WHITE;
      sheet.getRange(5 + i, 1, 1, 4).setBackground(bgColor);
      sheet.getRange(5 + i, 2, 1, 3).setHorizontalAlignment('center');
    }
  }
  
  // Update timestamp
  sheet.getRange(2, 1, 1, 4).merge()
       .setValue('Last Updated: ' + new Date().toLocaleString());
}

/**
 * Update Community Comparison tab
 */
function updateCommunityComparison(masterSpreadsheet) {
  // Placeholder for future implementation
  Logger.log('Community Comparison update - placeholder');
}

/**
 * Update Weekly Trends tab
 */
function updateWeeklyTrends(masterSpreadsheet) {
  // Placeholder for future implementation
  Logger.log('Weekly Trends update - placeholder');
}
