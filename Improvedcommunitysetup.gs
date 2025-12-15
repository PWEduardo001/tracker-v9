/**
 * IMPROVED COMMUNITY SETUP - NO TYPE COLUMN
 * Enhanced setup process:
 * - Handles existing tabs gracefully
 * - No multiple alerts
 * - No duplicate tabs
 * - Applies blue/white theme consistently
 * UPDATED: Removed all Type column references (7 columns instead of 8)
 */

/**
 * Generic setup function for any community
 * IMPROVED: Better handling of existing tabs, single alert at end
 */
function setupCommunity(communityName) {
  try {
    logSeparator('SETTING UP: ' + communityName);
    
    // Validate community
    if (!isValidCommunity(communityName)) {
      throw new Error('Invalid community: ' + communityName);
    }
    
    var sheetId = getCommunitySheetId(communityName);
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    
    Logger.log('Connected to: ' + spreadsheet.getName());
    Logger.log('  URL: ' + spreadsheet.getUrl());
    Logger.log('');
    
    // STEP 1: Setup Registrations tab
    Logger.log('STEP 1: Setting up Registrations tab...');
    var regSheet = spreadsheet.getSheetByName(TAB_NAMES.REGISTRATIONS);
    
    if (regSheet) {
      var lastRow = regSheet.getLastRow();
      if (lastRow > 1) {
        // Has data - just reformat headers to ensure blue/white theme
        Logger.log('Registrations tab exists with data. Updating format...');
        updateRegistrationsTabFormat(regSheet);
      } else {
        // Empty - recreate with blue/white theme
        spreadsheet.deleteSheet(regSheet);
        Logger.log('Removed empty Registrations tab');
        regSheet = spreadsheet.insertSheet(TAB_NAMES.REGISTRATIONS);
        formatCommunityRegistrationsTab(regSheet);
        Logger.log('New Registrations tab created');
      }
    } else {
      regSheet = spreadsheet.insertSheet(TAB_NAMES.REGISTRATIONS);
      formatCommunityRegistrationsTab(regSheet);
      Logger.log('Registrations tab created');
    }
    
    Logger.log('');
    
    // STEP 2: Setup Current Month tab
    Logger.log('STEP 2: Setting up Current Month tab...');
    setupOrUpdateTab(spreadsheet, TAB_NAMES.MONTHLY, updateCommunityMonthly, sheetId);
    Logger.log('');
    
    // STEP 3: Setup Yearly Summary tab
    Logger.log('STEP 3: Setting up Yearly Summary tab...');
    setupOrUpdateTab(spreadsheet, TAB_NAMES.YEARLY, updateCommunityYearly, sheetId);
    Logger.log('');
    
    // STEP 4: Setup Sunday Breakdown tab
    Logger.log('STEP 4: Setting up Sunday Breakdown tab...');
    setupOrUpdateTab(spreadsheet, TAB_NAMES.SUNDAY, updateCommunitySunday, sheetId);
    Logger.log('');
    
    logSeparator('SETUP COMPLETE: ' + communityName);
    Logger.log('Sheet URL: ' + spreadsheet.getUrl());
    Logger.log('');
    Logger.log('All 4 tabs ready:');
    Logger.log('  1. ' + TAB_NAMES.REGISTRATIONS + ' (main data)');
    Logger.log('  2. ' + TAB_NAMES.MONTHLY + ' (auto-updating)');
    Logger.log('  3. ' + TAB_NAMES.YEARLY + ' (auto-updating)');
    Logger.log('  4. ' + TAB_NAMES.SUNDAY + ' (auto-updating)');
    logSeparator();
    
    return {
      success: true,
      message: 'Setup completed successfully for ' + communityName,
      url: spreadsheet.getUrl()
    };
    
  } catch (error) {
    logSeparator('SETUP FAILED: ' + communityName);
    Logger.log('Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    logSeparator();
    
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * Helper function to setup or update a tab
 * Avoids duplicate tabs and applies consistent formatting
 */
function setupOrUpdateTab(spreadsheet, tabName, updateFunction, sheetId) {
  var sheet = spreadsheet.getSheetByName(tabName);
  
  if (sheet) {
    // Tab exists - clear and update
    Logger.log(tabName + ' tab exists. Updating...');
    sheet.clear();
  } else {
    // Create new tab
    sheet = spreadsheet.insertSheet(tabName);
    Logger.log(tabName + ' tab created');
  }
  
  // Populate with data
  updateFunction(sheetId);
  Logger.log(tabName + ' tab populated');
}

/**
 * Update formatting of existing Registrations tab
 * Ensures blue/white theme is applied
 */
function updateRegistrationsTabFormat(sheet) {
  // Update header row to blue/white theme
  var lastCol = sheet.getLastColumn();
  if (lastCol > 0) {
    var headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setFontWeight('bold')
              .setFontSize(12)
              .setBackground(SETTINGS.COLORS.PRIMARY)      // Dark blue
              .setFontColor('#ffffff')                      // White text
              .setHorizontalAlignment('center')
              .setVerticalAlignment('middle');
    
    sheet.setRowHeight(1, 40);
    sheet.setFrozenRows(1);
  }
  
  Logger.log('✓ Registrations tab format updated to blue/white theme');
}

/**
 * Update Current Month tab with blue/white theme
 * UPDATED: 7 columns (no Type column)
 */
function updateCommunityMonthly(sheetId) {
  try {
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var mainSheet = spreadsheet.getSheetByName(TAB_NAMES.REGISTRATIONS);
    var monthlySheet = spreadsheet.getSheetByName(TAB_NAMES.MONTHLY);
    
    if (!monthlySheet) {
      monthlySheet = spreadsheet.insertSheet(TAB_NAMES.MONTHLY);
    } else {
      monthlySheet.clear();
    }
    
    var now = new Date();
    var currentMonth = now.getMonth();
    var currentYear = now.getFullYear();
    
    // Title - BLUE THEME (7 columns instead of 8)
    monthlySheet.getRange(1, 1, 1, 7).merge()
                .setValue('REGISTRATIONS FOR ' + getMonthName(currentMonth) + ' ' + currentYear)
                .setFontWeight('bold')
                .setFontSize(14)
                .setBackground(SETTINGS.COLORS.PRIMARY)
                .setFontColor('#ffffff')
                .setHorizontalAlignment('center');
    
    monthlySheet.setRowHeight(1, 50);
    
    // Headers - BLUE THEME (7 columns)
    var headers = Object.values(COLUMNS); // Now has 7 headers (no Type)
    monthlySheet.getRange(2, 1, 1, headers.length).setValues([headers])
                .setFontWeight('bold')
                .setBackground(SETTINGS.COLORS.SECONDARY)
                .setFontColor('#ffffff')
                .setHorizontalAlignment('center');
    
    monthlySheet.setRowHeight(2, 40);
    monthlySheet.setFrozenRows(2);
    
    // Get data for current month
    var mainData = mainSheet.getDataRange().getValues();
    var monthlyData = [];
    
    for (var i = 1; i < mainData.length; i++) {
      var timestamp = new Date(mainData[i][0]);
      if (timestamp.getMonth() === currentMonth && timestamp.getFullYear() === currentYear) {
        monthlyData.push(mainData[i]);
      }
    }
    
    // Add data if exists (7 columns)
    if (monthlyData.length > 0) {
      monthlySheet.getRange(3, 1, monthlyData.length, 7).setValues(monthlyData);
      
      // Apply blue/white alternating rows
      for (var i = 0; i < monthlyData.length; i++) {
        var bgColor = i % 2 === 0 ? SETTINGS.COLORS.LIGHT_BG : SETTINGS.COLORS.WHITE;
        monthlySheet.getRange(3 + i, 1, 1, 7).setBackground(bgColor);
      }
    }
    
    // Set column widths (7 columns)
    monthlySheet.setColumnWidth(1, 180);  // Timestamp
    monthlySheet.setColumnWidth(2, 180);  // Community
    monthlySheet.setColumnWidth(3, 120);  // First Name
    monthlySheet.setColumnWidth(4, 120);  // Last Name
    monthlySheet.setColumnWidth(5, 250);  // Email
    monthlySheet.setColumnWidth(6, 180);  // Session
    monthlySheet.setColumnWidth(7, 180);  // Sunday Date
    
    Logger.log('✓ Current Month tab updated');
    
  } catch (error) {
    Logger.log('Error updating monthly summary: ' + error.toString());
  }
}

/**
 * Update Yearly Summary tab with blue/white theme
 * UPDATED: Removed Members/Guests columns (3 columns instead of 5)
 */
function updateCommunityYearly(sheetId) {
  try {
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var mainSheet = spreadsheet.getSheetByName(TAB_NAMES.REGISTRATIONS);
    var yearlySheet = spreadsheet.getSheetByName(TAB_NAMES.YEARLY);
    
    if (!yearlySheet) {
      yearlySheet = spreadsheet.insertSheet(TAB_NAMES.YEARLY);
    } else {
      yearlySheet.clear();
    }
    
    var currentYear = new Date().getFullYear();
    
    // Title - BLUE THEME
    yearlySheet.getRange(1, 1, 1, 3).merge()
               .setValue('YEARLY REGISTRATION SUMMARY - ' + currentYear)
               .setFontWeight('bold')
               .setFontSize(14)
               .setBackground(SETTINGS.COLORS.PRIMARY)
               .setFontColor('#ffffff')
               .setHorizontalAlignment('center');
    
    yearlySheet.setRowHeight(1, 50);
    yearlySheet.setRowHeight(2, 15); // Blank row
    
    // Headers - BLUE THEME (3 columns: Month, Total, Unique Emails)
    var headers = ['Month', 'Total Registrations', 'Unique Emails'];
    yearlySheet.getRange(3, 1, 1, headers.length).setValues([headers])
               .setFontWeight('bold')
               .setBackground(SETTINGS.COLORS.SECONDARY)
               .setFontColor('#ffffff')
               .setHorizontalAlignment('center');
    
    yearlySheet.setRowHeight(3, 40);
    
    // Get all data
    var mainData = mainSheet.getDataRange().getValues();
    
    // Calculate monthly stats (no Member/Guest breakdown)
    var monthlyStats = [];
    for (var month = 0; month < 12; month++) {
      var monthData = [];
      
      for (var i = 1; i < mainData.length; i++) {
        var timestamp = new Date(mainData[i][0]);
        if (timestamp.getMonth() === month && timestamp.getFullYear() === currentYear) {
          monthData.push(mainData[i]);
        }
      }
      
      if (monthData.length > 0 || month <= new Date().getMonth()) {
        monthlyStats.push([
          getMonthName(month),
          monthData.length,
          getUniqueEmails(monthData).length
        ]);
      }
    }
    
    // Add monthly data
    if (monthlyStats.length > 0) {
      yearlySheet.getRange(4, 1, monthlyStats.length, 3).setValues(monthlyStats);
      
      // Apply blue/white alternating rows
      for (var i = 0; i < monthlyStats.length; i++) {
        var bgColor = i % 2 === 0 ? SETTINGS.COLORS.LIGHT_BG : SETTINGS.COLORS.WHITE;
        yearlySheet.getRange(4 + i, 1, 1, 3).setBackground(bgColor);
      }
      
      // Bold month names
      yearlySheet.getRange(4, 1, monthlyStats.length, 1).setFontWeight('bold');
      
      // Center numeric columns
      yearlySheet.getRange(4, 2, monthlyStats.length, 2).setHorizontalAlignment('center');
    }
    
    // Set column widths
    yearlySheet.setColumnWidth(1, 120);  // Month
    yearlySheet.setColumnWidth(2, 150);  // Total Registrations
    yearlySheet.setColumnWidth(3, 150);  // Unique Emails
    
    Logger.log('✓ Yearly Summary tab updated');
    
  } catch (error) {
    Logger.log('Error updating yearly summary: ' + error.toString());
  }
}

/**
 * Update Sunday Breakdown tab with blue/white theme
 * UPDATED: Removed Members/Guests columns (4 columns instead of 6)
 */
function updateCommunitySunday(sheetId) {
  try {
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var mainSheet = spreadsheet.getSheetByName(TAB_NAMES.REGISTRATIONS);
    var sundaySheet = spreadsheet.getSheetByName(TAB_NAMES.SUNDAY);
    
    if (!sundaySheet) {
      sundaySheet = spreadsheet.insertSheet(TAB_NAMES.SUNDAY);
    } else {
      sundaySheet.clear();
    }
    
    // Title - BLUE THEME
    sundaySheet.getRange(1, 1, 1, 4).merge()
               .setValue('📊 SUNDAY SERVICE REGISTRATION BREAKDOWN')
               .setFontWeight('bold')
               .setFontSize(14)
               .setBackground(SETTINGS.COLORS.PRIMARY)
               .setFontColor('#ffffff')
               .setHorizontalAlignment('center');
    
    sundaySheet.setRowHeight(1, 50);
    
    // Last updated timestamp
    sundaySheet.getRange(2, 1, 1, 4).merge()
               .setValue('Last Updated: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'))
               .setFontSize(10)
               .setFontColor('#666666')
               .setHorizontalAlignment('center');
    
    sundaySheet.setRowHeight(3, 15); // Blank row
    
    // Headers - BLUE THEME (4 columns: Date, Total, Unique Emails, Details)
    var headers = ['Sunday Date', 'Total Registrations', 'Unique Emails', 'Details'];
    sundaySheet.getRange(4, 1, 1, headers.length).setValues([headers])
               .setFontWeight('bold')
               .setBackground(SETTINGS.COLORS.SECONDARY)
               .setFontColor('#ffffff')
               .setHorizontalAlignment('center');
    
    sundaySheet.setRowHeight(4, 40);
    sundaySheet.setFrozenRows(4);
    
    // Get all data
    var mainData = mainSheet.getDataRange().getValues();
    
    // Group by Sunday date
    var sundayGroups = {};
    
    for (var i = 1; i < mainData.length; i++) {
      var sundayDate = mainData[i][6]; // Sunday Date column (was column 7, now column 6)
      
      if (sundayDate) {
        if (!sundayGroups[sundayDate]) {
          sundayGroups[sundayDate] = [];
        }
        sundayGroups[sundayDate].push(mainData[i]);
      }
    }
    
    // Convert to sorted array
    var sundayDates = Object.keys(sundayGroups).sort().reverse(); // Most recent first
    
    var sundayData = [];
    for (var i = 0; i < sundayDates.length; i++) {
      var date = sundayDates[i];
      var dateData = sundayGroups[date];
      
      sundayData.push([
        date,
        dateData.length,
        getUniqueEmails(dateData).length,
        'View Details →'
      ]);
    }
    
    // Add data
    if (sundayData.length > 0) {
      sundaySheet.getRange(5, 1, sundayData.length, 4).setValues(sundayData);
      
      // Apply blue/white alternating rows
      for (var i = 0; i < sundayData.length; i++) {
        var rowNum = 5 + i;
        var bgColor = i % 2 === 0 ? SETTINGS.COLORS.LIGHT_BG : SETTINGS.COLORS.WHITE;
        
        sundaySheet.getRange(rowNum, 1, 1, 4).setBackground(bgColor);
        sundaySheet.getRange(rowNum, 1).setFontWeight('bold');
        sundaySheet.getRange(rowNum, 2, 1, 2).setHorizontalAlignment('center');
      }
    } else {
      sundaySheet.getRange(5, 1, 1, 4).merge()
                 .setValue('No registrations yet')
                 .setHorizontalAlignment('center')
                 .setFontStyle('italic')
                 .setFontColor('#999999');
    }
    
    // Set column widths
    sundaySheet.setColumnWidth(1, 200);  // Sunday Date
    sundaySheet.setColumnWidth(2, 150);  // Total Registrations
    sundaySheet.setColumnWidth(3, 150);  // Unique Emails
    sundaySheet.setColumnWidth(4, 150);  // Details
    
    Logger.log('✓ Sunday Breakdown tab updated');
    
  } catch (error) {
    Logger.log('Error updating Sunday breakdown: ' + error.toString());
  }
}

/**
 * Setup ALL communities at once
 * WARNING: This will take several minutes
 */
function setupAllCommunities() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Setup All Communities',
    'This will set up all 8 community sheets. This may take several minutes. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    logSeparator('SETTING UP ALL COMMUNITIES');
    
    var communities = getCommunityNames();
    var successCount = 0;
    var failCount = 0;
    
    for (var i = 0; i < communities.length; i++) {
      var communityName = communities[i];
      Logger.log('');
      Logger.log('Setting up ' + (i + 1) + '/' + communities.length + ': ' + communityName);
      
      try {
        setupCommunity(communityName);
        successCount++;
        Logger.log(communityName + ' completed');
        
        // Brief pause between setups to avoid rate limiting
        if (i < communities.length - 1) {
          Utilities.sleep(2000);
        }
      } catch (error) {
        failCount++;
        Logger.log(communityName + ' failed: ' + error.toString());
      }
    }
    
    logSeparator('ALL COMMUNITIES SETUP COMPLETE');
    Logger.log('Success: ' + successCount);
    Logger.log('Failed: ' + failCount);
    logSeparator();
    
    ui.alert(
      'Setup Complete',
      'Setup completed for all communities:\n\n' +
      'Success: ' + successCount + '\n' +
      'Failed: ' + failCount + '\n\n' +
      'Check the Execution log for details.',
      ui.ButtonSet.OK
    );
  }
}
