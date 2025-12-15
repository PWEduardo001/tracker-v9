/**
 * COMMUNITY SHEET OPERATIONS
 * Handles operations for individual community sheets
 * UPDATED: Type column completely removed - no more Member/Guest distinction
 */

/**
 * Write registration to community sheet
 */
function writeToCommunitySheet(communityName, registrationData) {
  try {
    Logger.log('Writing to community sheet: ' + communityName);
    
    var sheetId = getCommunitySheetId(communityName);
    if (!sheetId) {
      throw new Error('Invalid community: ' + communityName);
    }
    
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var sheet = getOrCreateCommunityRegistrationsTab(spreadsheet);
    
    var timestamp = new Date();
    
    // Add each registrant - 7 columns (no Type)
    for (var i = 0; i < registrationData.registrants.length; i++) {
      var person = registrationData.registrants[i];
      
      sheet.appendRow([
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
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var range = sheet.getRange(2, 1, lastRow - 1, 7);
      range.sort({column: 1, ascending: false});
    }
    
    Logger.log('✓ Written to community sheet: ' + communityName);
    
    // Update dashboard tabs
    updateCommunityDashboards(sheetId);
    
    return true;
    
  } catch (error) {
    Logger.log('Error writing to community sheet: ' + error.toString());
    throw error;
  }
}

/**
 * Get or create the Registrations tab in a community sheet
 */
function getOrCreateCommunityRegistrationsTab(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(TAB_NAMES.REGISTRATIONS);
  
  if (!sheet) {
    Logger.log('Creating Registrations tab');
    sheet = spreadsheet.insertSheet(TAB_NAMES.REGISTRATIONS);
    formatCommunityRegistrationsTab(sheet);
  }
  
  return sheet;
}

/**
 * Format the Registrations tab for a community sheet
 * UPDATED: Type column removed - 7 columns instead of 8
 */
function formatCommunityRegistrationsTab(sheet) {
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
  
  // Apply row banding
  var dataRange = sheet.getRange(2, 1, 1000, 7);
  var banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  banding.setHeaderRowColor(SETTINGS.COLORS.PRIMARY)
         .setFirstRowColor(SETTINGS.COLORS.LIGHT_BG)
         .setSecondRowColor(SETTINGS.COLORS.WHITE);
  
  // Format timestamp column
  sheet.getRange(2, 1, 1000, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  
  Logger.log('✓ Registrations tab formatted');
}

/**
 * Update all dashboard tabs for a community
 */
function updateCommunityDashboards(sheetId) {
  try {
    Logger.log('Updating community dashboards...');
    
    updateCommunityMonthly(sheetId);
    updateCommunityYearly(sheetId);
    updateCommunitySunday(sheetId);
    
    Logger.log('✓ Community dashboards updated');
    
  } catch (error) {
    Logger.log('Error updating community dashboards: ' + error.toString());
  }
}

/**
 * Update Current Month tab
 * UPDATED: Type column removed
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
    
    // Title
    monthlySheet.getRange(1, 1, 1, 7).merge()
                .setValue('REGISTRATIONS FOR ' + getMonthName(currentMonth) + ' ' + currentYear)
                .setFontWeight('bold')
                .setFontSize(14)
                .setBackground(SETTINGS.COLORS.PRIMARY)
                .setFontColor('#ffffff')
                .setHorizontalAlignment('center');
    
    monthlySheet.setRowHeight(1, 50);
    
    // Headers
    var headers = Object.values(COLUMNS);
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
    
    // Add data if exists
    if (monthlyData.length > 0) {
      monthlySheet.getRange(3, 1, monthlyData.length, 7).setValues(monthlyData);
      
      var dataRange = monthlySheet.getRange(3, 1, monthlyData.length, 7);
      var banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
      banding.setFirstRowColor(SETTINGS.COLORS.LIGHT_BG).setSecondRowColor(SETTINGS.COLORS.WHITE);
    }
    
    // Set column widths
    monthlySheet.setColumnWidth(1, 180);
    monthlySheet.setColumnWidth(2, 180);
    monthlySheet.setColumnWidth(3, 120);
    monthlySheet.setColumnWidth(4, 120);
    monthlySheet.setColumnWidth(5, 250);
    monthlySheet.setColumnWidth(6, 180);
    monthlySheet.setColumnWidth(7, 180);
    
    Logger.log('✓ Current Month tab updated');
    
  } catch (error) {
    Logger.log('Error updating monthly summary: ' + error.toString());
  }
}

/**
 * Update Yearly Summary tab
 * UPDATED: Members/Guests columns removed
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
    
    // Title
    yearlySheet.getRange(1, 1, 1, 3).merge()
               .setValue('YEARLY REGISTRATION SUMMARY - ' + currentYear)
               .setFontWeight('bold')
               .setFontSize(14)
               .setBackground(SETTINGS.COLORS.PRIMARY)
               .setFontColor('#ffffff')
               .setHorizontalAlignment('center');
    
    yearlySheet.setRowHeight(1, 50);
    
    // Headers - SIMPLIFIED (no Members/Guests)
    var headers = ['Month', 'Total Registrations', 'Unique Emails'];
    yearlySheet.getRange(3, 1, 1, headers.length).setValues([headers])
               .setFontWeight('bold')
               .setBackground(SETTINGS.COLORS.SECONDARY)
               .setFontColor('#ffffff')
               .setHorizontalAlignment('center');
    
    yearlySheet.setRowHeight(3, 40);
    
    // Get all data
    var mainData = mainSheet.getDataRange().getValues();
    
    // Calculate monthly stats
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
      
      for (var i = 0; i < monthlyStats.length; i++) {
        var bgColor = i % 2 === 0 ? SETTINGS.COLORS.LIGHT_BG : SETTINGS.COLORS.WHITE;
        yearlySheet.getRange(4 + i, 1, 1, 3).setBackground(bgColor);
      }
    }
    
    // Set column widths
    yearlySheet.setColumnWidth(1, 120);
    yearlySheet.setColumnWidth(2, 150);
    yearlySheet.setColumnWidth(3, 150);
    
    Logger.log('✓ Yearly Summary tab updated');
    
  } catch (error) {
    Logger.log('Error updating yearly summary: ' + error.toString());
  }
}

/**
 * Update Sunday Breakdown tab
 * UPDATED: Members/Guests columns removed
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
    
    // Title
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
    
    // Headers - SIMPLIFIED (no Members/Guests)
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
      var sundayDate = mainData[i][6]; // Sunday Date column (index 6)
      
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
      
      // Format data rows
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
    sundaySheet.setColumnWidth(1, 200);
    sundaySheet.setColumnWidth(2, 150);
    sundaySheet.setColumnWidth(3, 150);
    sundaySheet.setColumnWidth(4, 150);
    
    Logger.log('✓ Sunday Breakdown tab updated');
    
  } catch (error) {
    Logger.log('Error updating Sunday breakdown: ' + error.toString());
  }
}
