/**
 * FORMATTING AND DESIGN FUNCTIONS - NO TYPE COLUMN
 * Blue/white theme formatting functions
 * UPDATED: Removed all Type column validation (7 columns instead of 8)
 */

/**
 * Format the Registrations tab for a community sheet
 * UPDATED: No Type column, no Type validation
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
  
  // Set column widths (7 columns)
  sheet.setColumnWidth(1, 180);  // Timestamp
  sheet.setColumnWidth(2, 180);  // Community
  sheet.setColumnWidth(3, 120);  // First Name
  sheet.setColumnWidth(4, 120);  // Last Name
  sheet.setColumnWidth(5, 250);  // Email
  sheet.setColumnWidth(6, 180);  // Session
  sheet.setColumnWidth(7, 180);  // Sunday Date
  
  // NO TYPE VALIDATION (Type column removed)
  
  // Apply row banding (7 columns)
  var dataRange = sheet.getRange(2, 1, 1000, 7);
  var banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  banding.setHeaderRowColor(SETTINGS.COLORS.PRIMARY)
         .setFirstRowColor(SETTINGS.COLORS.LIGHT_BG)
         .setSecondRowColor(SETTINGS.COLORS.WHITE);
  
  // Format timestamp column
  sheet.getRange(2, 1, 1000, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  
  Logger.log('✓ Registrations tab formatted (7 columns, no Type)');
}

/**
 * Format dashboard title row
 */
function formatDashboardTitle(sheet, title, rowHeight, numCols) {
  sheet.getRange(1, 1, 1, numCols).merge()
       .setValue(title)
       .setFontWeight('bold')
       .setFontSize(14)
       .setBackground(SETTINGS.COLORS.PRIMARY)
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
  
  sheet.setRowHeight(1, rowHeight);
}

/**
 * Format dashboard headers
 */
function formatDashboardHeaders(sheet, headers, rowNum) {
  sheet.getRange(rowNum, 1, 1, headers.length).setValues([headers])
       .setFontWeight('bold')
       .setFontSize(12)
       .setBackground(SETTINGS.COLORS.SECONDARY)
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
  
  sheet.setRowHeight(rowNum, 40);
}

/**
 * Apply blue/white alternating row colors
 */
function applyDashboardRowColors(sheet, startRow, numRows, numCols) {
  for (var i = 0; i < numRows; i++) {
    var bgColor = i % 2 === 0 ? SETTINGS.COLORS.LIGHT_BG : SETTINGS.COLORS.WHITE;
    sheet.getRange(startRow + i, 1, 1, numCols).setBackground(bgColor);
  }
}

/**
 * Add timestamp to dashboard
 */
function addDashboardTimestamp(sheet, rowNum, numCols) {
  sheet.getRange(rowNum, 1, 1, numCols).merge()
       .setValue('Last Updated: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'))
       .setFontSize(10)
       .setFontColor('#666666')
       .setHorizontalAlignment('center');
}
