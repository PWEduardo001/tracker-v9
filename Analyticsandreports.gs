/**
 * ANALYTICS AND REPORTS
 * Functional implementations of Community Comparison and Weekly Trends
 * Replaces placeholder functions with real data aggregation
 */

// ==================== OVERALL SUMMARY ====================

/**
 * Update Overall Summary tab with current data
 * UPDATED: Blue/white theme
 */
function updateOverallSummary(masterSpreadsheet) {
  var sheet = masterSpreadsheet.getSheetByName(MASTER_TAB_NAMES.OVERALL_SUMMARY);
  if (!sheet) {
    createOverallSummaryTab(masterSpreadsheet);
    sheet = masterSpreadsheet.getSheetByName(MASTER_TAB_NAMES.OVERALL_SUMMARY);
  }
  
  // Clear existing data (keep headers)
  var lastRow = sheet.getLastRow();
  if (lastRow > 4) {
    sheet.getRange(5, 1, lastRow - 4, 6).clear();
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
        var members = 0;
        var guests = 0;
        var emails = [];
        var lastTimestamp = null;
        
        for (var j = 0; j < registrations.length; j++) {
          if (registrations[j][5] === SETTINGS.VALIDATION.MEMBER) members++;
          if (registrations[j][5] === SETTINGS.VALIDATION.GUEST) guests++;
          
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
          members,
          guests,
          emails.length,
          lastTimestamp ? Utilities.formatDate(new Date(lastTimestamp), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : 'N/A'
        ]);
      } else {
        summaryData.push([
          communityName,
          0,
          0,
          0,
          0,
          'No data'
        ]);
      }
    }
  }
  
  if (summaryData.length > 0) {
    sheet.getRange(5, 1, summaryData.length, 6).setValues(summaryData);
    
    // Format rows - BLUE/WHITE THEME
    applyDashboardRowColors(sheet, 5, summaryData.length, 6);
    
    // Center numeric columns
    for (var i = 0; i < summaryData.length; i++) {
      sheet.getRange(5 + i, 2, 1, 5).setHorizontalAlignment('center');
    }
  }
  
  // Update timestamp
  addDashboardTimestamp(sheet, 2, 6);
  
  Logger.log('✓ Overall Summary updated');
}

// ==================== COMMUNITY COMPARISON ====================

/**
 * Update Community Comparison tab with monthly breakdown
 * FULLY FUNCTIONAL: Shows each community's registrations by month
 */
function updateCommunityComparison(masterSpreadsheet) {
  var sheet = masterSpreadsheet.getSheetByName(MASTER_TAB_NAMES.COMMUNITY_COMPARISON);
  if (!sheet) {
    createCommunityComparisonTab(masterSpreadsheet);
    sheet = masterSpreadsheet.getSheetByName(MASTER_TAB_NAMES.COMMUNITY_COMPARISON);
  }
  
  // Clear existing content
  sheet.clear();
  
  // Title - BLUE THEME
  formatDashboardTitle(sheet, '📈 COMMUNITY COMPARISON - MONTHLY BREAKDOWN', 50, 10);
  
  // Last updated
  addDashboardTimestamp(sheet, 2, 10);
  
  sheet.setRowHeight(3, 15); // Blank row
  
  // Get current year
  var currentYear = new Date().getFullYear();
  var currentMonth = new Date().getMonth();
  
  // Prepare headers: Month + each community
  var communities = getCommunityNames();
  var headers = ['Month'];
  for (var i = 0; i < communities.length; i++) {
    headers.push(communities[i].replace(' Family Church', '')); // Shortened names
  }
  headers.push('Total');
  
  formatDashboardHeaders(sheet, headers, 4);
  sheet.setFrozenRows(4);
  
  // Collect data for each month
  var monthlyData = [];
  
  for (var month = 0; month <= currentMonth; month++) {
    var row = [getMonthName(month)];
    var monthTotal = 0;
    
    for (var i = 0; i < communities.length; i++) {
      var communityName = communities[i];
      var communityTab = masterSpreadsheet.getSheetByName(communityName);
      
      var count = 0;
      if (communityTab) {
        var data = communityTab.getDataRange().getValues();
        
        for (var j = 2; j < data.length; j++) { // Skip title and header
          var timestamp = new Date(data[j][0]);
          if (timestamp.getFullYear() === currentYear && timestamp.getMonth() === month) {
            count++;
          }
        }
      }
      
      row.push(count);
      monthTotal += count;
    }
    
    row.push(monthTotal);
    monthlyData.push(row);
  }
  
  // Write data
  if (monthlyData.length > 0) {
    sheet.getRange(5, 1, monthlyData.length, headers.length).setValues(monthlyData);
    
    // Apply formatting - BLUE/WHITE THEME
    applyDashboardRowColors(sheet, 5, monthlyData.length, headers.length);
    
    // Center all numeric columns
    sheet.getRange(5, 2, monthlyData.length, headers.length - 1).setHorizontalAlignment('center');
    
    // Bold the Total column
    sheet.getRange(5, headers.length, monthlyData.length, 1).setFontWeight('bold');
    
    // Bold month names
    sheet.getRange(5, 1, monthlyData.length, 1).setFontWeight('bold');
  }
  
  // Set column widths
  sheet.setColumnWidth(1, 120); // Month
  for (var i = 2; i <= communities.length + 1; i++) {
    sheet.setColumnWidth(i, 100); // Community columns
  }
  sheet.setColumnWidth(headers.length, 100); // Total column
  
  // Add totals row at bottom
  var totalsRow = sheet.getLastRow() + 2;
  var totalsData = ['TOTAL'];
  var grandTotal = 0;
  
  for (var i = 0; i < communities.length; i++) {
    var communityTotal = 0;
    for (var j = 0; j < monthlyData.length; j++) {
      communityTotal += monthlyData[j][i + 1];
    }
    totalsData.push(communityTotal);
    grandTotal += communityTotal;
  }
  totalsData.push(grandTotal);
  
  sheet.getRange(totalsRow, 1, 1, headers.length).setValues([totalsData])
       .setFontWeight('bold')
       .setBackground(SETTINGS.COLORS.SECONDARY)
       .setFontColor(SETTINGS.COLORS.TEXT_LIGHT)
       .setHorizontalAlignment('center');
  
  sheet.getRange(totalsRow, 1).setHorizontalAlignment('left');
  
  Logger.log('✓ Community Comparison updated with data for ' + monthlyData.length + ' months');
}

// ==================== WEEKLY TRENDS ====================

/**
 * Update Weekly Trends tab with Sunday-by-Sunday breakdown
 * FULLY FUNCTIONAL: Shows registrations for each Sunday across all communities
 */
function updateWeeklyTrends(masterSpreadsheet) {
  var sheet = masterSpreadsheet.getSheetByName(MASTER_TAB_NAMES.WEEKLY_TRENDS);
  if (!sheet) {
    createWeeklyTrendsTab(masterSpreadsheet);
    sheet = masterSpreadsheet.getSheetByName(MASTER_TAB_NAMES.WEEKLY_TRENDS);
  }
  
  // Clear existing content
  sheet.clear();
  
  // Title - BLUE THEME
  formatDashboardTitle(sheet, '📅 WEEKLY TRENDS - SUNDAY BY SUNDAY', 50, 6);
  
  // Last updated
  addDashboardTimestamp(sheet, 2, 6);
  
  sheet.setRowHeight(3, 15); // Blank row
  
  // Headers
  var headers = ['Sunday Date', 'Total Registrations', 'Members', 'Guests', 'Unique Emails', 'Communities Active'];
  formatDashboardHeaders(sheet, headers, 4);
  sheet.setFrozenRows(4);
  
  // Collect all Sunday dates across all communities
  var sundayData = {};
  var communities = getCommunityNames();
  
  for (var i = 0; i < communities.length; i++) {
    var communityName = communities[i];
    var communityTab = masterSpreadsheet.getSheetByName(communityName);
    
    if (communityTab) {
      var data = communityTab.getDataRange().getValues();
      
      for (var j = 2; j < data.length; j++) { // Skip title and header
        var sundayDate = data[j][7]; // Sunday Date column
        
        if (sundayDate) {
          if (!sundayData[sundayDate]) {
            sundayData[sundayDate] = {
              total: 0,
              members: 0,
              guests: 0,
              emails: [],
              communities: []
            };
          }
          
          sundayData[sundayDate].total++;
          
          if (data[j][5] === SETTINGS.VALIDATION.MEMBER) {
            sundayData[sundayDate].members++;
          }
          if (data[j][5] === SETTINGS.VALIDATION.GUEST) {
            sundayData[sundayDate].guests++;
          }
          
          var email = data[j][4];
          if (email && sundayData[sundayDate].emails.indexOf(email) === -1) {
            sundayData[sundayDate].emails.push(email);
          }
          
          if (sundayData[sundayDate].communities.indexOf(communityName) === -1) {
            sundayData[sundayDate].communities.push(communityName);
          }
        }
      }
    }
  }
  
  // Convert to sorted array (most recent first)
  var sundayDates = Object.keys(sundayData).sort().reverse();
  var trendsData = [];
  
  for (var i = 0; i < sundayDates.length; i++) {
    var date = sundayDates[i];
    var stats = sundayData[date];
    
    trendsData.push([
      date,
      stats.total,
      stats.members,
      stats.guests,
      stats.emails.length,
      stats.communities.length
    ]);
  }
  
  // Write data
  if (trendsData.length > 0) {
    sheet.getRange(5, 1, trendsData.length, 6).setValues(trendsData);
    
    // Apply formatting - BLUE/WHITE THEME
    applyDashboardRowColors(sheet, 5, trendsData.length, 6);
    
    // Bold Sunday dates
    sheet.getRange(5, 1, trendsData.length, 1).setFontWeight('bold');
    
    // Center numeric columns
    sheet.getRange(5, 2, trendsData.length, 5).setHorizontalAlignment('center');
  } else {
    sheet.getRange(5, 1, 1, 6).merge()
         .setValue('No registration data available yet')
         .setHorizontalAlignment('center')
         .setFontStyle('italic')
         .setFontColor('#999999');
  }
  
  // Set column widths
  sheet.setColumnWidth(1, 200); // Sunday Date
  sheet.setColumnWidth(2, 150); // Total
  sheet.setColumnWidth(3, 120); // Members
  sheet.setColumnWidth(4, 120); // Guests
  sheet.setColumnWidth(5, 150); // Unique Emails
  sheet.setColumnWidth(6, 150); // Communities Active
  
  Logger.log('✓ Weekly Trends updated with ' + trendsData.length + ' Sundays');
}

// ==================== REFRESH ALL MASTER SUMMARY TABS ====================

/**
 * Update all master sheet summary tabs at once
 */
function updateMasterSummaryTabs() {
  try {
    Logger.log('Updating master sheet summary tabs...');
    
    var masterSpreadsheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
    
    updateOverallSummary(masterSpreadsheet);
    updateCommunityComparison(masterSpreadsheet);
    updateWeeklyTrends(masterSpreadsheet);
    
    Logger.log('✓ All master summary tabs updated');
    
  } catch (error) {
    Logger.log('Error updating master summary tabs: ' + error.toString());
  }
}
