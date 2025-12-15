/**
 * UTILITY FUNCTIONS - UPDATED FOR YEAR-LONG SERVICE SETTINGS
 * Shared helper functions used across the project
 * UPDATED: Type column completely removed - countByType() function deleted
 */

/**
 * Get month name from index
 */
function getMonthName(monthIndex) {
  var months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  return months[monthIndex];
}

/**
 * Get unique emails from registration data
 * UPDATED: Uses correct column index (4) for email
 */
function getUniqueEmails(data) {
  var emails = [];
  for (var i = 0; i < data.length; i++) {
    var email = data[i][COL_INDEX.EMAIL];
    if (email && emails.indexOf(email) === -1) {
      emails.push(email);
    }
  }
  return emails;
}

/**
 * Calculate the appropriate Sunday date for registration
 * If it's Sunday before SERVICE_END_HOUR, show today
 * Otherwise, show next Sunday
 * 
 * Returns date in format: "November 30, 2025" (without day of week)
 */
function getNextSunday() {
  var now = new Date();
  var dayOfWeek = now.getDay();
  var currentHour = now.getHours();
  
  var targetDate = new Date(now);
  
  if (dayOfWeek === 0) {
    if (currentHour < SETTINGS.SERVICE_END_HOUR) {
      Logger.log('It\'s Sunday before ' + SETTINGS.SERVICE_END_HOUR + ':00, showing today');
    } else {
      Logger.log('It\'s Sunday after ' + SETTINGS.SERVICE_END_HOUR + ':00, showing next Sunday');
      targetDate.setDate(now.getDate() + 7);
    }
  } else {
    var daysUntilSunday = (7 - dayOfWeek) % 7;
    if (daysUntilSunday === 0) daysUntilSunday = 7;
    targetDate.setDate(now.getDate() + daysUntilSunday);
    Logger.log('Today is not Sunday, showing upcoming Sunday');
  }
  
  var options = { year: 'numeric', month: 'long', day: 'numeric' };
  
  return {
    formatted: targetDate.toLocaleDateString('en-US', options),
    date: targetDate.toISOString(),
    dateOnly: Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    isToday: dayOfWeek === 0 && currentHour < SETTINGS.SERVICE_END_HOUR
  };
}

/**
 * Get the Nth upcoming Sunday from a given date
 */
function getNextSundayFromDate(fromDate, weeksAhead) {
  if (!weeksAhead) weeksAhead = 0;
  
  var targetDate = new Date(fromDate);
  var dayOfWeek = targetDate.getDay();
  
  var daysUntilSunday = dayOfWeek === 0 ? 7 : (7 - dayOfWeek);
  
  targetDate.setDate(targetDate.getDate() + daysUntilSunday + (weeksAhead * 7));
  
  return targetDate;
}

/**
 * Validate email format
 */
function isValidEmail(email) {
  if (!email) return false;
  var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email);
}

/**
 * Get timestamp in readable format
 */
function getFormattedTimestamp(date) {
  if (!date) date = new Date();
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

/**
 * Sleep function for rate limiting
 */
function sleep(milliseconds) {
  Utilities.sleep(milliseconds);
}

/**
 * Log with timestamp
 */
function logWithTimestamp(message) {
  Logger.log('[' + getFormattedTimestamp() + '] ' + message);
}

/**
 * Create a section separator for logs
 */
function logSeparator(title) {
  Logger.log('========================================');
  if (title) {
    Logger.log(title.toUpperCase());
    Logger.log('========================================');
  }
}

/**
 * Delete a sheet if it exists
 */
function deleteSheetIfExists(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    spreadsheet.deleteSheet(sheet);
    Logger.log('Deleted sheet: ' + sheetName);
    return true;
  }
  return false;
}

/**
 * Check if a sheet exists
 */
function sheetExists(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  return sheet !== null;
}

/**
 * Create a bordered section
 */
function addBorder(sheet, row, col, numRows, numCols, color) {
  if (!color) color = SETTINGS.COLORS.GRAY;
  
  var range = sheet.getRange(row, col, numRows, numCols);
  range.setBorder(
    true, true, true, true, true, true,
    color,
    SpreadsheetApp.BorderStyle.SOLID
  );
}

/**
 * Convert date to YYYY-MM-DD format
 */
function formatDateYMD(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// ==================== YEAR-LONG SERVICE SETTINGS FUNCTIONS ====================

/**
 * Setup Service Settings tab for a community
 * NEW APPROACH: Year-long table with all Sunday services
 */
function setupServiceSettingsTab(spreadsheet, communityName) {
  var tabName = 'Service Settings';
  
  var existingSheet = spreadsheet.getSheetByName(tabName);
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
    Logger.log('Deleted old Service Settings tab');
  }
  
  var sheet = spreadsheet.insertSheet(tabName);
  
  sheet.getRange(1, 1, 1, 4).merge()
       .setValue('📅 YEARLY SERVICE SCHEDULE - ' + communityName.toUpperCase())
       .setFontWeight('bold')
       .setFontSize(16)
       .setBackground(SETTINGS.COLORS.PRIMARY)
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 60);
  
  sheet.getRange(2, 1, 1, 4).merge()
       .setValue('📝 Fill in service details for each Sunday. The system will automatically send registrants the correct link for the next upcoming Sunday.')
       .setFontSize(11)
       .setFontColor('#666666')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle')
       .setWrap(true);
  sheet.setRowHeight(2, 50);
  
  sheet.setRowHeight(3, 15);
  
  var headers = ['Sunday Date', 'Service Time', 'Speaker Name', 'YouTube/Zoom Link'];
  sheet.getRange(4, 1, 1, 4).setValues([headers])
       .setFontWeight('bold')
       .setFontSize(12)
       .setBackground(SETTINGS.COLORS.SECONDARY)
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
  sheet.setRowHeight(4, 40);
  sheet.setFrozenRows(4);
  
  var startRow = 5;
  var numRows = 60;
  
  for (var i = 0; i < numRows; i++) {
    sheet.setRowHeight(startRow + i, 35);
  }
  
  for (var i = 0; i < numRows; i++) {
    var bgColor = i % 2 === 0 ? SETTINGS.COLORS.LIGHT_BG : SETTINGS.COLORS.WHITE;
    sheet.getRange(startRow + i, 1, 1, 4).setBackground(bgColor);
  }
  
  addBorder(sheet, 4, 1, numRows + 1, 4, SETTINGS.COLORS.SECONDARY);
  
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 400);
  
  var today = new Date();
  var exampleData = [];
  
  for (var i = 0; i < 3; i++) {
    var exampleDate = getNextSundayFromDate(today, i);
    var formattedDate = Utilities.formatDate(exampleDate, Session.getScriptTimeZone(), 'MMMM d, yyyy');
    
    exampleData.push([
      formattedDate,
      '10:30 AM - 12:00 PM',
      i === 0 ? 'Rev. Dr. Example Speaker' : '',
      i === 0 ? 'https://www.youtube.com/watch?v=example' : ''
    ]);
  }
  
  sheet.getRange(startRow, 1, 3, 4).setValues(exampleData);
  
  var noteRow = startRow + numRows + 2;
  sheet.getRange(noteRow, 1, 1, 4).merge()
       .setValue('💡 TIP: Fill in dates and links as far in advance as possible. If a link is not available when someone registers, they will be notified to try again later.')
       .setFontSize(10)
       .setFontStyle('italic')
       .setFontColor('#666666')
       .setWrap(true)
       .setBackground('#fffbf0');
  sheet.setRowHeight(noteRow, 50);
  
  Logger.log('Service Settings tab created with year-long table format');
}

/**
 * Get service settings for next upcoming Sunday from the year-long table
 */
function getNextSundayServiceSettings(communityName) {
  try {
    var sheetId = getCommunitySheetId(communityName);
    if (!sheetId) {
      Logger.log('Invalid community: ' + communityName);
      return null;
    }
    
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var settingsSheet = spreadsheet.getSheetByName('Service Settings');
    
    if (!settingsSheet) {
      Logger.log('Service Settings tab not found for: ' + communityName);
      return null;
    }
    
    var nextSunday = getNextSunday();
    var targetDate = nextSunday.formatted;
    
    Logger.log('Looking for service info for: ' + targetDate);
    
    var dataRange = settingsSheet.getRange(5, 1, 60, 4);
    var data = dataRange.getValues();
    
    for (var i = 0; i < data.length; i++) {
      var rowDate = data[i][0];
      
      if (!rowDate) continue;
      
      var rowDateStr = '';
      if (rowDate instanceof Date) {
        rowDateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'MMMM d, yyyy');
      } else {
        rowDateStr = rowDate.toString().trim();
      }
      
      if (rowDateStr === targetDate) {
        var serviceTime = data[i][1];
        var speakerName = data[i][2];
        var liveStreamLink = data[i][3];
        
        if (!liveStreamLink || liveStreamLink.toString().trim() === '') {
          Logger.log('No livestream link available for ' + targetDate);
          return {
            sundayDate: targetDate,
            speakerName: speakerName || 'TBA',
            serviceTime: serviceTime || '10:30 AM - 12:00 PM',
            liveStreamLink: null,
            linkAvailable: false
          };
        }
        
        Logger.log('Found service settings for ' + targetDate);
        Logger.log('  Speaker: ' + speakerName);
        Logger.log('  Time: ' + serviceTime);
        Logger.log('  Link: ' + liveStreamLink);
        
        return {
          sundayDate: targetDate,
          speakerName: speakerName || 'TBA',
          serviceTime: serviceTime || '10:30 AM - 12:00 PM',
          liveStreamLink: liveStreamLink,
          linkAvailable: true
        };
      }
    }
    
    Logger.log('Sunday date ' + targetDate + ' not found in Service Settings table');
    return {
      sundayDate: targetDate,
      speakerName: 'TBA',
      serviceTime: '10:30 AM - 12:00 PM',
      liveStreamLink: null,
      linkAvailable: false
    };
    
  } catch (error) {
    Logger.log('Error getting next Sunday service settings: ' + error.toString());
    return null;
  }
}

/**
 * DEPRECATED: Old function for backward compatibility
 */
function getCommunityServiceSettings(communityName) {
  return getNextSundayServiceSettings(communityName);
}

/**
 * Get livestream link for a specific community
 */
function getCommunityLiveStreamLink(communityName) {
  var settings = getNextSundayServiceSettings(communityName);
  
  if (settings && settings.liveStreamLink && settings.linkAvailable) {
    return settings.liveStreamLink;
  }
  
  return null;
}
