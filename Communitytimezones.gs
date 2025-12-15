/**
 * COMMUNITY TIMEZONE CONFIGURATION
 * Stores time zone for each community for accurate reminder scheduling
 * 
 * WHY THIS IS IMPORTANT:
 * - Boston says "10:30 AM" → That's 10:30 AM Eastern Time
 * - Belvedere says "10:30 AM" → That's 10:30 AM Pacific Time
 * - Without timezone info, reminders would be scheduled incorrectly!
 * 
 * TIME ZONE FORMAT:
 * Use IANA time zone identifiers (e.g., "America/New_York")
 * Full list: https://en.wikipedia.org/wiki/List_of_tz_database_time_zones
 */

// ==================== COMMUNITY TIME ZONES ====================

/**
 * Time zone configuration for each community
 * 
 * ALL COMMUNITIES ARE IN EASTERN TIME (ET)
 * 
 * NOTE: If any community is in a different timezone in the future, you can change it here.
 * 
 * COMMON US TIME ZONES:
 * - Eastern Time: "America/New_York"
 * - Central Time: "America/Chicago"
 * - Mountain Time: "America/Denver"
 * - Pacific Time: "America/Los_Angeles"
 * - Alaska Time: "America/Anchorage"
 * - Hawaii Time: "Pacific/Honolulu"
 */
const COMMUNITY_TIMEZONES = {
  'Belvedere Family Church': 'America/New_York',           // Eastern Time (ET)
  'Boston Family Church': 'America/New_York',              // Eastern Time (ET)
  'Connecticut Family Church': 'America/New_York',         // Eastern Time (ET)
  'Elizabeth Family Church': 'America/New_York',           // Eastern Time (ET)
  'Manhattan Family Church': 'America/New_York',           // Eastern Time (ET)
  'New Jersey Family Church': 'America/New_York',          // Eastern Time (ET)
  'Philadelphia Family Church': 'America/New_York',        // Eastern Time (ET)
  'Worcester Family Church': 'America/New_York'            // Eastern Time (ET)
};

/**
 * Get time zone for a specific community
 * @param {string} communityName - Full community name
 * @returns {string} - IANA time zone identifier (e.g., "America/New_York")
 */
function getCommunityTimeZone(communityName) {
  var timezone = COMMUNITY_TIMEZONES[communityName];
  
  if (!timezone) {
    Logger.log('⚠️ No timezone configured for: ' + communityName);
    Logger.log('   Using default: America/New_York');
    return 'America/New_York'; // Default to Eastern Time
  }
  
  return timezone;
}

/**
 * Get readable time zone name for display
 * @param {string} communityName - Full community name
 * @returns {string} - Human-readable timezone (e.g., "Eastern Time (ET)")
 */
function getCommunityTimeZoneDisplay(communityName) {
  var timezone = getCommunityTimeZone(communityName);
  
  var displayNames = {
    'America/New_York': 'Eastern Time (ET)',
    'America/Chicago': 'Central Time (CT)',
    'America/Denver': 'Mountain Time (MT)',
    'America/Los_Angeles': 'Pacific Time (PT)',
    'America/Anchorage': 'Alaska Time (AKT)',
    'Pacific/Honolulu': 'Hawaii Time (HT)'
  };
  
  return displayNames[timezone] || timezone;
}

/**
 * Calculate local time for a community
 * Takes a Date object and converts it to that community's local time string
 * 
 * @param {Date} dateTime - The date/time to convert
 * @param {string} communityName - Community name
 * @returns {string} - Formatted local time (e.g., "7:30 AM")
 */
function getCommunityLocalTime(dateTime, communityName) {
  var timezone = getCommunityTimeZone(communityName);
  
  try {
    return Utilities.formatDate(dateTime, timezone, 'h:mm a');
  } catch (error) {
    Logger.log('Error formatting date for timezone: ' + error.toString());
    return Utilities.formatDate(dateTime, Session.getScriptTimeZone(), 'h:mm a');
  }
}

/**
 * Calculate local date for a community
 * 
 * @param {Date} dateTime - The date/time to convert
 * @param {string} communityName - Community name
 * @returns {string} - Formatted local date (e.g., "December 15, 2024")
 */
function getCommunityLocalDate(dateTime, communityName) {
  var timezone = getCommunityTimeZone(communityName);
  
  try {
    return Utilities.formatDate(dateTime, timezone, 'MMMM d, yyyy');
  } catch (error) {
    Logger.log('Error formatting date for timezone: ' + error.toString());
    return Utilities.formatDate(dateTime, Session.getScriptTimeZone(), 'MMMM d, yyyy');
  }
}

/**
 * Create a Date object from date string and time string in community's timezone
 * This is CRITICAL for accurate reminder scheduling
 * 
 * @param {string} dateStr - Date string (e.g., "December 15, 2024")
 * @param {string} timeStr - Time string (e.g., "10:30 AM")
 * @param {string} communityName - Community name
 * @returns {Date} - Date object in community's timezone
 */
function createCommunityDateTime(dateStr, timeStr, communityName) {
  try {
    var timezone = getCommunityTimeZone(communityName);
    
    // Parse the date
    var dateObj = new Date(dateStr);
    
    // Parse the time
    var timeParts = timeStr.match(/(\d+):(\d+)\s*(AM|PM)/i);
    if (!timeParts) {
      Logger.log('Could not parse time: ' + timeStr);
      return null;
    }
    
    var hours = parseInt(timeParts[1]);
    var minutes = parseInt(timeParts[2]);
    var meridiem = timeParts[3].toUpperCase();
    
    // Convert to 24-hour format
    if (meridiem === 'PM' && hours !== 12) {
      hours += 12;
    } else if (meridiem === 'AM' && hours === 12) {
      hours = 0;
    }
    
    // Create a date string in the format: "2024-12-15T10:30:00"
    var year = dateObj.getFullYear();
    var month = String(dateObj.getMonth() + 1).padStart(2, '0');
    var day = String(dateObj.getDate()).padStart(2, '0');
    var hoursStr = String(hours).padStart(2, '0');
    var minutesStr = String(minutes).padStart(2, '0');
    
    var dateTimeStr = year + '-' + month + '-' + day + 'T' + hoursStr + ':' + minutesStr + ':00';
    
    // Parse this date in the community's timezone
    // Note: JavaScript Date parsing can be tricky with timezones
    // We need to create a date that represents this local time
    
    // Get the timezone offset for this community at this date/time
    var tempDate = new Date(dateTimeStr + 'Z'); // Create in UTC first
    var formatter = Utilities.formatDate(tempDate, timezone, 'Z'); // Get offset like "-05:00"
    
    // Parse the offset
    var offsetMatch = formatter.match(/([+-])(\d{2}):(\d{2})/);
    if (offsetMatch) {
      var offsetSign = offsetMatch[1] === '+' ? 1 : -1;
      var offsetHours = parseInt(offsetMatch[2]);
      var offsetMinutes = parseInt(offsetMatch[3]);
      var totalOffsetMinutes = offsetSign * (offsetHours * 60 + offsetMinutes);
      
      // Adjust the date by the offset to get true local time
      var localDate = new Date(tempDate.getTime() - (totalOffsetMinutes * 60 * 1000));
      
      return localDate;
    }
    
    // Fallback: Use simpler approach
    dateObj.setHours(hours);
    dateObj.setMinutes(minutes);
    dateObj.setSeconds(0);
    dateObj.setMilliseconds(0);
    
    return dateObj;
    
  } catch (error) {
    Logger.log('Error creating community date/time: ' + error.toString());
    return null;
  }
}

/**
 * View all community timezones (for verification)
 */
function viewCommunityTimezones() {
  logSeparator('COMMUNITY TIME ZONES');
  
  var communities = getCommunityNames();
  
  Logger.log('Configured time zones for ' + communities.length + ' communities:');
  Logger.log('');
  
  // Group by timezone
  var byTimezone = {};
  
  for (var i = 0; i < communities.length; i++) {
    var community = communities[i];
    var timezone = getCommunityTimeZone(community);
    var display = getCommunityTimeZoneDisplay(community);
    
    if (!byTimezone[display]) {
      byTimezone[display] = [];
    }
    byTimezone[display].push(community);
  }
  
  // Display grouped by timezone
  var timezones = Object.keys(byTimezone);
  for (var i = 0; i < timezones.length; i++) {
    var tz = timezones[i];
    var communities = byTimezone[tz];
    
    Logger.log(tz + ':');
    for (var j = 0; j < communities.length; j++) {
      Logger.log('  • ' + communities[j]);
    }
    Logger.log('');
  }
  
  // Show example: What time is "10:30 AM" in each timezone?
  Logger.log('EXAMPLE: Current time in each timezone:');
  Logger.log('');
  
  var now = new Date();
  var allCommunities = getCommunityNames();
  
  for (var i = 0; i < allCommunities.length; i++) {
    var community = allCommunities[i];
    var localTime = getCommunityLocalTime(now, community);
    var timezone = getCommunityTimeZoneDisplay(community);
    Logger.log(community + ' (' + timezone + '): ' + localTime);
  }
  
  logSeparator();
}

/**
 * Test timezone conversion
 */
function testTimezoneConversion() {
  logSeparator('TESTING TIMEZONE CONVERSION');
  
  // Test 1: Create date/time in different timezones
  Logger.log('TEST 1: Creating service times in community timezones');
  Logger.log('Sunday, December 15, 2024 at 10:30 AM');
  Logger.log('');
  
  var testDate = 'December 15, 2024';
  var testTime = '10:30 AM';
  
  Logger.log('Boston (Eastern):');
  var bostonDateTime = createCommunityDateTime(testDate, testTime, 'Boston Family Church');
  Logger.log('  Local: ' + testTime);
  Logger.log('  UTC: ' + bostonDateTime.toUTCString());
  Logger.log('');
  
  Logger.log('Belvedere (Pacific):');
  var belvedereDateTime = createCommunityDateTime(testDate, testTime, 'Belvedere Family Church');
  Logger.log('  Local: ' + testTime);
  Logger.log('  UTC: ' + belvedereDateTime.toUTCString());
  Logger.log('');
  
  // Test 2: Calculate 3 hours before
  Logger.log('TEST 2: Reminder times (3 hours before)');
  Logger.log('');
  
  var bostonReminder = new Date(bostonDateTime.getTime() - (3 * 60 * 60 * 1000));
  Logger.log('Boston reminder:');
  Logger.log('  Local: ' + getCommunityLocalTime(bostonReminder, 'Boston Family Church'));
  Logger.log('  UTC: ' + bostonReminder.toUTCString());
  Logger.log('');
  
  var belvedereReminder = new Date(belvedereDateTime.getTime() - (3 * 60 * 60 * 1000));
  Logger.log('Belvedere reminder:');
  Logger.log('  Local: ' + getCommunityLocalTime(belvedereReminder, 'Belvedere Family Church'));
  Logger.log('  UTC: ' + belvedereReminder.toUTCString());
  Logger.log('');
  
  // Test 3: Time difference
  var timeDiff = (bostonDateTime.getTime() - belvedereDateTime.getTime()) / (60 * 60 * 1000);
  Logger.log('TEST 3: Time difference');
  Logger.log('Boston 10:30 AM ET vs Belvedere 10:30 AM PT:');
  Logger.log('  Difference: ' + Math.abs(timeDiff) + ' hours');
  Logger.log('  (This should be 3 hours)');
  Logger.log('');
  
  if (Math.abs(timeDiff) === 3) {
    Logger.log('✓ PASS - Timezone conversion working correctly!');
  } else {
    Logger.log('✗ FAIL - Timezone conversion needs adjustment');
  }
  
  logSeparator('TEST COMPLETE');
}

/**
 * Show what time a reminder will fire in each community's local time
 */
function showReminderTimes() {
  logSeparator('REMINDER TIMES BY COMMUNITY');
  
  var testDate = 'December 15, 2024';
  
  Logger.log('For Sunday, December 15, 2024');
  Logger.log('');
  Logger.log('Service Time → Reminder Time (Local)');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  // Example service times
  var serviceTimes = {
    'Boston Family Church': '10:30 AM',
    'Manhattan Family Church': '11:00 AM',
    'Philadelphia Family Church': '9:00 AM',
    'Connecticut Family Church': '10:00 AM',
    'Belvedere Family Church': '10:30 AM',
    'Worcester Family Church': '10:30 AM',
    'New Jersey Family Church': '10:00 AM',
    'Elizabeth Family Church': '10:00 AM'
  };
  
  var communities = Object.keys(serviceTimes);
  
  for (var i = 0; i < communities.length; i++) {
    var community = communities[i];
    var serviceTime = serviceTimes[community];
    var timezone = getCommunityTimeZoneDisplay(community);
    
    // Create service datetime
    var serviceDateTime = createCommunityDateTime(testDate, serviceTime, community);
    
    // Calculate reminder time
    var reminderDateTime = new Date(serviceDateTime.getTime() - (3 * 60 * 60 * 1000));
    var reminderTime = getCommunityLocalTime(reminderDateTime, community);
    
    Logger.log(community + ' (' + timezone + ')');
    Logger.log('  Service: ' + serviceTime);
    Logger.log('  Reminder: ' + reminderTime);
    Logger.log('  UTC: ' + reminderDateTime.toUTCString());
    Logger.log('');
  }
  
  logSeparator();
}

// ==================== CONFIGURATION HELP ====================

/**
 * Instructions for updating timezones
 */
function howToUpdateTimezones() {
  Logger.log('═══════════════════════════════════════════════════════');
  Logger.log('HOW TO UPDATE COMMUNITY TIME ZONES');
  Logger.log('═══════════════════════════════════════════════════════');
  Logger.log('');
  Logger.log('1. Open CommunityTimezones.gs file');
  Logger.log('');
  Logger.log('2. Find the COMMUNITY_TIMEZONES object');
  Logger.log('');
  Logger.log('3. Update the timezone for each community');
  Logger.log('');
  Logger.log('COMMON US TIME ZONES:');
  Logger.log('  Eastern Time:  "America/New_York"');
  Logger.log('  Central Time:  "America/Chicago"');
  Logger.log('  Mountain Time: "America/Denver"');
  Logger.log('  Pacific Time:  "America/Los_Angeles"');
  Logger.log('  Alaska Time:   "America/Anchorage"');
  Logger.log('  Hawaii Time:   "Pacific/Honolulu"');
  Logger.log('');
  Logger.log('EXAMPLE:');
  Logger.log('  \'Boston Family Church\': \'America/New_York\',');
  Logger.log('  \'Belvedere Family Church\': \'America/Los_Angeles\',');
  Logger.log('');
  Logger.log('4. Save the file');
  Logger.log('');
  Logger.log('5. Run viewCommunityTimezones() to verify');
  Logger.log('');
  Logger.log('Full timezone list:');
  Logger.log('https://en.wikipedia.org/wiki/List_of_tz_database_time_zones');
  Logger.log('');
  Logger.log('═══════════════════════════════════════════════════════');
}
