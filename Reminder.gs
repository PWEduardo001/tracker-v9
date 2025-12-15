/**
 * SUNDAY SERVICE REMINDER SYSTEM - WITH TIMEZONE SUPPORT
 * Sends reminder emails 3 hours before service starts
 * HANDLES DIFFERENT SERVICE TIMES AND TIMEZONES PER COMMUNITY
 * 
 * REQUIRES: CommunityTimezones.gs file
 * 
 * ⚠️ IMPORTANT: SERVICE TIMES ARE READ FROM EACH COMMUNITY'S SERVICE SETTINGS TABLE
 * - Each community has a "Service Settings" tab with year-long schedule
 * - Service times can vary Sunday to Sunday
 * - System reads the service time for the specific Sunday from that table
 * - Then calculates reminder time (3 hours before) based on that time
 * 
 * EXAMPLE:
 * - Dec 15: Boston service at 10:30 AM → Reminder at 7:30 AM
 * - Dec 22: Boston service at 9:00 AM → Reminder at 6:00 AM (different time!)
 * - Same Sunday: Philadelphia at 9:00 AM, Manhattan at 11:00 AM → Different reminders
 * 
 * HOW IT WORKS:
 * 1. When someone registers, gets community's timezone
 * 2. Gets service time from Service Settings table for that specific Sunday
 * 3. Interprets service time in THAT timezone
 * 4. Calculates reminder time (3 hours before) in THAT timezone
 * 5. Schedules trigger at correct UTC time
 * 6. Each community gets reminder at correct LOCAL time
 */

/**
 * Schedule reminder for a specific Sunday service FOR A SPECIFIC COMMUNITY
 * NOW WITH TIMEZONE SUPPORT
 * 
 * @param {string} sundayDate - The Sunday date (e.g., "December 15, 2024")
 * @param {string} serviceTime - Service time (e.g., "10:30 AM - 12:00 PM")
 * @param {string} community - Community name
 */
function scheduleReminderForSunday(sundayDate, serviceTime, community) {
  try {
    Logger.log('Scheduling reminder for: ' + sundayDate + ' (' + community + ')');
    
    // Get community timezone
    var timezone = getCommunityTimeZone(community);
    var timezoneDisplay = getCommunityTimeZoneDisplay(community);
    Logger.log('Timezone: ' + timezoneDisplay + ' (' + timezone + ')');
    
    // Check if reminder already scheduled for THIS COMMUNITY on THIS SUNDAY
    if (isReminderAlreadyScheduled(sundayDate, community)) {
      Logger.log('✓ Reminder already scheduled for ' + community + ' on ' + sundayDate);
      return;
    }
    
    // Parse service time to get start time
    var serviceStartTime = parseServiceStartTime(serviceTime);
    if (!serviceStartTime) {
      Logger.log('⚠️ Could not parse service time: ' + serviceTime);
      return;
    }
    
    Logger.log('Service starts at: ' + serviceStartTime + ' (' + timezoneDisplay + ')');
    
    // Calculate reminder time (3 hours before service) IN COMMUNITY'S TIMEZONE
    var reminderDateTime = calculateReminderTimeWithTimezone(sundayDate, serviceStartTime, community);
    
    if (!reminderDateTime) {
      Logger.log('⚠️ Could not calculate reminder time');
      return;
    }
    
    // Don't schedule if reminder time is in the past
    var now = new Date();
    if (reminderDateTime <= now) {
      Logger.log('⚠️ Reminder time is in the past, not scheduling');
      return;
    }
    
    // Show both local and UTC times
    var localReminderTime = getCommunityLocalTime(reminderDateTime, community);
    Logger.log('Reminder will be sent at:');
    Logger.log('  Local: ' + localReminderTime + ' (' + timezoneDisplay + ')');
    Logger.log('  UTC: ' + reminderDateTime.toUTCString());
    
    // Create trigger for that specific UTC time
    ScriptApp.newTrigger('sendCommunityReminder')
      .timeBased()
      .at(reminderDateTime)
      .create();
    
    // Store trigger info in Script Properties with community-specific key
    storeReminderInfo(sundayDate, community, reminderDateTime.toISOString());
    
    Logger.log('✓ Reminder scheduled for ' + community + ' on ' + sundayDate);
    
  } catch (error) {
    Logger.log('Error scheduling reminder: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
  }
}

/**
 * Calculate reminder time (3 hours before service) WITH TIMEZONE SUPPORT
 * This is the KEY function that makes timezones work correctly
 * 
 * @param {string} sundayDate - Date string (e.g., "December 15, 2024")
 * @param {string} serviceStartTime - Time string (e.g., "10:30 AM")
 * @param {string} community - Community name
 * @returns {Date} - Reminder datetime in UTC (for trigger scheduling)
 */
function calculateReminderTimeWithTimezone(sundayDate, serviceStartTime, community) {
  try {
    // Create service datetime in community's timezone
    var serviceDateTime = createCommunityDateTime(sundayDate, serviceStartTime, community);
    
    if (!serviceDateTime) {
      Logger.log('Could not create community datetime');
      return null;
    }
    
    // Subtract 3 hours for reminder
    var reminderTime = new Date(serviceDateTime.getTime() - (3 * 60 * 60 * 1000));
    
    return reminderTime;
    
  } catch (error) {
    Logger.log('Error calculating reminder time with timezone: ' + error.toString());
    return null;
  }
}

/**
 * Send reminder emails to registrants for upcoming Sunday service
 * This is called by EACH scheduled trigger (one per community per Sunday)
 */
function sendCommunityReminder() {
  try {
    logSeparator('SENDING COMMUNITY REMINDER');
    var startTime = new Date().getTime();
    
    // Get next Sunday date
    var nextSunday = getNextSunday();
    var sundayDate = nextSunday.formatted;
    
    Logger.log('Sending reminders for: ' + sundayDate);
    Logger.log('');
    
    // Find which community this trigger is for by checking Script Properties
    var targetCommunity = findCommunityForCurrentReminder(sundayDate);
    
    if (!targetCommunity) {
      Logger.log('⚠️ Could not determine target community');
      
      // Fallback: Send reminders for ALL communities
      sendRemindersForAllCommunities(sundayDate);
      return;
    }
    
    Logger.log('Target community: ' + targetCommunity);
    Logger.log('Timezone: ' + getCommunityTimeZoneDisplay(targetCommunity));
    
    // Process only this community
    var emailsSent = processReminderForCommunity(targetCommunity, sundayDate);
    
    var endTime = new Date().getTime();
    var duration = (endTime - startTime) / 1000;
    
    logSeparator('REMINDERS SENT');
    Logger.log('Community: ' + targetCommunity);
    Logger.log('Emails sent: ' + emailsSent);
    Logger.log('Time taken: ' + duration + ' seconds');
    logSeparator();
    
    // Clean up the Script Property for this reminder
    clearReminderInfo(sundayDate, targetCommunity);
    
  } catch (error) {
    Logger.log('ERROR in sendCommunityReminder: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    logSeparator('REMINDER SENDING FAILED');
  }
}

/**
 * Process reminder for a specific community
 * Returns number of emails sent
 */
function processReminderForCommunity(communityName, sundayDate) {
  try {
    Logger.log('Processing: ' + communityName);
    
    // Get registrations for this Sunday from this community
    var registrations = getRegistrationsForSunday(communityName, sundayDate);
    
    if (registrations.length === 0) {
      Logger.log('  No registrations for ' + sundayDate);
      return 0;
    }
    
    Logger.log('  Found ' + registrations.length + ' registrations');
    
    // Get service settings for reminder email
    var serviceSettings = getNextSundayServiceSettings(communityName);
    
    if (!serviceSettings) {
      Logger.log('  ⚠️ Could not get service settings');
      return 0;
    }
    
    // Group registrations by email (one email per person)
    var registrantsByEmail = groupRegistrationsByEmail(registrations);
    
    // Send reminder to each registrant
    var emailsSent = 0;
    var emails = Object.keys(registrantsByEmail);
    
    for (var j = 0; j < emails.length; j++) {
      var email = emails[j];
      var registrants = registrantsByEmail[email];
      
      var sent = sendReminderEmail(email, registrants, serviceSettings, communityName);
      if (sent) {
        emailsSent++;
      }
    }
    
    Logger.log('  ✓ Sent ' + emailsSent + ' reminder emails');
    return emailsSent;
    
  } catch (error) {
    Logger.log('  ✗ Error processing ' + communityName + ': ' + error.toString());
    return 0;
  }
}

/**
 * Fallback: Send reminders for all communities
 */
function sendRemindersForAllCommunities(sundayDate) {
  Logger.log('Sending reminders for ALL communities (fallback mode)');
  
  var communities = getCommunityNames();
  var totalEmailsSent = 0;
  
  for (var i = 0; i < communities.length; i++) {
    var communityName = communities[i];
    var emailsSent = processReminderForCommunity(communityName, sundayDate);
    totalEmailsSent += emailsSent;
  }
  
  Logger.log('Total emails sent: ' + totalEmailsSent);
}

/**
 * Get registrations for a specific Sunday from a community sheet
 */
function getRegistrationsForSunday(communityName, sundayDate) {
  try {
    var sheetId = getCommunitySheetId(communityName);
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var sheet = spreadsheet.getSheetByName('Registrations');
    
    if (!sheet) {
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    var registrations = [];
    
    // Skip header row
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowSundayDate = row[7]; // Column H: Sunday Date
      
      // Match Sunday date
      var rowDateStr = '';
      if (rowSundayDate instanceof Date) {
        rowDateStr = Utilities.formatDate(rowSundayDate, Session.getScriptTimeZone(), 'MMMM d, yyyy');
      } else {
        rowDateStr = rowSundayDate.toString().trim();
      }
      
      if (rowDateStr === sundayDate) {
        registrations.push({
          firstName: row[2],  // Column C
          lastName: row[3],   // Column D
          email: row[4]       // Column E
        });
      }
    }
    
    return registrations;
    
  } catch (error) {
    Logger.log('Error getting registrations: ' + error.toString());
    return [];
  }
}

/**
 * Group registrations by email address
 */
function groupRegistrationsByEmail(registrations) {
  var grouped = {};
  
  for (var i = 0; i < registrations.length; i++) {
    var reg = registrations[i];
    var email = reg.email;
    
    if (!grouped[email]) {
      grouped[email] = [];
    }
    
    grouped[email].push({
      firstName: reg.firstName,
      lastName: reg.lastName
    });
  }
  
  return grouped;
}

/**
 * Send reminder email to a registrant
 */
function sendReminderEmail(email, registrants, serviceSettings, communityName) {
  try {
    var htmlBody = buildReminderEmailHTML(registrants, serviceSettings, communityName);
    var subject = '⏰ Reminder: Sunday Service in 3 Hours - ' + communityName;
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
    
    Logger.log('    ✓ Sent reminder to ' + email);
    return true;
    
  } catch (error) {
    Logger.log('    ✗ Failed to send to ' + email + ': ' + error.toString());
    return false;
  }
}

/**
 * Build HTML reminder email
 */
function buildReminderEmailHTML(registrants, serviceSettings, communityName) {
  var html = '';
  
  // Header
  html += '<div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 40px 20px; text-align: center;">';
  html += '<h1 style="color: white; margin: 0; font-size: 28px; font-family: Arial, sans-serif;">⏰ Service Starting Soon!</h1>';
  html += '<p style="color: rgba(255,255,255,0.9); margin: 10px 0 0 0; font-size: 16px; font-family: Arial, sans-serif;">' + communityName + '</p>';
  html += '</div>';
  
  // Main content
  html += '<div style="padding: 30px 20px; font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">';
  
  // Reminder message
  html += '<div style="background: #fff3cd; border: 2px solid #ffc107; border-radius: 8px; padding: 20px; margin-bottom: 30px; text-align: center;">';
  html += '<p style="margin: 0; color: #856404; font-size: 18px; font-weight: 600;">Your Sunday service starts in 3 hours!</p>';
  html += '</div>';
  
  // Service Information
  html += '<div style="background: #f8f9fa; border-radius: 8px; padding: 25px; margin-bottom: 30px;">';
  html += '<h2 style="margin: 0 0 20px 0; color: #333; font-size: 20px; border-bottom: 2px solid #667eea; padding-bottom: 10px;">📅 Service Details</h2>';
  
  html += '<table style="width: 100%; border-collapse: collapse;">';
  
  html += '<tr>';
  html += '<td style="padding: 12px 0; color: #666; font-weight: 600; width: 140px;">Date:</td>';
  html += '<td style="padding: 12px 0; color: #333;">' + serviceSettings.sundayDate + '</td>';
  html += '</tr>';
  
  html += '<tr>';
  html += '<td style="padding: 12px 0; color: #666; font-weight: 600;">Time:</td>';
  html += '<td style="padding: 12px 0; color: #333;">' + serviceSettings.serviceTime + '</td>';
  html += '</tr>';
  
  html += '<tr>';
  html += '<td style="padding: 12px 0; color: #666; font-weight: 600;">Speaker:</td>';
  html += '<td style="padding: 12px 0; color: #333;">' + serviceSettings.speakerName + '</td>';
  html += '</tr>';
  
  html += '</table>';
  html += '</div>';
  
  // Registered attendees
  html += '<div style="background: #f8f9fa; border-radius: 8px; padding: 25px; margin-bottom: 30px;">';
  html += '<h2 style="margin: 0 0 20px 0; color: #333; font-size: 20px; border-bottom: 2px solid #667eea; padding-bottom: 10px;">👥 Your Registered Attendees</h2>';
  
  html += '<table style="width: 100%; border-collapse: collapse;">';
  
  for (var i = 0; i < registrants.length; i++) {
    var person = registrants[i];
    var bgColor = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    
    html += '<tr style="background: ' + bgColor + ';">';
    html += '<td style="padding: 12px; border-bottom: 1px solid #e9ecef;">';
    html += '<strong style="color: #333;">' + person.firstName + ' ' + person.lastName + '</strong>';
    html += '</td>';
    html += '</tr>';
  }
  
  html += '</table>';
  html += '</div>';
  
  // Livestream button or message
  if (serviceSettings.linkAvailable && serviceSettings.liveStreamLink) {
    html += '<div style="text-align: center; margin: 30px 0;">';
    html += '<a href="' + serviceSettings.liveStreamLink + '" style="display: inline-block; background: #ff0000; color: white; padding: 15px 40px; text-decoration: none; border-radius: 8px; font-size: 18px; font-weight: 600; box-shadow: 0 4px 15px rgba(255,0,0,0.3);">📺 Join Livestream</a>';
    html += '<p style="margin: 15px 0 0 0; color: #666; font-size: 14px;">Service starts in 3 hours - see you there!</p>';
    html += '</div>';
  } else {
    html += '<div style="background: #fff3cd; border: 2px solid #ffc107; border-radius: 8px; padding: 20px; margin: 30px 0; text-align: center;">';
    html += '<p style="margin: 0; color: #856404; font-size: 16px;">Livestream link will be available soon. Please contact your community coordinator.</p>';
    html += '</div>';
  }
  
  // Footer
  html += '<div style="margin-top: 40px; padding-top: 20px; border-top: 2px solid #e9ecef; text-align: center;">';
  html += '<p style="margin: 0 0 10px 0; color: #666; font-size: 14px;">Looking forward to worshipping with you!</p>';
  html += '<p style="margin: 0; color: #999; font-size: 12px;">This is an automated reminder from the Sunday Service Registration System</p>';
  html += '</div>';
  
  html += '</div>';
  
  return html;
}

// ==================== HELPER FUNCTIONS ====================

/**
 * Parse service start time from service time string
 * Example: "10:30 AM - 12:00 PM" → "10:30 AM"
 */
function parseServiceStartTime(serviceTime) {
  if (!serviceTime) return null;
  
  var timeStr = serviceTime.toString().trim();
  
  // Split on dash or hyphen
  var parts = timeStr.split(/[-–]/);
  if (parts.length > 0) {
    return parts[0].trim();
  }
  
  return timeStr;
}

/**
 * Format date for trigger name (no special characters)
 */
function formatDateForTriggerName(dateStr) {
  return dateStr.replace(/[^a-zA-Z0-9]/g, '_');
}

/**
 * Format community name for trigger (no special characters)
 */
function formatCommunityForTrigger(communityName) {
  return communityName.replace(/[^a-zA-Z0-9]/g, '_');
}

/**
 * Check if reminder already scheduled for THIS COMMUNITY on THIS SUNDAY
 */
function isReminderAlreadyScheduled(sundayDate, community) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var key = 'reminder_' + formatCommunityForTrigger(community) + '_' + formatDateForTriggerName(sundayDate);
  var value = scriptProperties.getProperty(key);
  return value !== null;
}

/**
 * Store reminder info in Script Properties (community-specific)
 */
function storeReminderInfo(sundayDate, community, reminderDateTime) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var key = 'reminder_' + formatCommunityForTrigger(community) + '_' + formatDateForTriggerName(sundayDate);
  scriptProperties.setProperty(key, reminderDateTime);
}

/**
 * Clear reminder info after sending
 */
function clearReminderInfo(sundayDate, community) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var key = 'reminder_' + formatCommunityForTrigger(community) + '_' + formatDateForTriggerName(sundayDate);
  scriptProperties.deleteProperty(key);
}

/**
 * Find which community this reminder is for
 */
function findCommunityForCurrentReminder(sundayDate) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var communities = getCommunityNames();
  
  // Check which community has a reminder scheduled for this Sunday around now
  var now = new Date();
  var thirtyMinutesAgo = new Date(now.getTime() - (30 * 60 * 1000));
  var thirtyMinutesLater = new Date(now.getTime() + (30 * 60 * 1000));
  
  for (var i = 0; i < communities.length; i++) {
    var community = communities[i];
    var key = 'reminder_' + formatCommunityForTrigger(community) + '_' + formatDateForTriggerName(sundayDate);
    var value = scriptProperties.getProperty(key);
    
    if (value) {
      var scheduledTime = new Date(value);
      
      // Check if scheduled time is within 30 minutes of now
      if (scheduledTime >= thirtyMinutesAgo && scheduledTime <= thirtyMinutesLater) {
        return community;
      }
    }
  }
  
  return null;
}

// ==================== MANAGEMENT FUNCTIONS ====================

/**
 * View all scheduled reminders (organized by community WITH TIMEZONES)
 */
function viewScheduledReminders() {
  logSeparator('SCHEDULED REMINDERS');
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var keys = scriptProperties.getKeys();
  var reminderKeys = [];
  
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].indexOf('reminder_') === 0) {
      reminderKeys.push(keys[i]);
    }
  }
  
  Logger.log('Total reminders scheduled: ' + reminderKeys.length);
  Logger.log('');
  
  // Group by community
  var communities = getCommunityNames();
  
  for (var i = 0; i < communities.length; i++) {
    var community = communities[i];
    var timezone = getCommunityTimeZoneDisplay(community);
    var communityReminders = [];
    
    for (var j = 0; j < reminderKeys.length; j++) {
      if (reminderKeys[j].indexOf(formatCommunityForTrigger(community)) !== -1) {
        communityReminders.push(reminderKeys[j]);
      }
    }
    
    if (communityReminders.length > 0) {
      Logger.log(community + ' (' + timezone + '):');
      for (var k = 0; k < communityReminders.length; k++) {
        var value = scriptProperties.getProperty(communityReminders[k]);
        var scheduledTime = new Date(value);
        var localTime = getCommunityLocalTime(scheduledTime, community);
        Logger.log('  • Local: ' + localTime);
        Logger.log('    UTC: ' + scheduledTime.toUTCString());
      }
      Logger.log('');
    }
  }
  
  logSeparator();
}

/**
 * Clear all reminder triggers (for testing)
 */
function clearAllReminderTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendCommunityReminder') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  
  Logger.log('Removed ' + removed + ' reminder trigger(s)');
  
  // Also clear Script Properties
  var scriptProperties = PropertiesService.getScriptProperties();
  var keys = scriptProperties.getKeys();
  var clearedProps = 0;
  
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].indexOf('reminder_') === 0) {
      scriptProperties.deleteProperty(keys[i]);
      clearedProps++;
    }
  }
  
  Logger.log('Cleared ' + clearedProps + ' reminder properties');
}

/**
 * Test the reminder system WITH TIMEZONES
 */
function testReminderSystem() {
  logSeparator('TESTING REMINDER SYSTEM WITH TIMEZONES');
  
  // Test 1: Parse service time
  Logger.log('TEST 1: Parse service time');
  var testTime = '10:30 AM - 12:00 PM';
  var startTime = parseServiceStartTime(testTime);
  Logger.log('Input: ' + testTime);
  Logger.log('Output: ' + startTime);
  Logger.log(startTime === '10:30 AM' ? '✓ PASS' : '✗ FAIL');
  Logger.log('');
  
  // Test 2: Calculate reminder time WITH TIMEZONE
  Logger.log('TEST 2: Calculate reminder time with timezone');
  var testDate = 'December 15, 2024';
  
  Logger.log('Boston (Eastern Time):');
  var bostonReminder = calculateReminderTimeWithTimezone(testDate, '10:30 AM', 'Boston Family Church');
  var bostonLocal = getCommunityLocalTime(bostonReminder, 'Boston Family Church');
  Logger.log('  Service: 10:30 AM ET');
  Logger.log('  Reminder: ' + bostonLocal + ' ET');
  Logger.log('  UTC: ' + bostonReminder.toUTCString());
  Logger.log('');
  
  Logger.log('Belvedere (Pacific Time):');
  var belvedereReminder = calculateReminderTimeWithTimezone(testDate, '10:30 AM', 'Belvedere Family Church');
  var belvedereLocal = getCommunityLocalTime(belvedereReminder, 'Belvedere Family Church');
  Logger.log('  Service: 10:30 AM PT');
  Logger.log('  Reminder: ' + belvedereLocal + ' PT');
  Logger.log('  UTC: ' + belvedereReminder.toUTCString());
  Logger.log('');
  
  // Test 3: Verify 3-hour offset
  var bostonServiceTime = new Date(bostonReminder.getTime() + (3 * 60 * 60 * 1000));
  var bostonServiceLocal = getCommunityLocalTime(bostonServiceTime, 'Boston Family Church');
  Logger.log('TEST 3: Verify 3-hour offset');
  Logger.log('Boston reminder + 3 hours = ' + bostonServiceLocal);
  Logger.log(bostonServiceLocal === '10:30 AM' ? '✓ PASS' : '✗ FAIL');
  
  logSeparator('TEST COMPLETE');
}
