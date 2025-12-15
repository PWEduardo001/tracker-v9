/**
 * FORM HANDLER - GOOGLE FORMS VERSION
 * Sends instant HTML confirmation emails when Google Form is submitted
 * UPDATED: Type column completely removed - no more Member/Guest distinction
 * 
 * TRIGGER: Form Submit (runs automatically when someone submits the Google Form)
 */

/**
 * Main function triggered when Google Form is submitted
 */
function onFormSubmit(e) {
  try {
    logSeparator('FORM SUBMISSION STARTED');
    
    if (!e) {
      Logger.log('ERROR: No event object');
      return;
    }
    
    Logger.log('Event object type: ' + typeof e);
    Logger.log('Event has namedValues: ' + (e.namedValues ? 'YES' : 'NO'));
    Logger.log('Event has values: ' + (e.values ? 'YES' : 'NO'));
    
    var formData = parseGoogleFormSubmission(e);
    
    if (!formData.community) {
      Logger.log('ERROR: No community selected');
      Logger.log('Available data: ' + JSON.stringify(e.namedValues || e.values || {}));
      return;
    }
    
    Logger.log('Community: ' + formData.community);
    Logger.log('Number of registrants: ' + formData.registrants.length);
    
    var serviceSettings = getNextSundayServiceSettings(formData.community);
    
    if (!serviceSettings) {
      Logger.log('WARNING: Could not get service settings for ' + formData.community);
      serviceSettings = {
        sundayDate: getNextSunday().formatted,
        speakerName: 'TBA',
        serviceTime: '10:00 AM - 12:00 PM',
        liveStreamLink: SETTINGS.LIVE_STREAM_LINK,
        linkAvailable: true
      };
    }
    
    Logger.log('Sunday Date: ' + serviceSettings.sundayDate);
    Logger.log('Speaker: ' + serviceSettings.speakerName);
    Logger.log('Link Available: ' + serviceSettings.linkAvailable);
    
    var emailSent = sendConfirmationEmail(formData, serviceSettings);
    
    if (emailSent) {
      Logger.log('✓ Email sent successfully');
      logSeparator('FORM SUBMISSION COMPLETE');
    } else {
      Logger.log('✗ Email failed to send');
      logSeparator('FORM SUBMISSION FAILED');
    }
    
  } catch (error) {
    Logger.log('ERROR in onFormSubmit: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    logSeparator('FORM SUBMISSION FAILED');
  }
}

/**
 * Parse Google Form submission into structured data
 * UPDATED: No Type field parsing
 */
function parseGoogleFormSubmission(e) {
  var formData = {
    community: null,
    email: null,
    registrants: []
  };
  
  var responses = e.namedValues || {};
  
  Logger.log('Form responses received:');
  Logger.log(JSON.stringify(responses));
  
  function getValue(key) {
    if (responses[key] && responses[key].length > 0) {
      return responses[key][0];
    }
    return null;
  }
  
  formData.community = getValue('Which community are you registering for?');
  formData.email = getValue('Email Address');
  
  // Person 1 - NO TYPE
  var person1FirstName = getValue('First Name');
  var person1LastName = getValue('Last Name');
  
  if (person1FirstName && person1LastName) {
    formData.registrants.push({
      firstName: person1FirstName,
      lastName: person1LastName,
      email: formData.email
    });
  }
  
  // Person 2 - NO TYPE
  var addPerson2 = getValue('Will anyone else be joining with you?');
  if (addPerson2 === 'Yes') {
    var person2FirstName = getValue('#2 First Name');
    var person2LastName = getValue('#2 Last Name');
    
    if (person2FirstName && person2LastName) {
      formData.registrants.push({
        firstName: person2FirstName,
        lastName: person2LastName,
        email: formData.email
      });
    }
  }
  
  // Person 3 - NO TYPE
  var addPerson3 = getValue('#2 Add another person?');
  if (addPerson3 === 'Yes') {
    var person3FirstName = getValue('#3 First Name');
    var person3LastName = getValue('#3 Last Name');
    
    if (person3FirstName && person3LastName) {
      formData.registrants.push({
        firstName: person3FirstName,
        lastName: person3LastName,
        email: formData.email
      });
    }
  }
  
  // Person 4 - NO TYPE
  var addPerson4 = getValue('#3 Add another person?');
  if (addPerson4 === 'Yes') {
    var person4FirstName = getValue('#4 First Name');
    var person4LastName = getValue('#4 Last Name');
    
    if (person4FirstName && person4LastName) {
      formData.registrants.push({
        firstName: person4FirstName,
        lastName: person4LastName,
        email: formData.email
      });
    }
  }
  
  // Person 5 - NO TYPE
  var addPerson5 = getValue('#4 Add another person?');
  if (addPerson5 === 'Yes') {
    var person5FirstName = getValue('#5 First Name');
    var person5LastName = getValue('#5 Last Name');
    
    if (person5FirstName && person5LastName) {
      formData.registrants.push({
        firstName: person5FirstName,
        lastName: person5LastName,
        email: formData.email
      });
    }
  }
  
  Logger.log('Parsed form data:');
  Logger.log('  Community: ' + formData.community);
  Logger.log('  Email: ' + formData.email);
  Logger.log('  Registrants: ' + formData.registrants.length);
  
  return formData;
}

/**
 * Send confirmation email with HTML template
 */
function sendConfirmationEmail(formData, serviceSettings) {
  try {
    if (!formData.email) {
      Logger.log('ERROR: No email address provided');
      return false;
    }
    
    var htmlBody = buildConfirmationEmailHTML(formData, serviceSettings);
    
    var subject = '✓ Sunday Service Registration Confirmed - ' + formData.community;
    
    MailApp.sendEmail({
      to: formData.email,
      subject: subject,
      htmlBody: htmlBody
    });
    
    Logger.log('Email sent to: ' + formData.email);
    Logger.log('Subject: ' + subject);
    
    return true;
    
  } catch (error) {
    Logger.log('ERROR sending email: ' + error.toString());
    return false;
  }
}

/**
 * Build beautiful HTML email
 * UPDATED: No Member/Guest badge
 */
function buildConfirmationEmailHTML(formData, serviceSettings) {
  var html = '';
  
  // Header with gradient
  html += '<div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 40px 20px; text-align: center;">';
  html += '<h1 style="color: white; margin: 0; font-size: 28px; font-family: Arial, sans-serif;">🙏 Registration Confirmed</h1>';
  html += '<p style="color: rgba(255,255,255,0.9); margin: 10px 0 0 0; font-size: 16px; font-family: Arial, sans-serif;">' + formData.community + '</p>';
  html += '</div>';
  
  // Main content
  html += '<div style="padding: 30px 20px; font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">';
  
  // Success message
  html += '<div style="background: #d4edda; border: 2px solid #28a745; border-radius: 8px; padding: 20px; margin-bottom: 30px;">';
  html += '<p style="margin: 0; color: #155724; font-size: 16px; font-weight: 600;">✓ You\'re all set for Sunday service!</p>';
  html += '</div>';
  
  // Service Information
  html += '<div style="background: #f8f9fa; border-radius: 8px; padding: 25px; margin-bottom: 30px;">';
  html += '<h2 style="margin: 0 0 20px 0; color: #333; font-size: 20px; border-bottom: 2px solid #667eea; padding-bottom: 10px;">📅 Service Information</h2>';
  
  html += '<table style="width: 100%; border-collapse: collapse;">';
  
  // Sunday Date
  html += '<tr>';
  html += '<td style="padding: 12px 0; color: #666; font-weight: 600; width: 140px;">Sunday Date:</td>';
  html += '<td style="padding: 12px 0; color: #333;">' + serviceSettings.sundayDate + '</td>';
  html += '</tr>';
  
  // Service Time
  html += '<tr>';
  html += '<td style="padding: 12px 0; color: #666; font-weight: 600;">Service Time:</td>';
  html += '<td style="padding: 12px 0; color: #333;">' + serviceSettings.serviceTime + '</td>';
  html += '</tr>';
  
  // Speaker
  html += '<tr>';
  html += '<td style="padding: 12px 0; color: #666; font-weight: 600;">Speaker:</td>';
  html += '<td style="padding: 12px 0; color: #333;">' + serviceSettings.speakerName + '</td>';
  html += '</tr>';
  
  html += '</table>';
  html += '</div>';
  
  // Registrants - NO MEMBER/GUEST BADGES
  html += '<div style="background: #f8f9fa; border-radius: 8px; padding: 25px; margin-bottom: 30px;">';
  html += '<h2 style="margin: 0 0 20px 0; color: #333; font-size: 20px; border-bottom: 2px solid #667eea; padding-bottom: 10px;">👥 Registered Attendees</h2>';
  
  html += '<table style="width: 100%; border-collapse: collapse;">';
  
  for (var i = 0; i < formData.registrants.length; i++) {
    var person = formData.registrants[i];
    var bgColor = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    
    html += '<tr style="background: ' + bgColor + ';">';
    html += '<td style="padding: 12px; border-bottom: 1px solid #e9ecef;">';
    html += '<strong style="color: #333;">' + person.firstName + ' ' + person.lastName + '</strong>';
    html += '</td>';
    html += '</tr>';
  }
  
  html += '</table>';
  html += '</div>';
  
  // Livestream Link or Not Available Message
  if (serviceSettings.linkAvailable && serviceSettings.liveStreamLink) {
    // YouTube button
    html += '<div style="text-align: center; margin: 30px 0;">';
    html += '<a href="' + serviceSettings.liveStreamLink + '" style="display: inline-block; background: #ff0000; color: white; padding: 15px 40px; text-decoration: none; border-radius: 8px; font-size: 18px; font-weight: 600; box-shadow: 0 4px 15px rgba(255,0,0,0.3);">📺 Join Livestream</a>';
    html += '<p style="margin: 15px 0 0 0; color: #666; font-size: 14px;">Click the button above on Sunday to join the service</p>';
    html += '</div>';
  } else {
    // Link not available message
    html += '<div style="background: #fff3cd; border: 2px solid #ffc107; border-radius: 8px; padding: 20px; margin: 30px 0; text-align: center;">';
    html += '<p style="margin: 0; color: #856404; font-size: 16px; font-weight: 600;">📅 Livestream Link Not Yet Available</p>';
    html += '<p style="margin: 10px 0 0 0; color: #856404; font-size: 14px;">The livestream link for ' + serviceSettings.sundayDate + ' has not been added yet. Please check back tomorrow or contact your community coordinator.</p>';
    html += '</div>';
  }
  
  // Footer
  html += '<div style="margin-top: 40px; padding-top: 20px; border-top: 2px solid #e9ecef; text-align: center;">';
  html += '<p style="margin: 0 0 10px 0; color: #666; font-size: 14px;">Questions? Contact your community coordinator</p>';
  html += '<p style="margin: 0; color: #999; font-size: 12px;">This is an automated confirmation email from the Sunday Service Registration System</p>';
  html += '</div>';
  
  html += '</div>';
  
  return html;
}

function installFormTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(spreadsheet)
    .onFormSubmit()
    .create();
  
  Logger.log('✓ Form submit trigger installed');
  
  Browser.msgBox(
    'Form Trigger Installed',
    'The form submit trigger has been installed.\n\n' +
    'Confirmation emails will now be sent automatically when someone submits the Google Form.\n\n' +
    'Emails will arrive in less than 1 second!',
    Browser.Buttons.OK
  );
}

function removeFormTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  
  Logger.log('Removed ' + removed + ' form trigger(s)');
  
  Browser.msgBox(
    'Form Trigger Removed',
    'The form submit trigger has been removed.\n\n' +
    'Automatic emails are now disabled.\n' +
    'You can reinstall from the menu.',
    Browser.Buttons.OK
  );
}

function testFormHandler() {
  Logger.log('===== TESTING FORM HANDLER =====');
  
  var testData = {
    community: 'Boston Family Church',
    email: Session.getActiveUser().getEmail(),
    registrants: [
      {
        firstName: 'Test',
        lastName: 'User',
        email: Session.getActiveUser().getEmail()
      }
    ]
  };
  
  var serviceSettings = getNextSundayServiceSettings(testData.community);
  
  if (!serviceSettings) {
    Logger.log('WARNING: Could not get service settings');
    serviceSettings = {
      sundayDate: getNextSunday().formatted,
      speakerName: 'TBA',
      serviceTime: '10:00 AM - 12:00 PM',
      liveStreamLink: SETTINGS.LIVE_STREAM_LINK,
      linkAvailable: true
    };
  }
  
  Logger.log('Test data:');
  Logger.log('  Community: ' + testData.community);
  Logger.log('  Email: ' + testData.email);
  Logger.log('  Sunday: ' + serviceSettings.sundayDate);
  
  var result = sendConfirmationEmail(testData, serviceSettings);
  
  if (result) {
    Logger.log('✓ Test email sent successfully');
    Browser.msgBox(
      'Test Successful!',
      'Test email sent to: ' + testData.email + '\n\n' +
      'Check your inbox to see the confirmation email.\n\n' +
      'If it looks good, the system is working!',
      Browser.Buttons.OK
    );
  } else {
    Logger.log('✗ Test email failed');
    Browser.msgBox(
      'Test Failed',
      'Could not send test email.\n\n' +
      'Check the execution log for errors.',
      Browser.Buttons.OK
    );
  }
  
  Logger.log('===== TEST COMPLETE =====');
}
