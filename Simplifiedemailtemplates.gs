/**
 * SIMPLIFIED EMAIL TEMPLATES
 * Clean, professional design with blue/white/red color palette
 * Replaces old gradient/purple email design
 */

/**
 * Build confirmation email with simplified blue/white/red design
 * UPDATED: Much cleaner, simpler styling
 */
function buildConfirmationEmailHTML(formData, serviceSettings) {
  var html = '';
  
  // Header - SIMPLE BLUE
  html += '<div style="background: #1e3a8a; padding: 30px 20px; text-align: center;">';
  html += '<h1 style="color: white; margin: 0; font-size: 24px; font-family: Arial, sans-serif;">✓ Registration Confirmed</h1>';
  html += '<p style="color: white; margin: 10px 0 0 0; font-size: 14px; font-family: Arial, sans-serif;">' + formData.community + '</p>';
  html += '</div>';
  
  // Main content - WHITE BACKGROUND
  html += '<div style="padding: 30px 20px; font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background: white;">';
  
  // Success message - SIMPLE
  html += '<div style="background: #dbeafe; border-left: 4px solid #3b82f6; padding: 15px; margin-bottom: 25px;">';
  html += '<p style="margin: 0; color: #1e3a8a; font-size: 16px; font-weight: 600;">You\'re registered for Sunday service</p>';
  html += '</div>';
  
  // Service Information - CLEAN LAYOUT
  html += '<h2 style="margin: 0 0 15px 0; color: #1e3a8a; font-size: 18px; border-bottom: 2px solid #3b82f6; padding-bottom: 8px;">Service Information</h2>';
  
  html += '<table style="width: 100%; margin-bottom: 25px;">';
  
  // Sunday Date
  html += '<tr><td style="padding: 10px 0; color: #1e3a8a; font-weight: 600; width: 130px;">Sunday:</td>';
  html += '<td style="padding: 10px 0; color: #1e293b;">' + serviceSettings.sundayDate + '</td></tr>';
  
  // Service Time
  html += '<tr><td style="padding: 10px 0; color: #1e3a8a; font-weight: 600;">Time:</td>';
  html += '<td style="padding: 10px 0; color: #1e293b;">' + serviceSettings.serviceTime + '</td></tr>';
  
  // Speaker
  html += '<tr><td style="padding: 10px 0; color: #1e3a8a; font-weight: 600;">Speaker:</td>';
  html += '<td style="padding: 10px 0; color: #1e293b;">' + serviceSettings.speakerName + '</td></tr>';
  
  html += '</table>';
  
  // Registrants - SIMPLE TABLE
  html += '<h2 style="margin: 0 0 15px 0; color: #1e3a8a; font-size: 18px; border-bottom: 2px solid #3b82f6; padding-bottom: 8px;">Registered Attendees</h2>';
  
  html += '<table style="width: 100%; margin-bottom: 25px;">';
  
  for (var i = 0; i < formData.registrants.length; i++) {
    var person = formData.registrants[i];
    var bgColor = i % 2 === 0 ? '#ffffff' : '#dbeafe';
    
    html += '<tr style="background: ' + bgColor + ';">';
    html += '<td style="padding: 10px; border-bottom: 1px solid #e5e7eb;">';
    html += '<strong style="color: #1e3a8a;">' + person.firstName + ' ' + person.lastName + '</strong>';
    
    // Type badge - BLUE for Member, RED for Guest
    var badgeColor = person.type === 'Guest' ? '#dc2626' : '#3b82f6';
    html += '<span style="margin-left: 10px; padding: 3px 8px; background: ' + badgeColor + '; color: white; border-radius: 3px; font-size: 11px; font-weight: 600;">' + person.type + '</span>';
    
    html += '</td></tr>';
  }
  
  html += '</table>';
  
  // Livestream Link or Not Available Message
  if (serviceSettings.linkAvailable && serviceSettings.liveStreamLink) {
    // Simple blue button
    html += '<div style="text-align: center; margin: 25px 0;">';
    html += '<a href="' + serviceSettings.liveStreamLink + '" style="display: inline-block; background: #dc2626; color: white; padding: 12px 30px; text-decoration: none; border-radius: 5px; font-size: 16px; font-weight: 600;">Join Livestream</a>';
    html += '<p style="margin: 12px 0 0 0; color: #666; font-size: 13px;">Click to join on Sunday</p>';
    html += '</div>';
  } else {
    // Link not available message - SIMPLE
    html += '<div style="background: #dbeafe; border-left: 4px solid #3b82f6; padding: 15px; margin: 25px 0; text-align: center;">';
    html += '<p style="margin: 0; color: #1e3a8a; font-size: 14px; font-weight: 600;">Livestream link not yet available</p>';
    html += '<p style="margin: 8px 0 0 0; color: #1e3a8a; font-size: 13px;">The link for ' + serviceSettings.sundayDate + ' will be added soon. Check back tomorrow.</p>';
    html += '</div>';
  }
  
  // Footer - SIMPLE
  html += '<div style="margin-top: 30px; padding-top: 20px; border-top: 2px solid #e5e7eb; text-align: center;">';
  html += '<p style="margin: 0 0 8px 0; color: #666; font-size: 13px;">Questions? Contact your community coordinator</p>';
  html += '<p style="margin: 0; color: #999; font-size: 11px;">Sunday Service Registration System</p>';
  html += '</div>';
  
  html += '</div>';
  
  return html;
}

/**
 * Send confirmation email with simplified template
 */
function sendConfirmationEmail(formData, serviceSettings) {
  try {
    if (!formData.email) {
      Logger.log('ERROR: No email address provided');
      return false;
    }
    
    // Build HTML email with new simplified design
    var htmlBody = buildConfirmationEmailHTML(formData, serviceSettings);
    
    // Email subject - SIMPLE
    var subject = '✓ Sunday Service Registration - ' + formData.community;
    
    // Send email
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
