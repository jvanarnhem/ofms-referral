/**
 * @OnlyCurrentDoc
 * Enhanced School Discipline Referral System
 */

const SPREADSHEET_ID = '1oMgE9AdGDR_YgVTkeOhYrENDo0K1lt1sWZxz3Hzko4A';
const RATE_LIMIT_CACHE = {};

const FIELD_LIMITS = {
  staffName: 100,
  staffEmail: 100,
  descriptionOfInfraction: 2000,
  adminComments: 2000,
  comments: 1000,
  infractionOtherText: 500,
  actionsTakenOtherText: 500
};

// Cache for spreadsheet headers to handle column reordering
let HEADER_CACHE = {};

/**
 * Format date string from form (YYYY-MM-DD) to "Mar 1, 2025" format
 */
function formatDateDisplay(dateValue) {
  if (!dateValue) return '';

  let date;
  if (typeof dateValue === 'string' && dateValue.includes('-')) {
    // Handle YYYY-MM-DD format from HTML date input
    const [year, month, day] = dateValue.split('-');
    date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
  } else {
    date = dateValue instanceof Date ? dateValue : new Date(dateValue);
  }

  if (isNaN(date.getTime())) return 'Invalid date';

  const options = {
    month: 'short',
    day: 'numeric',
    year: 'numeric'
  };
  return date.toLocaleDateString('en-US', options);
}

/**
 * Format time string from form (HH:MM) to "2:30 PM" format
 */
function formatTimeDisplay(timeValue) {
  if (!timeValue) return '';

  let date;
  if (typeof timeValue === 'string' && timeValue.includes(':') && !timeValue.includes(' ')) {
    // Handle HH:MM format from HTML time input
    const [hours, minutes] = timeValue.split(':');
    date = new Date(2000, 0, 1, parseInt(hours), parseInt(minutes)); // Use dummy date
  } else {
    date = timeValue instanceof Date ? timeValue : new Date(timeValue);
  }

  if (isNaN(date.getTime())) return '';

  const options = {
    hour: 'numeric',
    minute: '2-digit',
    hour12: true
  };
  return date.toLocaleTimeString('en-US', options);
}

// --- ROUTING & SERVING HTML ---
function doGet(e) {
  try {
    if (e.parameter.page === 'admin' && e.parameter.id) {
      const template = HtmlService.createTemplateFromFile('admin.html');
      template.submissionId = e.parameter.id;
      return template.evaluate()
        .setTitle('Admin Referral Review')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
    } else {
      const template = HtmlService.createTemplateFromFile('forms2.html');
      template.scriptUrl = ScriptApp.getService().getUrl();
      return template.evaluate()
        .setTitle('School Discipline Referral Form')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
    }
  } catch (error) {
    Logger.log('Error in doGet: ' + error.message);
    return HtmlService.createHtmlOutput('<h1>Service Temporarily Unavailable</h1><p>Please try again later.</p>');
  }
}

// --- FIXED: DYNAMIC HEADER MAPPING FUNCTIONS ---
function getSheetHeaders(sheetName) {
  try {
    const cacheKey = `headers_${sheetName}`;
    if (HEADER_CACHE[cacheKey]) {
      return HEADER_CACHE[cacheKey];
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    HEADER_CACHE[cacheKey] = headers;

    return headers;
  } catch (error) {
    Logger.log('Error getting sheet headers: ' + error.message);
    throw error;
  }
}

function getColumnIndex(headers, columnName) {
  const index = headers.indexOf(columnName);
  if (index === -1) {
    Logger.log(`Warning: Column "${columnName}" not found in headers`);
  }
  return index;
}

function mapDataToHeaders(headers, data) {
  const mappedData = {};
  headers.forEach((header, index) => {
    mappedData[header] = data[index] || '';
  });
  return mappedData;
}

// --- SECURITY & VALIDATION FUNCTIONS ---
function validateUserAccess() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      throw new Error('User not authenticated');
    }

    console.log(`Access by user: ${userEmail} at ${new Date().toISOString()}`);
    return userEmail;
  } catch (error) {
    Logger.log('Authentication error: ' + error.message);
    throw new Error('Authentication required');
  }
}

function checkRateLimit(userEmail, maxRequests = 10, windowMinutes = 5) {
  const now = Date.now();
  const windowMs = windowMinutes * 60 * 1000;

  if (!RATE_LIMIT_CACHE[userEmail]) {
    RATE_LIMIT_CACHE[userEmail] = [];
  }

  RATE_LIMIT_CACHE[userEmail] = RATE_LIMIT_CACHE[userEmail].filter(
    timestamp => now - timestamp < windowMs
  );

  if (RATE_LIMIT_CACHE[userEmail].length >= maxRequests) {
    throw new Error(`Rate limit exceeded. Maximum ${maxRequests} requests per ${windowMinutes} minutes.`);
  }

  RATE_LIMIT_CACHE[userEmail].push(now);
}

function sanitizeInput(input) {
  if (typeof input !== 'string') return input;

  // Don't sanitize formatted time strings (they contain colons and spaces)
  if (input.match(/^\d{1,2}:\d{2}\s?(AM|PM)$/i)) {
    return input.trim();
  }

  // Don't sanitize formatted date strings (they contain commas and spaces)
  if (input.match(/^[A-Za-z]{3}\s\d{1,2},\s\d{4}$/)) {
    return input.trim();
  }

  return input.replace(/<[^>]*>/g, '')
    .replace(/[<>]/g, '')
    .trim();
}

function validateFormData(formObject) {
  const errors = [];
  const sanitizedData = {};

  const requiredFields = ['studentName', 'staffName', 'staffEmail', 'dateOfIncident', 'descriptionOfInfraction'];

  for (const field of requiredFields) {
    const value = formObject[field];
    if (!value || typeof value !== 'string' || value.trim().length === 0) {
      errors.push(`${field} is required`);
    } else {
      sanitizedData[field] = sanitizeInput(value);

      if (FIELD_LIMITS[field] && sanitizedData[field].length > FIELD_LIMITS[field]) {
        errors.push(`${field} exceeds maximum length of ${FIELD_LIMITS[field]} characters`);
      }
    }
  }

  if (formObject.staffEmail) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(formObject.staffEmail)) {
      errors.push('Valid staff email is required');
    }
  }

  if (formObject.dateOfIncident) {
    const incidentDate = new Date(formObject.dateOfIncident);
    const today = new Date();
    const oneYearAgo = new Date(today.getFullYear() - 1, today.getMonth(), today.getDate());

    if (isNaN(incidentDate.getTime())) {
      errors.push('Valid date of incident is required');
    } else if (incidentDate > today) {
      errors.push('Date of incident cannot be in the future');
    } else if (incidentDate < oneYearAgo) {
      errors.push('Date of incident cannot be more than one year ago');
    }
  }

  Object.keys(formObject).forEach(key => {
    if (!sanitizedData[key] && typeof formObject[key] === 'string') {
      sanitizedData[key] = sanitizeInput(formObject[key]);
    }
  });

  return { errors, sanitizedData };
}

// --- PERFORMANCE OPTIMIZED DATA RETRIEVAL ---
function getStudentNamesOptimized() {
  try {
    validateUserAccess();

    const cache = CacheService.getScriptCache();
    const cachedNames = cache.get('student_names');

    if (cachedNames) {
      Logger.log('Using cached student names');
      return JSON.parse(cachedNames);
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('StudentInfo');
    if (!sheet) throw new Error('Sheet "StudentInfo" not found.');

    // FIXED: Use dynamic headers
    const headers = getSheetHeaders('StudentInfo');
    const dropDownColIndex = getColumnIndex(headers, 'DropDown');

    if (dropDownColIndex === -1) {
      throw new Error('Column "DropDown" not found in "StudentInfo" sheet.');
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const data = sheet.getRange(2, dropDownColIndex + 1, lastRow - 1, 1).getValues();
    const names = data.map(row => row[0])
      .filter(name => name && name.toString().trim())
      .sort();

    cache.put('student_names', JSON.stringify(names), 1800);
    return names;
  } catch (e) {
    Logger.log('Error in getStudentNamesOptimized: ' + e.message);
    throw e;
  }
}

function getStudentInfoByNameOptimized(lookupName) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = `student_info_${lookupName}`;
    const cachedInfo = cache.get(cacheKey);

    if (cachedInfo) {
      return JSON.parse(cachedInfo);
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('StudentInfo');
    if (!sheet) throw new Error('Sheet "StudentInfo" not found for lookup.');

    // FIXED: Use dynamic headers
    const headers = getSheetHeaders('StudentInfo');
    const lookupColIndex = getColumnIndex(headers, 'DropDown');
    const nameColIndex = getColumnIndex(headers, 'StudentName');
    const idColIndex = getColumnIndex(headers, 'StudentID');
    const gradeColIndex = getColumnIndex(headers, 'GradeLevel');
    const teamColIndex = getColumnIndex(headers, 'Team');

    const data = sheet.getRange("A:E").getValues();
    data.shift(); // Remove header row

    for (const row of data) {
      if (row[lookupColIndex] === lookupName) {
        const studentInfo = {
          studentName: row[nameColIndex] || '',
          studentID: row[idColIndex] || '',
          gradeLevel: row[gradeColIndex] || '',
          team: row[teamColIndex] || ''
        };

        cache.put(cacheKey, JSON.stringify(studentInfo), 3600);
        return studentInfo;
      }
    }

    const notFoundInfo = {
      studentName: lookupName,
      studentID: 'NOT FOUND',
      gradeLevel: '',
      team: ''
    };

    cache.put(cacheKey, JSON.stringify(notFoundInfo), 300);
    return notFoundInfo;
  } catch (e) {
    Logger.log(`Error in getStudentInfoByNameOptimized: ${e.message}`);
    return { studentName: lookupName, studentID: 'ERROR', gradeLevel: '', team: '' };
  }
}

// FIXED: Updated getSubmissionDataOptimized to use dynamic headers
function getSubmissionDataOptimized(submissionId) {
  try {
    validateUserAccess();

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = [
      { sheet: ss.getSheetByName('Major'), name: 'Major' },
      { sheet: ss.getSheetByName('Minor'), name: 'Minor' },
      { sheet: ss.getSheetByName('Positive'), name: 'Positive' } // Add this line
    ];

    for (const { sheet, name } of sheets) {
      if (!sheet) continue;

      const textFinder = sheet.createTextFinder(submissionId).matchEntireCell(true);
      const foundRange = textFinder.findNext();

      if (foundRange) {
        const rowNum = foundRange.getRow();
        const headers = getSheetHeaders(name);
        const data = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];

        const result = {
          data: data,
          headers: headers,
          sheet: name,
          submissionId: submissionId,
          retrievedAt: new Date().toISOString()
        };

        return JSON.stringify(result);
      }
    }

    throw new Error(`Submission ID '${submissionId}' was not found in any sheet.`);

  } catch (e) {
    Logger.log('Error in getSubmissionDataOptimized: ' + e.message);
    return JSON.stringify({ error: e.message });
  }
}

// --- FIXED: SECURE FORM PROCESSING WITH DYNAMIC HEADERS ---
function processFormSecure(formObject) {
  try {
    const userEmail = validateUserAccess();
    checkRateLimit(userEmail, 5, 10);

    // For Positive referrals, skip some validation
    if (formObject.referralType === 'Positive') {
      // Simplified validation for Positive referrals
      const positiveErrors = [];

      if (!formObject.studentName?.trim()) {
        positiveErrors.push('Student name is required');
      }
      if (!formObject.staffName?.trim()) {
        positiveErrors.push('Staff name is required');
      }
      if (!formObject.staffEmail?.trim()) {
        positiveErrors.push('Staff email is required');
      }
      if (!formObject.dateOfIncident) {
        positiveErrors.push('Date is required');
      }

      if (positiveErrors.length > 0) {
        throw new Error('Validation failed: ' + positiveErrors.join(', '));
      }

      // Process positive referral
      return processPositiveReferral(formObject, userEmail);
    }

    const validation = validateFormData(formObject);
    if (validation.errors.length > 0) {
      throw new Error('Validation failed: ' + validation.errors.join(', '));
    }

    const cleanData = { ...formObject, ...validation.sanitizedData };
    // Add this right after the line: const cleanData = { ...formObject, ...validation.sanitizedData };

    Logger.log('PROCESS FORM DEBUG - timeOfIncident in formObject:', formObject.timeOfIncident);
    Logger.log('PROCESS FORM DEBUG - timeOfIncident in cleanData:', cleanData.timeOfIncident);
    Logger.log('PROCESS FORM DEBUG - typeof timeOfIncident:', typeof cleanData.timeOfIncident);
    const studentDetails = getStudentInfoByNameOptimized(cleanData.studentName);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = cleanData.referralType;
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }

    // FIXED: Process infractions with proper handling to prevent duplication
    let infractions = '';
    if (cleanData.MinorInfractions) {
      infractions = Array.isArray(cleanData.MinorInfractions) ?
        cleanData.MinorInfractions.join(', ') : cleanData.MinorInfractions;
    } else if (cleanData.MajorInfractions) {
      infractions = Array.isArray(cleanData.MajorInfractions) ?
        cleanData.MajorInfractions.join(', ') : cleanData.MajorInfractions;
    }

    if (cleanData.infractionOtherText && cleanData.infractionOtherText.trim()) {
      infractions = infractions.replace(/\bOther\b/, `Other: ${sanitizeInput(cleanData.infractionOtherText)}`);
    }

    // FIXED: Process actions taken with proper handling
    let actionsTaken = '';
    if (cleanData.MinorActionsTaken) {
      actionsTaken = Array.isArray(cleanData.MinorActionsTaken) ?
        cleanData.MinorActionsTaken.join(', ') : cleanData.MinorActionsTaken;
    } else if (cleanData.MajorActionsTaken) {
      actionsTaken = Array.isArray(cleanData.MajorActionsTaken) ?
        cleanData.MajorActionsTaken.join(', ') : cleanData.MajorActionsTaken;
    }

    if (cleanData.actionsTakenOtherText && cleanData.actionsTakenOtherText.trim()) {
      actionsTaken = actionsTaken.replace(/\bOther\b/, `Other: ${sanitizeInput(cleanData.actionsTakenOtherText)}`);
    }

    // FIXED: Use dynamic headers
    const headers = getSheetHeaders(sheetName);
    const submissionId = `REF-${Date.now()}`;
    console.log(cleanData.timeOfIncident);
    console.log(typeof cleanData.timeOfIncident);
    const newRow = headers.map(header => {
      Logger.log('DEBUG - Right before switch, cleanData has: ' + JSON.stringify(cleanData));
      Logger.log('DEBUG - cleanData.timeOfIncident specifically: ' + (cleanData.timeOfIncident || 'UNDEFINED'));

      switch (header) {
        case 'StudentName': return studentDetails.studentName;
        case 'StudentID': return studentDetails.studentID;
        case 'GradeLevel': return studentDetails.gradeLevel;
        case 'Team': return studentDetails.team;
        case 'StaffName': return cleanData.staffName;
        case 'StaffEmail': return cleanData.staffEmail;
        case 'DateOfIncident': return `'${formatDateDisplay(cleanData.dateOfIncident)}`;
        case 'TimeOfIncident':
          let timeValue = cleanData.timeOfIncident ||
            cleanData.TimeOfIncident ||
            formObject.timeOfIncident ||
            formObject.TimeOfIncident;

          Logger.log('DEBUG - timeValue: ' + timeValue);

          if (timeValue) {
            // Skip formatTimeDisplay - just use the raw value
            const result = `'${timeValue}`;
            Logger.log('DEBUG - Returning raw time: ' + result);
            return result;
          }
          return '';
        case 'Location': return cleanData.Location;
        case 'DescriptionOfInfraction': return cleanData.descriptionOfInfraction;
        // FIXED: Home communication logic - only set "Yes" if checkbox is actually checked for Major referrals
        case 'HomeCommunications':
          if (cleanData.referralType === 'Major') {
            return cleanData.HomeCommunications === 'on' ? 'Yes' : 'No';
          } else {
            // For minor referrals, leave blank instead of "No"
            return '';
          }
        case 'DetailsOfCommunication': return cleanData.DetailsOfCommunication || '';
        case 'Comments': return cleanData.Comments || '';
        // FIXED: Process perceived motivation properly
        case 'PerceivedMotivation':
          return Array.isArray(cleanData.PerceivedMotivation) ?
            cleanData.PerceivedMotivation.join(', ') : (cleanData.PerceivedMotivation || '');
        case 'Infraction': return infractions;
        case 'ActionsTaken': return actionsTaken;
        case 'SubmissionNumber': return submissionId;
        case 'Timestamp': return `'${formatDateDisplay(new Date())} at ${formatTimeDisplay(new Date())}`;
        case 'Status': return 'Pending Admin Review';
        case 'DayOfWeek': return new Date(cleanData.dateOfIncident).toLocaleString('en-US', { weekday: 'long' });
        case 'SubmittedBy': return userEmail;
        case 'DocURL': return '';
        case 'RewardTripPts': return '';
        case 'DCPoints': return '';
        default: return '';
      }
    });

    sheet.appendRow(newRow);

    const cache = CacheService.getScriptCache();
    cache.remove(`student_info_${cleanData.studentName}`);

    try {
      sendNotificationEmail(submissionId, studentDetails, cleanData, sheetName, userEmail);
    } catch (emailError) {
      Logger.log('Email notification failed: ' + emailError.message);
    }

    Logger.log(`Successful submission: ${submissionId} by ${userEmail} for student ${studentDetails.studentName}`);

    return `Success! Referral ${submissionId} has been submitted and administrator has been notified.`;

  } catch (error) {
    Logger.log('Error in processFormSecure: ' + error.message);
    Logger.log('User: ' + Session.getActiveUser().getEmail());
    Logger.log('Data: ' + JSON.stringify(formObject));
    throw error;
  }
}

function sendNotificationEmail(submissionId, studentDetails, formData, referralType, submitterEmail) {
  try {
    const webAppUrl = ScriptApp.getService().getUrl();
    const adminLink = `${webAppUrl}?page=admin&id=${submissionId}`;

    const infractions = formData.MinorInfractions || formData.MajorInfractions || 'Not specified';

    const emailSubject = `🚨 New ${referralType} Referral: ${studentDetails.studentName} (${submissionId})`;

    const emailBody = `
A new ${referralType} discipline referral has been submitted.

STUDENT INFORMATION:
• Name: ${studentDetails.studentName}
• ID: ${studentDetails.studentID}
• Grade: ${studentDetails.gradeLevel}
• Team: ${studentDetails.team}

INCIDENT DETAILS:
• Date: ${formData.dateOfIncident}
• Time: ${formData.timeOfIncident || 'Not specified'}
• Location: ${formData.Location}
• Submitted by: ${formData.staffName} (${submitterEmail})

INFRACTIONS:
${infractions}

DESCRIPTION:
${formData.descriptionOfInfraction}

PRIORITY: ${referralType === 'Major' ? 'HIGH' : 'STANDARD'}

To review and process this referral, click here:
${adminLink}

Submission ID: ${submissionId}
Submitted: ${formatDateDisplay(new Date())} at ${formatTimeDisplay(new Date())}

---
This is an automated notification from the School Discipline System.
    `;

    MailApp.sendEmail({
      to: SETTINGS.ADMIN_EMAIL,
      subject: emailSubject,
      body: emailBody
    });

  } catch (error) {
    Logger.log('Email notification error: ' + error.message);
    throw error;
  }
}

// --- FIXED: PDF CREATION WITH ADMINISTRATOR'S GMAIL ---
function createEmailDraftWithPDF(referralData) {
  try {
    const userEmail = validateUserAccess();

    Logger.log(`Creating PDF for submission: ${referralData.submissionNumber}`);

    const submissionResponse = getSubmissionDataOptimized(referralData.submissionNumber);
    const submissionData = JSON.parse(submissionResponse);

    if (submissionData.error) {
      throw new Error(submissionData.error);
    }

    // FIXED: Use dynamic headers instead of hard-coded array
    const headers = submissionData.headers;
    const data = submissionData.data;
    const fullSubmissionData = mapDataToHeaders(headers, data);

    Logger.log('Full submission data extracted successfully');

    const tempDoc = DocumentApp.create(`TEMP_${referralData.submissionNumber}_${Date.now()}`);
    const body = tempDoc.getBody();

    // Style the temporary document
    const headerStyle = {};
    headerStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
    headerStyle[DocumentApp.Attribute.FONT_SIZE] = 16;
    headerStyle[DocumentApp.Attribute.BOLD] = true;
    headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#d9534f';

    const normalStyle = {};
    normalStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
    normalStyle[DocumentApp.Attribute.FONT_SIZE] = 10;

    const labelStyle = {};
    labelStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
    labelStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
    labelStyle[DocumentApp.Attribute.BOLD] = true;

    body.clear();
    const header = body.appendParagraph('DISCIPLINE REFERRAL REPORT');
    header.setAttributes(headerStyle);
    header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    body.appendParagraph('').setSpacingAfter(10);

    body.appendParagraph(`Student: ${referralData.studentName}`).setAttributes(normalStyle);
    body.appendParagraph(`Submission ID: ${referralData.submissionNumber}`).setAttributes(normalStyle);
    body.appendParagraph(`Date of Incident: ${referralData.dateOfIncident}`).setAttributes(normalStyle);
    body.appendParagraph(`Location: ${referralData.location}`).setAttributes(normalStyle);
    body.appendParagraph(`Staff Member: ${referralData.staffName}`).setAttributes(normalStyle);
    body.appendParagraph(`Type: ${referralData.referralType} Referral`).setAttributes(normalStyle);

    body.appendParagraph('').setSpacingAfter(10);

    const infractionsLabel = body.appendParagraph('INFRACTIONS:');
    infractionsLabel.setAttributes(labelStyle);
    body.appendParagraph(referralData.infraction || fullSubmissionData.Infraction || 'Not specified').setAttributes(normalStyle);

    body.appendParagraph('').setSpacingAfter(10);

    const descriptionLabel = body.appendParagraph('DESCRIPTION:');
    descriptionLabel.setAttributes(labelStyle);
    body.appendParagraph(referralData.description || fullSubmissionData.DescriptionOfInfraction || 'Not provided').setAttributes(normalStyle);

    if (fullSubmissionData.AdminConsequence) {
      body.appendParagraph('').setSpacingAfter(10);
      const adminActionsLabel = body.appendParagraph('ADMINISTRATIVE CONSEQUENCES:');
      adminActionsLabel.setAttributes(labelStyle);
      body.appendParagraph(fullSubmissionData.AdminConsequence).setAttributes(normalStyle);
    }

    if (referralData.adminComments || fullSubmissionData.AdminComments) {
      body.appendParagraph('').setSpacingAfter(10);
      const adminCommentsLabel = body.appendParagraph('ADMINISTRATIVE COMMENTS:');
      adminCommentsLabel.setAttributes(labelStyle);
      body.appendParagraph(referralData.adminComments || fullSubmissionData.AdminComments).setAttributes(normalStyle);
    }

    if (fullSubmissionData.RewardTripPts && fullSubmissionData.RewardTripPts !== '') {
      body.appendParagraph('').setSpacingAfter(10);
      const rewardLabel = body.appendParagraph('REWARD TRIP POINTS DEDUCTED:');
      rewardLabel.setAttributes(labelStyle);
      body.appendParagraph(fullSubmissionData.RewardTripPts + ' points').setAttributes(normalStyle);
    }

    if (fullSubmissionData.DCPoints && fullSubmissionData.DCPoints !== '') {
      body.appendParagraph('').setSpacingAfter(10);
      const dcLabel = body.appendParagraph('DC TRIP POINTS DEDUCTED:');
      dcLabel.setAttributes(labelStyle);
      body.appendParagraph(fullSubmissionData.DCPoints + ' points').setAttributes(normalStyle);
    }

    body.appendParagraph('').setSpacingAfter(10);
    body.appendParagraph(`Generated: ${formatDateDisplay(new Date())} at ${formatTimeDisplay(new Date())}`).setAttributes(normalStyle);

    tempDoc.saveAndClose();

    Logger.log('Temporary document created and saved');

    const tempFile = DriveApp.getFileById(tempDoc.getId());
    const pdfBlob = tempFile.getAs('application/pdf');
    pdfBlob.setName(`Discipline_Referral_${referralData.submissionNumber}.pdf`);

    Logger.log('PDF blob created');

    const subject = `Discipline Referral Report - ${referralData.studentName} (${referralData.submissionNumber})`;
    const body_text = `
Dear Colleague,

Please find attached the discipline referral report for ${referralData.studentName}.

Summary:
• Student: ${referralData.studentName}
• Submission ID: ${referralData.submissionNumber}
• Date of Incident: ${referralData.dateOfIncident}
• Type: ${referralData.referralType} Referral
• Location: ${referralData.location}

This report has been generated from the school discipline tracking system.

Best regards,
${userEmail}
    `;

    // FIXED: Create Gmail draft in the administrator's account using Gmail API
    try {
      const draft = GmailApp.createDraft(
        '', // No recipients - admin will add them
        subject,
        body_text,
        {
          attachments: [pdfBlob],
          from: userEmail // FIXED: Specify sender as current user
        }
      );

      Logger.log('Gmail draft created with attachment in administrator account');
    } catch (gmailError) {
      Logger.log('Gmail draft creation error: ' + gmailError.message);
      // Fallback: Try without the 'from' parameter
      const draft = GmailApp.createDraft(
        '',
        subject,
        body_text,
        {
          attachments: [pdfBlob]
        }
      );
      Logger.log('Gmail draft created with attachment (fallback method)');
    }

    DriveApp.getFileById(tempDoc.getId()).setTrashed(true);

    Logger.log(`Email draft created with PDF for ${referralData.submissionNumber}`);

    return 'Email draft with PDF attachment created successfully in your Gmail drafts.';

  } catch (error) {
    Logger.log('Error creating email draft with PDF: ' + error.message);
    Logger.log('Error stack: ' + error.stack);
    throw new Error(`Failed to create PDF: ${error.message}`);
  }
}

// --- GOOGLE DOC EXPORT FUNCTIONALITY ---
function createGoogleDocExport(formData, submissionData) {
  try {
    const folder = DriveApp.getFolderById(SETTINGS.DESTINATION_FOLDER_ID);

    const docTitle = `Discipline Referral - ${submissionData.StudentName} - ${submissionData.SubmissionNumber}`;

    const doc = DocumentApp.create(docTitle);
    const body = doc.getBody();

    const file = DriveApp.getFileById(doc.getId());
    file.moveTo(folder);

    const headerStyle = {};
    headerStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
    headerStyle[DocumentApp.Attribute.FONT_SIZE] = 18;
    headerStyle[DocumentApp.Attribute.BOLD] = true;
    headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#111184';

    const normalStyle = {};
    normalStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
    normalStyle[DocumentApp.Attribute.FONT_SIZE] = 11;

    const labelStyle = {};
    labelStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
    labelStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
    labelStyle[DocumentApp.Attribute.BOLD] = true;

    body.clear();

    const header = body.appendParagraph('DISCIPLINE REFERRAL REPORT');
    header.setAttributes(headerStyle);
    header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    body.appendParagraph('').setSpacingAfter(10);

    body.appendParagraph(`Student: ${submissionData.StudentName}`).setAttributes(normalStyle);
    body.appendParagraph(`Submission ID: ${submissionData.SubmissionNumber}`).setAttributes(normalStyle);
    body.appendParagraph(`Date of Incident: ${submissionData.DateOfIncident || 'Not specified'}`).setAttributes(normalStyle);
    body.appendParagraph(`Location: ${submissionData.Location}`).setAttributes(normalStyle);
    body.appendParagraph(`Staff Member: ${submissionData.StaffName}`).setAttributes(normalStyle);
    body.appendParagraph(`Type: ${formData.sheetName || 'Not specified'} Referral`).setAttributes(normalStyle);

    body.appendParagraph('').setSpacingAfter(15);

    const studentHeader = body.appendParagraph('STUDENT INFORMATION');
    studentHeader.setAttributes(labelStyle);
    studentHeader.setSpacingBefore(10);

    body.appendParagraph(`Name: ${submissionData.StudentName}`).setAttributes(normalStyle);
    body.appendParagraph(`ID: ${submissionData.StudentID}`).setAttributes(normalStyle);
    body.appendParagraph(`Grade: ${submissionData.GradeLevel}`).setAttributes(normalStyle);
    body.appendParagraph(`Team: ${submissionData.Team}`).setAttributes(normalStyle);

    body.appendParagraph('').setSpacingAfter(15);

    const incidentHeader = body.appendParagraph('INCIDENT INFORMATION');
    incidentHeader.setAttributes(labelStyle);
    incidentHeader.setSpacingBefore(10);

    body.appendParagraph(`Date of Incident: ${submissionData.DateOfIncident || 'Not specified'}`).setAttributes(normalStyle);
    body.appendParagraph(`Time of Incident: ${submissionData.TimeOfIncident || 'Not specified'}`).setAttributes(normalStyle);
    body.appendParagraph(`Day of Week: ${submissionData.DayOfWeek}`).setAttributes(normalStyle);
    body.appendParagraph(`Location: ${submissionData.Location}`).setAttributes(normalStyle);
    body.appendParagraph(`Referring Staff: ${submissionData.StaffName} (${submissionData.StaffEmail})`).setAttributes(normalStyle);

    body.appendParagraph('').setSpacingAfter(15);

    const infractionHeader = body.appendParagraph('INFRACTION DETAILS');
    infractionHeader.setAttributes(labelStyle);
    infractionHeader.setSpacingBefore(10);

    body.appendParagraph(`Type: ${formData.sheetName || 'Not specified'} Referral`).setAttributes(normalStyle);
    body.appendParagraph(`Infractions: ${submissionData.Infraction || 'Not specified'}`).setAttributes(normalStyle);

    const descriptionLabel = body.appendParagraph('Description:');
    descriptionLabel.setAttributes(labelStyle);
    const descriptionText = body.appendParagraph(submissionData.DescriptionOfInfraction || 'Not provided');
    descriptionText.setAttributes(normalStyle);
    descriptionText.setSpacingAfter(10);

    body.appendParagraph(`Previous Actions Taken: ${submissionData.ActionsTaken || 'None specified'}`).setAttributes(normalStyle);
    body.appendParagraph(`Perceived Motivation: ${submissionData.PerceivedMotivation || 'Not specified'}`).setAttributes(normalStyle);

    body.appendParagraph('').setSpacingAfter(15);

    const communicationHeader = body.appendParagraph('COMMUNICATION');
    communicationHeader.setAttributes(labelStyle);
    communicationHeader.setSpacingBefore(10);

    body.appendParagraph(`Home Communication: ${submissionData.HomeCommunications || 'No'}`).setAttributes(normalStyle);
    if (submissionData.DetailsOfCommunication) {
      body.appendParagraph(`Communication Details: ${submissionData.DetailsOfCommunication}`).setAttributes(normalStyle);
    }

    if (formData.ParentContactDate || formData.ParentContactMethod) {
      body.appendParagraph(`Admin Parent Contact Date: ${formData.ParentContactDate || 'Not specified'}`).setAttributes(normalStyle);
      body.appendParagraph(`Admin Contact Method: ${formData.ParentContactMethod || 'Not specified'}`).setAttributes(normalStyle);
      if (formData.ParentContactNotes) {
        body.appendParagraph(`Admin Contact Notes: ${formData.ParentContactNotes}`).setAttributes(normalStyle);
      }
    }

    body.appendParagraph('').setSpacingAfter(15);

    const adminHeader = body.appendParagraph('ADMINISTRATIVE ACTION');
    adminHeader.setAttributes(labelStyle);
    adminHeader.setSpacingBefore(10);

    if (formData.CodeOfConduct) {
      body.appendParagraph(`Code of Conduct Violations: ${formData.CodeOfConduct}`).setAttributes(normalStyle);
    }

    if (formData.CRDCReporting) {
      body.appendParagraph(`Ohio State Reporting Codes: ${formData.CRDCReporting}`).setAttributes(normalStyle);
    }

    if (formData.AdminConsequence) {
      body.appendParagraph(`Administrative Consequences: ${formData.AdminConsequence}`).setAttributes(normalStyle);
    }

    if (formData.AdminComments) {
      const adminCommentsLabel = body.appendParagraph('Administrator Comments:');
      adminCommentsLabel.setAttributes(labelStyle);
      const adminCommentsText = body.appendParagraph(formData.AdminComments);
      adminCommentsText.setAttributes(normalStyle);
      adminCommentsText.setSpacingAfter(10);
    }

    if (formData.RewardTripPts !== undefined && formData.RewardTripPts !== '') {
      body.appendParagraph(`Reward Trip Points Deducted: ${formData.RewardTripPts}`).setAttributes(normalStyle);
    }

    if (formData.DCPoints !== undefined && formData.DCPoints !== '') {
      body.appendParagraph(`DC Trip Points Deducted: ${formData.DCPoints}`).setAttributes(normalStyle);
    }

    body.appendParagraph('').setSpacingAfter(15);

    if (submissionData.Comments) {
      const commentsHeader = body.appendParagraph('ADDITIONAL COMMENTS');
      commentsHeader.setAttributes(labelStyle);
      commentsHeader.setSpacingBefore(10);

      const commentsText = body.appendParagraph(submissionData.Comments);
      commentsText.setAttributes(normalStyle);
    }

    body.appendParagraph('').setSpacingAfter(20);
    const footer = body.appendParagraph('--- End of Discipline Referral Report ---');
    footer.setAttributes(normalStyle);
    footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    doc.saveAndClose();

    return doc.getUrl();

  } catch (error) {
    Logger.log('Error creating Google Doc: ' + error.message);
    throw error;
  }
}

// --- FIXED: SECURE RECORD UPDATES WITH DYNAMIC HEADERS ---
function updateRecordSecure(formObject) {
  try {
    Logger.log(`Form Object ${Object.keys(formObject)}`);
    const userEmail = validateUserAccess();
    checkRateLimit(userEmail, 10, 5);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = formObject.sheetName;
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }

    const textFinder = sheet.createTextFinder(formObject.submissionNumber).matchEntireCell(true);
    const foundRange = textFinder.findNext();

    if (!foundRange) {
      throw new Error(`Submission ID ${formObject.submissionNumber} not found.`);
    }

    const rowNum = foundRange.getRow();
    // FIXED: Use dynamic headers
    const headers = getSheetHeaders(sheetName);

    const currentRowData = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
    const submissionData = mapDataToHeaders(headers, currentRowData);

    const sanitizedFormObject = {};
    Object.keys(formObject).forEach(key => {
      if (typeof formObject[key] === 'string') {
        sanitizedFormObject[key] = sanitizeInput(formObject[key]);

        if (FIELD_LIMITS[key] && sanitizedFormObject[key].length > FIELD_LIMITS[key]) {
          throw new Error(`${key} exceeds maximum length of ${FIELD_LIMITS[key]} characters`);
        }
      } else {
        sanitizedFormObject[key] = formObject[key];
      }
    });

    if (sanitizedFormObject.adminConsequenceOtherText && sanitizedFormObject.adminConsequenceOtherText.trim()) {
      sanitizedFormObject.AdminConsequence = (sanitizedFormObject.AdminConsequence || '')
        .replace('Other', `Other: ${sanitizedFormObject.adminConsequenceOtherText}`);
    }

    const updatedBy = userEmail;
    const updatedAt = formatDateDisplay(new Date()) + ' at ' + formatTimeDisplay(new Date());

    let docURL = '';
    try {
      const exportData = { ...submissionData, ...sanitizedFormObject };
      docURL = createGoogleDocExport(sanitizedFormObject, exportData);
      Logger.log(`Google Doc created: ${docURL}`);
    } catch (docError) {
      Logger.log('Google Doc creation failed: ' + docError.message);
    }

    // FIXED: Update spreadsheet using dynamic header mapping
    headers.forEach((header, index) => {
      const col = index + 1;
      let valueToSet;

      switch (header) {
        case 'CodeOfConduct':
          valueToSet = sanitizedFormObject.CodeOfConduct;
          break;
        case 'CRDCReporting':
          valueToSet = sanitizedFormObject.CRDCReporting;
          break;
        case 'AdminConsequence':
          valueToSet = sanitizedFormObject.AdminConsequence;
          break;
        case 'AdminComments':
          valueToSet = sanitizedFormObject.AdminComments;
          break;
        case 'RewardTripPts':
          valueToSet = sanitizedFormObject.RewardTripPts;
          break;
        case 'DCPoints':
          valueToSet = sanitizedFormObject.DCPoints;
          break;
        case 'Status':
          valueToSet = 'Completed';
          break;
        case 'UpdatedBy':
          valueToSet = updatedBy;
          break;
        case 'UpdatedAt':
          valueToSet = updatedAt;
          break;
        case 'ParentContactDate':
          valueToSet = sanitizedFormObject.ParentContactDate;
          break;
        case 'ParentContactMethod':
          valueToSet = sanitizedFormObject.ParentContactMethod;
          break;
        case 'ParentContactNotes':
          valueToSet = sanitizedFormObject.ParentContactNotes;
          break;
        case 'DocURL':
          valueToSet = docURL;
          break;
      }

      if (valueToSet !== undefined && valueToSet !== null && valueToSet !== '') {
        sheet.getRange(rowNum, col).setValue(valueToSet);
      }
    });

    Logger.log(`Successful update: ${sanitizedFormObject.submissionNumber} by ${userEmail}`);

    try {
      sendCompletionEmail(sanitizedFormObject, updatedBy, docURL, submissionData.StaffEmail);
    } catch (emailError) {
      Logger.log('Completion email failed: ' + emailError.message);
    }

    return `Success! Referral ${sanitizedFormObject.submissionNumber} has been updated and marked as completed. Google Doc created: ${docURL}`;

  } catch (e) {
    Logger.log('Error in updateRecordSecure: ' + e.message);
    Logger.log('User: ' + Session.getActiveUser().getEmail());
    throw e;
  }
}

function sendCompletionEmail(formData, adminEmail, docURL, staffEmail) {
  try {
    const emailSubject = `✅ Referral Completed: ${formData.submissionNumber}`;

    const emailBody = `
The discipline referral has been processed and completed.

Submission ID: ${formData.submissionNumber}
Completed by: ${adminEmail}
Completed at: ${formatDateDisplay(new Date())} at ${formatTimeDisplay(new Date())}

Administrative Comments: ${formData.AdminComments || 'None'}

Google Doc Report: ${docURL}

This referral is now closed.
    `;

    // Send to both the final email and the staff member who submitted it
    const recipients = [SETTINGS.FINAL_EMAIL];
    if (staffEmail && staffEmail.trim() && staffEmail !== SETTINGS.FINAL_EMAIL) {
      recipients.push(staffEmail);
    }

    MailApp.sendEmail({
      to: recipients.join(','),
      subject: emailSubject,
      body: emailBody
    });

    Logger.log(`Completion email sent to: ${recipients.join(', ')}`);

  } catch (error) {
    Logger.log('Completion email error: ' + error.message);
  }
}

// --- AUTO-SAVE FUNCTIONALITY ---
function autoSaveRecord(formObject) {
  try {
    const userEmail = validateUserAccess();

    const cache = CacheService.getScriptCache();
    const cacheKey = `draft_${formObject.submissionNumber}_${userEmail}`;

    cache.put(cacheKey, JSON.stringify(formObject), 3600);

    Logger.log(`Auto-saved draft for ${formObject.submissionNumber} by ${userEmail}`);

    return 'Draft saved successfully';

  } catch (error) {
    Logger.log('Auto-save error: ' + error.message);
    throw error;
  }
}

// --- ERROR HANDLING & MONITORING ---
function logSystemError(errorMessage, userEmail, functionName) {
  try {
    const errorLog = {
      timestamp: new Date().toISOString(),
      user: userEmail,
      function: functionName,
      error: errorMessage,
      userAgent: Session.getActiveUser().getEmail()
    };

    Logger.log('System Error: ' + JSON.stringify(errorLog));

  } catch (e) {
    Logger.log('Error logging failed: ' + e.message);
  }
}

// Add this new function to check if a referral is already completed
function checkReferralStatus(submissionId) {
  try {
    validateUserAccess();

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = [
      { sheet: ss.getSheetByName('Major'), name: 'Major' },
      { sheet: ss.getSheetByName('Minor'), name: 'Minor' }
    ];

    for (const { sheet, name } of sheets) {
      if (!sheet) continue;

      const textFinder = sheet.createTextFinder(submissionId).matchEntireCell(true);
      const foundRange = textFinder.findNext();

      if (foundRange) {
        const rowNum = foundRange.getRow();
        const headers = getSheetHeaders(name);
        const statusColIndex = getColumnIndex(headers, 'Status');

        if (statusColIndex !== -1) {
          const status = sheet.getRange(rowNum, statusColIndex + 1).getValue();
          return {
            found: true,
            status: status,
            isCompleted: status === 'Completed'
          };
        }
      }
    }

    return { found: false, status: null, isCompleted: false };

  } catch (error) {
    Logger.log('Error checking referral status: ' + error.message);
    return { found: false, status: null, isCompleted: false, error: error.message };
  }
}

// --- LEGACY COMPATIBILITY FUNCTIONS ---
function getStudentNames() {
  return getStudentNamesOptimized();
}

function getStudentInfoByName(lookupName) {
  return getStudentInfoByNameOptimized(lookupName);
}

function getSubmissionData(submissionId) {
  return getSubmissionDataOptimized(submissionId);
}

function processForm(formObject) {
  return processFormSecure(formObject);
}

function updateRecord(formObject) {
  return updateRecordSecure(formObject);
}

function clearCacheNames() {
  const cache = CacheService.getScriptCache();
  const cachedNames = cache.get('student_names');
  console.log(cachedNames);
  CacheService.getScriptCache().remove('student_names');
  //console.log(cachedNames);
}

function processPositiveReferral(formObject, userEmail) {
  try {
    const studentDetails = getStudentInfoByNameOptimized(formObject.studentName);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Positive');

    if (!sheet) {
      // Create the Positive sheet if it doesn't exist
      const newSheet = ss.insertSheet('Positive');
      const headers = [
        'StudentName', 'StudentID', 'GradeLevel', 'Team',
        'StaffName', 'StaffEmail', 'DateOfIncident',
        'PositiveBehavior', 'PositiveDetails', 'Comments',
        'SubmissionNumber', 'Timestamp', 'Status', 'SubmittedBy'
      ];
      newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      newSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }

    const submissionId = `POS-${Date.now()}`;

    // Get the positive behavior type (now a single value from radio button)
    const positiveBehavior = sanitizeInput(formObject.PositiveBehavior || '');
    const positiveDetails = sanitizeInput(formObject.PositiveDetails || '');

    // Get headers for dynamic mapping
    const headers = getSheetHeaders('Positive');

    const newRow = headers.map(header => {
      switch (header) {
        case 'StudentName': return studentDetails.studentName;
        case 'StudentID': return studentDetails.studentID;
        case 'GradeLevel': return studentDetails.gradeLevel;
        case 'Team': return studentDetails.team;
        case 'StaffName': return sanitizeInput(formObject.staffName);
        case 'StaffEmail': return sanitizeInput(formObject.staffEmail);
        case 'DateOfIncident': return `'${formatDateDisplay(formObject.dateOfIncident)}`;
        case 'PositiveBehavior': return positiveBehavior;
        case 'PositiveDetails': return positiveDetails;
        case 'Comments': return sanitizeInput(formObject.Comments || '');
        case 'SubmissionNumber': return submissionId;
        case 'Timestamp': return `'${formatDateDisplay(new Date())} at ${formatTimeDisplay(new Date())}`;
        case 'Status': return 'Completed';
        case 'SubmittedBy': return userEmail;
        default: return '';
      }
    });

    sheet.appendRow(newRow);

    // Clear cache
    const cache = CacheService.getScriptCache();
    cache.remove(`student_info_${formObject.studentName}`);

    // Send positive notification email
    try {
      sendPositiveNotificationEmail(submissionId, studentDetails, formObject, userEmail);
    } catch (emailError) {
      Logger.log('Positive email notification failed: ' + emailError.message);
    }

    Logger.log(`Successful positive submission: ${submissionId} by ${userEmail} for student ${studentDetails.studentName}`);

    return `Success! Positive behavior recognition ${submissionId} has been submitted. Thank you for recognizing positive student behavior!`;

  } catch (error) {
    Logger.log('Error in processPositiveReferral: ' + error.message);
    throw error;
  }
}

function sendPositiveNotificationEmail(submissionId, studentDetails, formData, submitterEmail) {
  try {
    const emailSubject = `🌟 Positive Behavior Recognition: ${studentDetails.studentName} - ${formData.PositiveBehavior}`;

    const emailBody = `
Great news! A student has been recognized for positive behavior.

STUDENT INFORMATION:
- Name: ${studentDetails.studentName}
- ID: ${studentDetails.studentID}
- Grade: ${studentDetails.gradeLevel}
- Team: ${studentDetails.team}

RECOGNITION DETAILS:
- Date: ${formData.dateOfIncident}
- Submitted by: ${formData.staffName} (${submitterEmail})
- Behavior Type: ${formData.PositiveBehavior}

${formData.PositiveDetails ? 'SPECIFIC DETAILS:\n' + formData.PositiveDetails + '\n\n' : ''}

${formData.Comments ? 'ADDITIONAL COMMENTS:\n' + formData.Comments + '\n\n' : ''}

Submission ID: ${submissionId}
Submitted: ${formatDateDisplay(new Date())} at ${formatTimeDisplay(new Date())}

---
This positive behavior recognition has been logged in the system.
Thank you for taking the time to recognize positive student behaviors!
    `;

    MailApp.sendEmail({
      to: SETTINGS.ADMIN_EMAIL,
      cc: submitterEmail,
      subject: emailSubject,
      body: emailBody
    });

  } catch (error) {
    Logger.log('Positive email notification error: ' + error.message);
    throw error;
  }
}