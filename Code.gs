var CACHE_PROP = CacheService.getPublicCache();
var ss = SpreadsheetApp.getActiveSpreadsheet();
var SETTINGS_SHEET = "_Settings";
var CACHE_SETTINGS = false;
var SETTINGS_CACHE_TTL = 900;
var cache = JSONCacheService();
var SETTINGS = getSettings();
var minorSheet = ss.getSheetByName("Minor"); // Reference to Minor sheet
var majorSheet = ss.getSheetByName("Major"); // Reference to Major sheet
var compSheet = ss.getSheetByName("Completed"); // Still referenced, but rows won't be moved here
var studentInfoSheet = ss.getSheetByName("StudentInfo"); // Reference to StudentInfo sheet

/**
 * Handles HTTP GET requests to the web app.
 * Renders the appropriate HTML form (admin or main form)
 * and passes necessary data like student names.
 * @param {GoogleAppsScript.Events.AppsScriptHttpRequestEvent} e The event object.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The HTML output.
 */
function doGet(e) {
  var idVal = e.parameter.idNum;
  var template;

  if (idVal) {
    // If idNum parameter is present, render the admin form for a specific referral
    template = HtmlService.createTemplateFromFile("admin.html");
    var foundInfo = null;
    var targetSheetName = null; // Store sheet name here

    // Search Minor sheet first
    if (minorSheet) {
      var minorData = minorSheet.getDataRange().getValues();
      for (var i = 0; i < minorData.length; i++) {
        if (minorData[i][0] == idVal) {
          foundInfo = minorData[i];
          targetSheetName = "Minor";
          break;
        }
      }
    }

    // If not found in Minor, search Major sheet
    if (!foundInfo && majorSheet) {
      var majorData = majorSheet.getDataRange().getValues();
      for (var i = 0; i < majorData.length; i++) {
        if (majorData[i][0] == idVal) {
          foundInfo = majorData[i];
          targetSheetName = "Major";
          break;
        }
      }
    }

    if (foundInfo) {
      // Check the Status column (index 2) of the found referral
      var status = foundInfo[2]; // Assuming Status is at index 2
      if (status === "Completed") {
        template = HtmlService.createTemplateFromFile("DoneAlready.html");
      } else {
        // Add referral type at the end of the array for admin.html to use for updates
        // This will be info[26] in admin.html (since 2 new columns added before and 2 more after previous info[22])
        foundInfo.push(targetSheetName); // info[26] will be 'Minor' or 'Major'
        template.info = foundInfo;
      }
    } else {
      template = HtmlService.createTemplateFromFile("DoneAlready.html");
    }
  } else {
    // If no idNum, render the main submission form
    template = HtmlService.createTemplateFromFile('forms.html');
    // Pass student names to the forms.html template for the Tom Select dropdown (Column A only)
    template.studentNames = getStudentNames();
  }

  var html = template.evaluate();
  var output = HtmlService.createHtmlOutput(html).setTitle("PBIS Office Referral");
  return output;
}

/**
 * Fetches student names from the "StudentInfo" sheet (Column A).
 * @returns {Array<string>} An array of student names.
 */
function getStudentNames() {
  if (!studentInfoSheet) {
    Logger.log("StudentInfo sheet not found.");
    return [];
  }
  var data = studentInfoSheet.getDataRange().getValues();
  var studentList = [];
  if (data.length > 1) { // Assuming header row, so start from second row
    for (var i = 1; i < data.length; i++) {
      var name = data[i][0] ? data[i][0].toString().trim() : ''; // Column A (index 0) for student name
      if (name) {
        studentList.push(name);
      }
    }
  }
  return studentList;
}

/**
 * Fetches student's details from the "StudentInfo" sheet based on student NAME.
 * Assumes Student Name in Column A (index 0), StudentID in Column B (index 1),
 * GradeLevel in Column E (index 4), Team in Column F (index 5).
 * @param {string} studentName The student's full name to search for.
 * @returns {Object} An object with 'studentId', 'gradeLevel', and 'team', or null if not found.
 */
function getStudentInfo(studentName) {
  if (!studentInfoSheet) {
    Logger.log("StudentInfo sheet not found for student details lookup.");
    return null;
  }
  var data = studentInfoSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { // Skip header row
    // Compare student name from form with Column A of StudentInfo sheet
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === studentName.toLowerCase()) {
      return {
        studentId: data[i][1] ? data[i][1].toString().trim() : '', // Column B (index 1) for StudentID
        gradeLevel: data[i][4] ? data[i][4].toString().trim() : '', // Column E (index 4) for GradeLevel
        team: data[i][5] ? data[i][5].toString().trim() : '' // Column F (index 5) for Team
      };
    }
  }
  return null; // Student name not found
}

/**
 * Submits a new referral report (either Minor or Major) to the "Minor" or "Major" sheet.
 * @param {string} data JSON string of form data.
 * @returns {string} Success or failure message.
 */
function submitReport(data) {
  var data3 = JSON.parse(data);
  var newStuff = [];
  var subNumber = +new Date(); // Unique submission ID
  var timeStamp = new Date();
  var submitterEmail = Session.getActiveUser().getEmail(); // Get the email of the user submitting
  var targetSheet;

  if (data3.referralType === "Minor") {
    targetSheet = minorSheet;
  } else if (data3.referralType === "Major") {
    targetSheet = majorSheet;
  } else {
    return "Error: Invalid referral type provided.";
  }

  if (!targetSheet) {
    Logger.log("Target sheet not found for referral type: " + data3.referralType);
    return "Error: The '" + data3.referralType + "' sheet does not exist in the spreadsheet. Please create it.";
  }

  // Retrieve student details using the selected student name
  var studentNameFromForm = data3.studentName; // This is the plain student name from Tom Select
  var studentDetails = getStudentInfo(studentNameFromForm);
  var studentId = studentDetails ? studentDetails.studentId : '';
  var gradeLevel = studentDetails ? studentDetails.gradeLevel : '';
  var team = studentDetails ? studentDetails.team : '';

  // Determine the date of incident for Day of Week calculation
  var incidentDateStr = data3.referralType === "Minor" ? data3.minorDateOfIncident : data3.majorDateOfIncident;
  var dayOfWeek = '';
  if (incidentDateStr) {
    try {
      var dateParts = incidentDateStr.split('-'); //YYYY-MM-DD
      var incidentDate = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
      var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
      dayOfWeek = days[incidentDate.getDay()];
    } catch (e) {
      Logger.log("Error parsing date for Day of Week: " + e.message);
      dayOfWeek = 'Unknown';
    }
  }

  // Column mapping for new sheet structure
  // 0: SubmissionNumber
  newStuff.push(subNumber);
  // 1: Timestamp
  newStuff.push("'" + timeStamp.toLocaleDateString('en-US'));
  // 2: Status
  newStuff.push("Pending");
  // 3: StudentName
  newStuff.push(studentNameFromForm);
  // 4: StudentID
  newStuff.push(studentId);
  // 5: StaffName
  newStuff.push(data3.referralType === "Minor" ? data3.minorStaffName : data3.majorStaffName || '');
  // 6: StaffEmail
  newStuff.push(data3.referralType === "Minor" ? data3.minorStaffEmail : data3.majorStaffEmail || submitterEmail);
  // 7: DateOfIncident
  newStuff.push("'" + (data3.referralType === "Minor" ? data3.minorDateOfIncident : data3.majorDateOfIncident) || '');
  // 8: TimeOfIncident
  newStuff.push("'" + (data3.referralType === "Minor" ? data3.minorTimeOfIncident : data3.majorTimeOfIncident) || '');
  // 9: GradeLevel
  newStuff.push(gradeLevel); // Populated from StudentInfo
  // 10: Team
  newStuff.push(team); // Populated from StudentInfo
  // 11: Location
  newStuff.push(data3.referralType === "Minor" ? data3.minorLocation : data3.majorLocation || '');
  // 12: HomeCommunications (Only for Major)
  newStuff.push(data3.referralType === "Major" ? (Array.isArray(data3.homeCommunicationPhone) ? data3.homeCommunicationPhone.join(", ") : data3.homeCommunicationPhone || '') : '');
  // 13: DetailsOfCommunication
  newStuff.push(data3.referralType === "Major" ? data3.majorCommunicate : '');
  // 14: Infraction
  if (data3.referralType === "Minor") {
      newStuff.push(Array.isArray(data3.minorInfraction) ? data3.minorInfraction.join(", ") : '');
  } else { // Major
      newStuff.push(Array.isArray(data3.majorInfraction) ? data3.majorInfraction.join(", ") : '');
  }
  // 15: DescriptionOfInfraction
  newStuff.push(data3.referralType === "Minor" ? data3.minorDescription : data3.majorDescription || '');
  // 16: Previous Action(s) Taken
  newStuff.push(Array.isArray(data3.referralType === "Minor" ? data3.minorActionTaken : data3.majorActionTaken) ? (data3.referralType === "Minor" ? data3.minorActionTaken : data3.majorActionTaken).join(", ") : '');
  // 17: PerceivedMotivation (NEW COLUMN)
  newStuff.push(Array.isArray(data3.referralType === "Minor" ? data3.minorPerceivedMotivation : data3.majorPerceivedMotivation) ? (data3.referralType === "Minor" ? data3.minorPerceivedMotivation : data3.majorPerceivedMotivation).join(", ") : '');
  // 18: DayOfWeek (NEW COLUMN)
  newStuff.push(dayOfWeek);
  // 19: Comments (OLD 17)
  newStuff.push(data3.referralType === "Minor" ? data3.minorComments : data3.majorComments || '');
  // 20: StatusCodeOfConduct (Admin field - OLD 18)
  newStuff.push("");
  // 21: CRDCReporting (Admin field - OLD 19)
  newStuff.push("");
  // 22: AdminConsequence (Admin field - OLD 20)
  newStuff.push("");
  // 23: AdminComments (Admin field - OLD 21)
  newStuff.push("");
  // 24: PBIS Reward (NEW - admin input)
  newStuff.push("");
  // 25: Washington DC (NEW - admin input, only for 8th grade)
  newStuff.push("");


  try {
    targetSheet.appendRow(newStuff);

    // Prepare email notification subject and body
    var subject = (data3.referralType === "Minor" ? "Minor PBIS Office Referral" : "Major PBIS Office Referral") + " for " + studentNameFromForm + " #" + subNumber;
    var htmlBody = "<h2>PBIS Office Referral Submitted. </h2>";
    // Pass the referral type in the URL for doGet to know which sheet to query
    htmlBody += '<p><strong>Click <a href="' + ScriptApp.getService().getUrl() + '?idNum=' + subNumber + '">on this link</a> to see full details of referral and to add administrative feedback.</strong></p>';
    htmlBody += "<p>&nbsp;</p>";
    htmlBody += "<h4>Summary: </h4>";
    htmlBody += "<p>Referral Type: " + data3.referralType + "<br>";
    htmlBody += "Student: " + studentNameFromForm + "<br>";
    htmlBody += "Student ID: " + studentId + "<br>";
    htmlBody += "Grade Level: " + gradeLevel + "<br>";
    htmlBody += "Team: " + team + "<br>";
    htmlBody += "Date of Incident: " + (data3.referralType === "Minor" ? data3.minorDateOfIncident : data3.majorDateOfIncident)  + "<br>";
    htmlBody += "Submitted By: " + (data3.referralType === "Minor" ? data3.minorStaffName : data3.majorStaffName)  + "<br>";
    htmlBody += "Date submitted: " + timeStamp.toLocaleDateString('en-US') + "<br>";

    MailApp.sendEmail({
      to: SETTINGS.ADMIN_EMAIL,
      subject: subject,
      htmlBody: htmlBody
    });
    return "Submission successful. You may close this window now.";

  } catch(err) {
    Logger.log("Error in submitReport: " + err.message);
    return "Something went wrong during submission: " + err.message;
  }
}

/**
 * Handles administrative feedback and marks a referral as complete.
 * Now receives referralType to target the correct sheet.
 * @param {string} data JSON string of admin form data.
 * @returns {string} Success or failure message.
 */
function adminReport (data) {
  Logger.log("adminReport function called.");
  var data3 = JSON.parse(data);
  var consequences = Array.isArray(data3.adminConsequence) ? data3.adminConsequence.join(", ") : '';
  var conducts = Array.isArray(data3.codeConduct) ? data3.codeConduct.join(", ") : '';
  var codes = Array.isArray(data3.CRDCcode) ? data3.CRDCcode.join(", ") : '';
  var referralType = data3.referralTypeDisplay; // Get referral type from hidden field in admin.html
  var pbisReward = data3.pbisReward || ''; // New PBIS Reward field
  var washingtonDC = data3.washingtonDC || ''; // New Washington DC field

  try {
    // The starting index for admin fields will now be 20 (for StatusCodeOfConduct)
    // We explicitly pass the starting index (20) and let 'update' handle the rest.
    update(data3.subNum, referralType, 20, conducts, codes, consequences, data3.adminComments, pbisReward, washingtonDC);

    // studentDisplay is info[3] from admin.html
    var studentDisplay = data3.studentName || "Unknown Student";
    // studentIdDisplay is info[4] from admin.html
    var studentIdDisplay = data3.studentId || "";

    var finalDoc = doMerge(
      data3.subNum,
      studentDisplay,
      SETTINGS.DESTINATION_FOLDER_ID,
      SETTINGS.TEMPLATE_ID,
      SpreadsheetApp.getActive().getId() // ← always the file you just wrote to
    );

    var htmlBody = "<p>The following Office Referral Form was completed by an administrator:</p>";
    htmlBody += "Date submitted: " + data3.timeStamp + "<br>";
    htmlBody += "<p>Student: " + studentDisplay + " (ID: " + studentIdDisplay + ")<br>";
    htmlBody += "Date of Incident: " + data3.dateOfIncident + "<br>";
    htmlBody += "<p>Administrator Consequence:" + consequences + "</p>";
    htmlBody += "<p>Administrator Comments:" + data3.adminComments + "</p>";
    if (pbisReward) {
      htmlBody += "<p>PBIS Reward: " + pbisReward + "</p>";
    }
    if (washingtonDC) {
      htmlBody += "<p>Washington DC: " + washingtonDC + "</p>";
    }


    MailApp.sendEmail({
      to: data3.staffEmail + ", " + SETTINGS.FINAL_EMAIL,
      subject: "Completed PBIS Office Referral for "+ studentDisplay + " #" + data3.subNum,
      htmlBody: htmlBody,
      attachments: [finalDoc.getAs(MimeType.PDF)]
    });

    return "Your submission is complete. You may close this window now.";
  } catch (err) {
    Logger.log("Error in adminReport: " + err.message);
    return "Something went wrong during admin submission: " + err.message;
  }
}

/**
 * Updates a row in the Minor or Major sheet with admin feedback and changes status to "Completed".
 * @param {number} num Submission ID.
 * @param {string} referralType The type of referral ("Minor" or "Major").
 * @param {number} startAdminColIndex Index of the first admin-editable column (StatusCodeOfConduct - 20).
 * @param {string} conducts Code of Conduct value.
 * @param {string} CRDCcodes CRDC Codes value.
 * @param {string} adminConsequence Administrative Consequence value.
 * @param {string} adminComments Administrative Comments value.
 * @param {string} pbisReward PBIS Reward value.
 * @param {string} washingtonDC Washington DC value (optional).
 */
function update (num, referralType, startAdminColIndex, conducts, CRDCcodes, adminConsequence, adminComments, pbisReward, washingtonDC) {
  var targetSheet = (referralType === "Minor") ? minorSheet : majorSheet;

  if (!targetSheet) {
    Logger.log("Target sheet not found for update: " + referralType);
    throw new Error("Target sheet for update not found.");
  }

  var data = targetSheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == num) {
      // Status column (index 2, 3rd column)
      targetSheet.getRange(i+1, 3).setValue("Completed").setBackground("#00FF00");

      // Admin data starts from startAdminColIndex (which is array index 20 for StatusCodeOfConduct)
      // Spreadsheet columns are 1-indexed for getRange, so add 1 to array index
      targetSheet.getRange(i+1, startAdminColIndex + 1).setValue(conducts); // Column U (index 20 in 0-indexed, 21 in 1-indexed)
      targetSheet.getRange(i+1, startAdminColIndex + 2).setValue(CRDCcodes); // Column V (index 21 in 0-indexed, 22 in 1-indexed)
      targetSheet.getRange(i+1, startAdminColIndex + 3).setValue(adminConsequence); // Column W (index 22 in 0-indexed, 23 in 1-indexed)
      targetSheet.getRange(i+1, startAdminColIndex + 4).setValue(adminComments); // Column X (index 23 in 0-indexed, 24 in 1-indexed)
      targetSheet.getRange(i+1, startAdminColIndex + 5).setValue(pbisReward); // Column Y (index 24 in 0-indexed, 25 in 1-indexed)
      targetSheet.getRange(i+1, startAdminColIndex + 6).setValue(washingtonDC); // Column Z (index 25 in 0-indexed, 26 in 1-indexed)
      break;
    }
  }
}

/**
 * Retrieves settings from the _Settings sheet.
 * @returns {Object} An object containing settings key-value pairs.
 */
function getSettings() {
  if(CACHE_SETTINGS) {
    var settings = cache.get("_settings");
  }

  if(settings == undefined) {
    var sheet = ss.getSheetByName(SETTINGS_SHEET);
    if (!sheet) {
      Logger.log("Settings sheet not found: " + SETTINGS_SHEET);
      return {}; // Return empty object if sheet not found
    }
    var values = sheet.getDataRange().getValues();

    var settings = {};
    for (var i = 1; i < values.length; i++) { // Start from 1 to skip header row
      var row = values[i];
      if (row[0]) { // Ensure setting name exists
        settings[row[0]] = row[1];
      }
    }

    cache.put("_settings", settings, SETTINGS_CACHE_TTL);
  }
  return settings;
}

/**
 * Simple JSON cache service.
 * @returns {Object} Cache service with get and put methods.
 */
function JSONCacheService() {
  var _cache = CacheService.getPublicCache();
  var _key_prefix = "_json#";

  var get = function(k) {
    var payload = _cache.get(_key_prefix+k);
    if(payload !== null) { // Check for null explicitly, not undefined
      return JSON.parse(payload);
    }
    return undefined; // Return undefined if not found
  }

  var put = function(k, d, t) {
    _cache.put(_key_prefix+k, JSON.stringify(d), t);
  }

  return {
    'get': get,
    'put': put
  }
}

// Utility function to test settings retrieval
function testStuff() {
  Logger.log(SETTINGS.ADMIN_EMAIL);
}

// Utility function to test merge (requires specific values for testing)
function testMerge() {
  // Replace with actual values for testing
  var dummySubNumber = 1234567890;
  // studentDisplay will be the full student name (from info[3] now)
  var dummyStudentName = "John Doe";
  // Ensure SETTINGS contains these IDs or hardcode for testing
  // var dummyFolderID = SETTINGS.DESTINATION_FOLDER_ID;
  // var dummyTemplateID = SETTINGS.TEMPLATE_ID;
  // var dummySpreadsheetID = ss.getId();

  // Example: If SETTINGS are not configured or for isolated testing
  var dummyFolderID = "YOUR_DESTINATION_FOLDER_ID"; // Replace with a real folder ID for testing
  var dummyTemplateID = "YOUR_TEMPLATE_DOC_ID"; // Replace with a real template Doc ID for testing
  var dummySpreadsheetID = ss.getId(); // Use the active spreadsheet ID

  try {
    // The doMerge function is in merge.gs, which is implicitly available if merge.gs exists.
    // Assuming doMerge is accessible, you'd call it like this:
    // var doc = doMerge(dummySubNumber, dummyStudentName, dummyFolderID, dummyTemplateID, dummySpreadsheetID);
    // Logger.log("Merge test successful. Document ID: " + doc.getId());
    Logger.log("doMerge test not run directly. Please ensure merge.gs is in your Apps Script project.");
  } catch (e) {
    Logger.log("Merge test failed: " + e.message);
  }
}
