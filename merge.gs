/*** Minimal no-op helpers to avoid ReferenceErrors in merge.gs ***/
function enforceDomainAccess() {
  // No-op for now (keeps original behavior).
  // Later, we can enforce domain like:
  // const email = Session.getActiveUser().getEmail() || '';
  // if (!/@ofcs\.net$/i.test(email)) throw new Error('Unauthorized domain');
}

function logAction(tag, message) {
  // Simple logger; keeps your existing calls working.
  Logger.log('[' + tag + '] ' + (message || ''));
}

/****************************************************
 *  MERGE FUNCTION - Generates Completed Referral PDF
 *  Fix: search in "Minor" then "Major" sheets (not at spreadsheet level)
 ****************************************************/
function doMerge(subNumber, studentLastName, folderID, templateID, spreadsheetID) {
  enforceDomainAccess();
  logAction("Merge Start", "Submission#: " + subNumber + ", Student: " + studentLastName);

  try {
    if (!subNumber) throw new Error("Submission number is required.");
    if (!templateID || !folderID || !spreadsheetID) {
      throw new Error("Missing template, folder, or spreadsheet ID.");
    }

    // Open assets
    var templateFile = DriveApp.getFileById(templateID);
    var targetFolder = DriveApp.getFolderById(folderID);
    var mergedFile = templateFile.makeCopy(studentLastName + " Discipline Report - " + subNumber, targetFolder);

    var mergedDoc = DocumentApp.openById(mergedFile.getId());
    var bodyElement = mergedDoc.getBody();
    var bodyCopy = bodyElement.copy();
    bodyElement.clear();

    // --- Find the row in the proper sheet (Minor → Major) ---
    var ss = SpreadsheetApp.openById(spreadsheetID);
    var sheetNames = ["Minor", "Major"];
    var found = null;     // { sheet: Sheet, headers: array, row: array, headerMap: object }

    for (var s = 0; s < sheetNames.length && !found; s++) {
      var sh = ss.getSheetByName(sheetNames[s]);
      if (!sh) continue;
      var values = sh.getDataRange().getValues();
      if (!values || values.length < 2) continue;

      var headers = values[0];
      var headerMap = {};
      for (var i = 0; i < headers.length; i++) headerMap[String(headers[i]).trim()] = i;

      // Prefer header name; fall back to first column if missing
      var subCol = (headerMap.hasOwnProperty("SubmissionNumber")) ? headerMap["SubmissionNumber"] : 0;

      for (var r = 1; r < values.length; r++) { // skip header
        if (String(values[r][subCol]) === String(subNumber)) {
          found = { sheet: sh, headers: headers, row: values[r], headerMap: headerMap };
          break;
        }
      }
    }

    if (!found) throw new Error("No matching row for Submission#: " + subNumber);

    // --- Replace placeholders using header names ---
    var workingBody = bodyCopy.copy();
    var headersArr = found.headers;
    for (var h = 0; h < headersArr.length; h++) {
      var key = String(headersArr[h]).trim();
      if (!key) continue;
      var idx = found.headerMap[key];
      var val = (idx != null && idx < found.row.length && found.row[idx] != null) ? String(found.row[idx]) : "";
      // Replace tokens like $begin:math:display$SubmissionNumber\$end:math:display$
      // (keeping your original token pattern)
      workingBody.replaceText("\\$begin:math:display\\$" + key + "\\\\$end:math:display\\$", val);
    }

    // Replace [Today] with script TZ date
    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM-dd-yyyy");
    workingBody.replaceText("\\[Today\\]", todayStr);

    // --- Write merged content into the new doc ---
    for (var c = 0; c < workingBody.getNumChildren(); c++) {
      var child = workingBody.getChild(c).copy();
      switch (child.getType()) {
        case DocumentApp.ElementType.HORIZONTAL_RULE:
          mergedDoc.appendHorizontalRule(child);
          break;
        case DocumentApp.ElementType.INLINE_IMAGE:
          mergedDoc.appendImage(child.getBlob());
          break;
        case DocumentApp.ElementType.PARAGRAPH:
          mergedDoc.appendParagraph(child);
          break;
        case DocumentApp.ElementType.LIST_ITEM:
          mergedDoc.appendListItem(child);
          break;
        case DocumentApp.ElementType.TABLE:
          mergedDoc.appendTable(child);
          break;
        default:
          Logger.log("Unknown element type: " + child.getType());
      }
    }

    mergedDoc.saveAndClose();
    logAction("Merge Success", "File ID: " + mergedDoc.getId());
    return mergedDoc;

  } catch (err) {
    logAction("Merge Error", err.message);
    throw err;
  }
}
