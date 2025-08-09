/************************************************************
 * OFMS Discipline Project – Code.gs (revised)
 * - Routing for web app (forms vs. admin view)
 * - Safe submit with LockService and header-based writes
 * - Fast, robust lookups by SubmissionNumber (TextFinder)
 * - Admin update functions (Minor/Major/generic) by header
 * - Small utilities: header map, read/write by header, etc.
 ************************************************************/

/** ====== CONFIG ====== */
const SHEET_MINOR = "Minor";
const SHEET_MAJOR = "Major";
const SHEET_STUDENT_INFO = "StudentInfo"; // used by getStudentNames()
const ADMIN_EMAIL_DOMAIN_REGEX = /@ofcs\.net$/i; // adjust or disable in enforceDomain_()
const WEBAPP_TITLE = "OFMS Discipline Referral";

/** If you want to restrict access to your domain, set ENFORCE_DOMAIN = true */
const ENFORCE_DOMAIN = true;

/** ====== WEB APP ENTRY ====== */
function doGet(e) {
  if (ENFORCE_DOMAIN) enforceDomain_();

  const idNum = e && e.parameter && e.parameter.idNum ? String(e.parameter.idNum) : "";
  const tmpl = idNum ? HtmlService.createTemplateFromFile("admin")
                     : HtmlService.createTemplateFromFile("forms");

  // Data available to templates
  tmpl.include = include;
  tmpl.studentNames = idNum ? [] : getStudentNames_(); // Only load names on form page
  tmpl.idNum = idNum;

  const page = tmpl.evaluate()
    .setTitle(WEBAPP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return page;
}

/** HtmlService helper to inline another HTML file by name. */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Optional: restrict to a Google Workspace domain (teachers/staff). */
function enforceDomain_() {
  const email = Session.getActiveUser().getEmail() || "";
  if (!ADMIN_EMAIL_DOMAIN_REGEX.test(email)) {
    throw new Error("Unauthorized: please sign in with your school account.");
  }
}

/** ====== STUDENT LIST (used by forms.html) ====== */
/**
 * Returns an array of student display names for the form (Tom Select, etc.).
 * Adjust this to match your StudentInfo sheet schema. Currently assumes names are in column A.
 */
function getStudentNames_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_STUDENT_INFO);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, 1).getValues(); // col A
  return values.map(r => String(r[0]).trim()).filter(Boolean);
}

/** ====== SUBMIT: CREATE REFERRAL ====== */
/**
 * Submit a referral.
 * `payload` is expected to be an object from the client with keys matching your headers
 * (e.g., StudentName, StudentID, StaffName, StaffEmail, DateOfIncident, etc.) and
 * a `referralType` of "Minor" or "Major".
 */
function submitReport(payload) {
  if (ENFORCE_DOMAIN) enforceDomain_();

  if (!payload || typeof payload !== "object") {
    throw new Error("Invalid submission payload.");
  }
  const referralType = String(payload.referralType || "").trim();
  const sheetName = referralType === "Minor" ? SHEET_MINOR :
                    referralType === "Major" ? SHEET_MAJOR : null;
  if (!sheetName) {
    throw new Error('Unknown referralType. Expected "Minor" or "Major".');
  }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("Target sheet not found: " + sheetName);

  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    const { header, map } = headerIndex_(sh);
    const colCount = header.length;

    // Build a row array sized to the header length
    const row = new Array(colCount).fill("");

    // Required system fields
    const submissionNumber = Date.now(); // numeric unique id (matches your prior approach)
    if (map["submissionnumber"]) row[map["submissionnumber"] - 1] = submissionNumber;
    if (map["timestamp"])        row[map["timestamp"] - 1] = new Date();
    if (map["status"])           row[map["status"] - 1] = "Submitted";

    // Copy all provided payload keys into matching headers (by exact or normalized match)
    Object.keys(payload).forEach(k => {
      if (k === "referralType") return;
      const c = map[_norm_(k)];
      if (c) row[c - 1] = payload[k];
    });

    // Append in one write for speed
    const lastRow = sh.getLastRow();
    sh.getRange(lastRow + 1, 1, 1, colCount).setValues([row]);

    // Email the submitter (if address available) with admin link
    const to = String(payload.StaffEmail || payload.staffEmail || "").trim();
    if (to) {
      const link = buildAdminLink_(submissionNumber);
      const subject = "Discipline referral submitted: #" + submissionNumber;
      const body = "Your referral was received.\n\nAdmin link for follow-up:\n" + link;
      try {
        MailApp.sendEmail(to, subject, body, { name: WEBAPP_TITLE, noReply: true });
      } catch (err) {
        console.error("MailApp.sendEmail failed:", err);
      }
    }

    return {
      ok: true,
      referralType,
      submissionNumber,
      link: buildAdminLink_(submissionNumber)
    };
  } finally {
    lock.releaseLock();
  }
}

/** Build a stable admin link using the deployed web app URL. */
function buildAdminLink_(submissionNumber) {
  // ScriptApp.getService().getUrl() returns the webapp URL for the latest deployment (if deployed).
  // If you need a hard-coded URL, replace this with the production URL string.
  var base = "";
  try {
    base = ScriptApp.getService().getUrl() || "";
  } catch (e) {
    // Fallback to empty; caller should handle missing URL in dev
  }
  if (!base) return "DEPLOYED_WEBAPP_URL?idNum=" + encodeURIComponent(submissionNumber);
  const sep = base.includes("?") ? "&" : "?";
  return base + sep + "idNum=" + encodeURIComponent(submissionNumber);
}

/** ====== ADMIN UPDATE HELPERS ====== */

/** Normalize a key to compare with headers: "Admin Comments" -> "admincomments" */
function _norm_(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

/** Build a 1-based header index map and return { header[], map{normalizedKey -> colIndex} }. */
function headerIndex_(sh) {
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  header.forEach((h, i) => (map[_norm_(h)] = i + 1));
  return { header, map };
}

/** Use TextFinder on the correct SubmissionNumber column to get the row index. */
function findRowBySubmissionNumber_(sh, submissionNumber) {
  const { map } = headerIndex_(sh);
  const subCol = map["submissionnumber"];
  if (!subCol) throw new Error('Header "SubmissionNumber" not found on sheet: ' + sh.getName());
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const colRange = sh.getRange(2, subCol, lastRow - 1, 1); // data only
  const cell = colRange.createTextFinder(String(submissionNumber)).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : null;
}

/** Read one row into an object keyed by original header text; also in obj._n by normalized key. */
function readRowAsObject_(sh, row) {
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const obj = { _n: {} };
  header.forEach((h, i) => {
    obj[h] = vals[i];
    obj._n[_norm_(h)] = vals[i];
  });
  return obj;
}

/** Write updates by header name(s). Keys can be exact header text OR normalized. */
function writeByHeader_(sh, row, updates) {
  const { map } = headerIndex_(sh);
  const toWrite = [];
  Object.keys(updates).forEach((k) => {
    const c = map[_norm_(k)];
    if (c) toWrite.push([c, updates[k]]);
  });
  if (!toWrite.length) return;

  const cols = toWrite.map(p => p[0]);
  const minCol = Math.min.apply(null, cols);
  const maxCol = Math.max.apply(null, cols);
  const width = maxCol - minCol + 1;

  const rowVals = new Array(width).fill(null);
  toWrite.forEach(([c, v]) => (rowVals[c - minCol] = v));
  sh.getRange(row, minCol, 1, width).setValues([rowVals]);
}

/** Find a SubmissionNumber in any of the provided sheets; return {sheet,row,info} or null. */
function findSubmissionInSheets_(submissionNumber, sheetNames) {
  const ss = SpreadsheetApp.getActive();
  for (const name of sheetNames) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;
    const row = findRowBySubmissionNumber_(sh, submissionNumber);
    if (row) return { sheet: sh, row, info: readRowAsObject_(sh, row) };
  }
  return null;
}

/** ====== ADMIN UPDATE API ====== */
/**
 * Generic admin update. Writes by header name and sets Status="Completed".
 * @param {string|number} submissionNumber
 * @param {"Minor"|"Major"} referralType
 * @param {Object} updates - e.g. { AdminConsequence: "...", AdminComments: "...", CRDCReporting: "...", CodeOfConduct: "..." }
 * @return {{sheet:string,row:number,data:Object}}
 */
function updateReferralAdmin_(submissionNumber, referralType, updates) {
  if (ENFORCE_DOMAIN) enforceDomain_();

  const sheetName = referralType === "Minor" ? SHEET_MINOR :
                    referralType === "Major" ? SHEET_MAJOR : null;
  if (!sheetName) throw new Error('Unknown referralType: ' + referralType);

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet not found: ' + sheetName);

  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    const row = findRowBySubmissionNumber_(sh, submissionNumber);
    if (!row) throw new Error(`No matching row for SubmissionNumber: ${submissionNumber} in ${sheetName}`);

    // Set Status and optional background
    writeByHeader_(sh, row, { Status: "Completed" });
    const { map } = headerIndex_(sh);
    if (map.status) sh.getRange(row, map.status).setBackground("#c6efce");

    // Write admin-provided fields
    if (updates && typeof updates === "object") {
      writeByHeader_(sh, row, updates);
    }
    return { sheet: sheetName, row, data: readRowAsObject_(sh, row) };
  } finally {
    lock.releaseLock();
  }
}

/** Convenience wrappers if you call these directly from the client. */
function updateMinorAdmin_(submissionNumber, updates) {
  return updateReferralAdmin_(submissionNumber, "Minor", updates);
}
function updateMajorAdmin_(submissionNumber, updates) {
  return updateReferralAdmin_(submissionNumber, "Major", updates);
}

/** ====== UTIL / DIAGNOSTICS ====== */
/** Verify that Minor/Major headers are present (based on your provided schema). */
function verifyHeaders_() {
  const required = [
    "SubmissionNumber","Timestamp","Status","StudentName","StudentID","StaffName","StaffEmail",
    "DateOfIncident","TimeOfIncident","GradeLevel","Team","Location",
    "HomeCommunications","DetailsOfCommunication",
    "Infraction","DescriptionOfInfraction","ActionsTaken","Perceived Motivation",
    "DayOfWeek","Comments","CodeOfConduct","CRDCReporting","AdminConsequence","AdminComments"
  ].map(_norm_);

  const ss = SpreadsheetApp.getActive();
  [SHEET_MINOR, SHEET_MAJOR].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) throw new Error("Missing sheet: " + name);
    const { header } = headerIndex_(sh);
    const have = new Set(header.map(_norm_));
    const missing = required.filter(k => !have.has(k));
    if (missing.length) throw new Error(`Missing headers on ${name}: ${missing.join(", ")}`);
  });

  Logger.log("Minor/Major headers OK.");
}

/** ====== EXAMPLES (remove or keep for testing) ====== */
function demoSubmit_() {
  const res = submitReport({
    referralType: "Minor",
    StudentName: "Doe, Jane",
    StudentID: "123456",
    StaffName: "Smith, A.",
    StaffEmail: "asmith@ofcs.net",
    DateOfIncident: new Date(),
    TimeOfIncident: "09:30",
    GradeLevel: "7",
    Team: "Alpha",
    Location: "Hallway",
    Infraction: "Class disruption",
    DescriptionOfInfraction: "Talking loudly during transition",
    ActionsTaken: "Redirected twice",
    "Perceived Motivation": "Attention",
    DayOfWeek: "Wednesday",
    Comments: "Frequent issue",
    CodeOfConduct: "",
    CRDCReporting: ""
  });
  Logger.log(res);
}

function demoUpdateMinor_() {
  const sub = 1754680579517; // replace with a real SubmissionNumber
  const res = updateMinorAdmin_(sub, {
    AdminConsequence: "Lunch Detention 9/10",
    AdminComments: "Parent contacted; plan agreed.",
    CRDCReporting: "Bullying",
    CodeOfConduct: "2.1B"
  });
  Logger.log(res);
}
