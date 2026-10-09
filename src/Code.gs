/***********************************************************************
 * Grading Checks Tracker
 * Google Apps Script — bound to a Google Sheet
 *
 * Supports two school-system presets out of the box:
 *   "AMERICAN" — Grades 1-12, divisions ES / MS / HS
 *   "BRITISH"  — Years 1-9, divisions Primary / Key Stage 3
 * Pick one with SCHOOL_SYSTEM below, or add your own preset.
 ***********************************************************************/

/* ---------- Configuration ---------- */
var SCHOOL_SYSTEM = "AMERICAN"; // "AMERICAN" | "BRITISH"

var SYSTEM_PRESETS = {
  AMERICAN: {
    label: "American System",
    levelLabel: "Grade Level",
    levelWord: "grade level",
    levels: ["Grade 1", "Grade 2", "Grade 3", "Grade 4", "Grade 5",
             "Grade 6", "Grade 7", "Grade 8", "Grade 9", "Grade 10", "Grade 11", "Grade 12"],
    subjects: ["English", "Social Studies", "Math", "Science", "French", "German",
               "Economics", "Sociology", "Psychology", "Business", "Political Science"],
    // [division, first level number, last level number]
    divisions: [["ES", 1, 5], ["MS", 6, 8], ["HS", 9, 12]],
    divisionHelp: "ES: Grades 1-5 | MS: Grades 6-8 | HS: Grades 9-12",
    hasVicePrincipal: false
  },
  BRITISH: {
    label: "British System",
    levelLabel: "Year Level",
    levelWord: "year level",
    levels: ["Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "Year 6",
             "Year 7", "Year 8", "Year 9"],
    subjects: ["English", "Math", "Science", "Global Perspectives", "French", "German"],
    divisions: [["Primary", 1, 6], ["Key Stage 3", 7, 9]],
    divisionHelp: "Primary: Years 1-6 | Key Stage 3: Years 7-9",
    hasVicePrincipal: true
  }
};
var SYS = SYSTEM_PRESETS[SCHOOL_SYSTEM];

var TRACKER_SHEET_NAME = "Grading Checks Tracker";
var LOG_SHEET_NAME = "Email Log";
var DASHBOARD_SHEET_NAME = "Dashboard";
var SETTINGS_SHEET_NAME = "Settings";
var SENDER_NAME = "Academic Dean's Office";

var HEADERS = [
  "Term #", "Week #", "Check Date", "Teacher's Name", "Teacher's Email",
  SYS.levelLabel, "Subject", "Assignment Type", "Assignment Title", "Grading Quality",
  "Feedback", "Grading Status", "Consistency - Fair Grading",
  "Notes", "Action Needed", "Next Review Date", "Email Status"
];
var LOG_HEADERS = HEADERS.concat(["CC Recipients", "Date Sent"]);

var COL = {
  TERM: 1, WEEK: 2, CHECK_DATE: 3, TEACHER_NAME: 4, TEACHER_EMAIL: 5,
  GRADE_LEVEL: 6, SUBJECT: 7, ASSIGN_TYPE: 8, ASSIGN_TITLE: 9, GRADING_QUALITY: 10,
  FEEDBACK: 11, GRADING_STATUS: 12, CONSISTENCY: 13,
  NOTES: 14, ACTION_NEEDED: 15, NEXT_REVIEW: 16, EMAIL_STATUS: 17
};
var COL_WIDTHS = [70, 70, 100, 150, 210, 90, 160, 140, 220, 140, 90, 115, 165, 280, 250, 120, 120];

var GRADE_LEVELS = SYS.levels;
var SUBJECTS = SYS.subjects;
var ASSIGNMENT_TYPES = ["Classwork", "Homework", "Audio", "Video Presentation", "In-Class Quiz"];
var GRADING_QUALITY_OPTIONS = ["Satisfying", "Flags Highlighted", "Dissatisfying"];
var YES_NO = ["Yes", "No"];
var GRADING_STATUS_OPTIONS = ["Complete", "Incomplete"];
var TERM_OPTIONS = ["1", "2", "3"];
var DIVISIONS = SYS.divisions.map(function (d) { return d[0]; });

var ACTION = {
  THANK_YOU: "Send Thank You Email",
  AMENDMENT: "Send Amendments Required",
  RECORRECTION: "Send Re-Correction Required",
  GRADEBOOK: "Send Gradebook Incomplete Alert"
};
var ACTION_OPTIONS = [ACTION.THANK_YOU, ACTION.AMENDMENT, ACTION.RECORRECTION, ACTION.GRADEBOOK];

/* ---------- Settings tab ---------- */
var SETTINGS_HEADER_ROW = 4;
var SETTINGS_HEADERS = ["Role", "Name", "Email", "Division", "Subject (HODs only)", "Status"];
var S_COL = { ROLE: 1, NAME: 2, EMAIL: 3, DIVISION: 4, SUBJECT: 5, STATUS: 6 };
var ROLE = {
  PRINCIPAL: "Principal",
  VICE_PRINCIPAL: "Vice Principal",
  OFFICE_ASSISTANT: "Office Assistant",
  CAMPUS_DIRECTOR: "Campus Director",
  DCI: "Director of Curriculum & Instruction",
  HOD: "Head of Department"
};
var ROLE_OPTIONS = SYS.hasVicePrincipal
  ? [ROLE.PRINCIPAL, ROLE.VICE_PRINCIPAL, ROLE.OFFICE_ASSISTANT, ROLE.CAMPUS_DIRECTOR, ROLE.DCI, ROLE.HOD]
  : [ROLE.PRINCIPAL, ROLE.OFFICE_ASSISTANT, ROLE.CAMPUS_DIRECTOR, ROLE.DCI, ROLE.HOD];
var ROLE_LIMITS = {};
ROLE_LIMITS[ROLE.PRINCIPAL] = 3;
ROLE_LIMITS[ROLE.CAMPUS_DIRECTOR] = 1;
ROLE_LIMITS[ROLE.DCI] = 3;
var SETTINGS_DIVISIONS = ["All"].concat(DIVISIONS);
var STATUS_OPTIONS = ["Active", "Inactive"];

var NAVY = "#1c3a5e";
var TEAL_LIGHT = "#d6f2ef";

/* =====================================================================
 * MENU
 * ===================================================================== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Grading Tracker")
    .addItem("Set Up Tracker (run once)", "setupTracker")
    .addItem("Set Up / Repair Settings Tab", "setupSettingsTab")
    .addSeparator()
    .addItem("Send Action Emails", "sendActionEmails")
    .addItem("Preview CC List for Selected Row", "previewCcForSelectedRow")
    .addSeparator()
    .addItem("Refresh Dashboard", "refreshDashboard")
    .addToUi();
}

/* =====================================================================
 * SETUP
 * ===================================================================== */
function setupTracker() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TRACKER_SHEET_NAME) || ss.getSheets()[0];

  if (sheet.getLastRow() > 1) {
    var resp = ui.alert("Set Up Tracker",
      "The tracker already has data. Rebuilding it will clear all rows. Continue?",
      ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) { return; }
  }

  sheet.setName(TRACKER_SHEET_NAME);
  if (sheet.getFilter()) { sheet.getFilter().remove(); }
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();

  var lastRow = Math.max(sheet.getMaxRows(), 200);
  if (sheet.getMaxRows() < lastRow) {
    sheet.insertRowsAfter(sheet.getMaxRows(), lastRow - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.length - sheet.getMaxColumns());
  }

  styleHeader_(sheet, HEADERS);
  setWidths_(sheet, COL_WIDTHS);

  var n = lastRow - 1;
  applyList_(sheet, COL.TERM, TERM_OPTIONS, n);
  applyList_(sheet, COL.GRADE_LEVEL, GRADE_LEVELS, n);
  applyList_(sheet, COL.SUBJECT, SUBJECTS, n);
  applyList_(sheet, COL.ASSIGN_TYPE, ASSIGNMENT_TYPES, n);
  applyList_(sheet, COL.GRADING_QUALITY, GRADING_QUALITY_OPTIONS, n);
  applyList_(sheet, COL.FEEDBACK, YES_NO, n);
  applyList_(sheet, COL.GRADING_STATUS, GRADING_STATUS_OPTIONS, n);
  applyList_(sheet, COL.CONSISTENCY, YES_NO, n);
  applyList_(sheet, COL.ACTION_NEEDED, ACTION_OPTIONS, n);

  var weekRule = SpreadsheetApp.newDataValidation().requireNumberBetween(1, 40)
    .setAllowInvalid(false).setHelpText("Enter a week number (1-40).").build();
  sheet.getRange(2, COL.WEEK, n, 1).setDataValidation(weekRule);

  var dateRule = SpreadsheetApp.newDataValidation().requireDate()
    .setAllowInvalid(false).setHelpText("Enter a valid date.").build();
  sheet.getRange(2, COL.CHECK_DATE, n, 1).setDataValidation(dateRule).setNumberFormat("dd-MMM-yyyy");
  sheet.getRange(2, COL.NEXT_REVIEW, n, 1).setDataValidation(dateRule).setNumberFormat("dd-MMM-yyyy");

  var emailRule = SpreadsheetApp.newDataValidation().requireTextIsEmail()
    .setAllowInvalid(false).setHelpText("Enter a valid email address.").build();
  sheet.getRange(2, COL.TEACHER_EMAIL, n, 1).setDataValidation(emailRule);

  // Free-text columns
  sheet.getRange(2, COL.ASSIGN_TITLE, n, 1).setNumberFormat("@").setWrap(true);
  sheet.getRange(2, COL.NOTES, n, 1).setNumberFormat("@").setWrap(true);

  sheet.getRange(2, 1, n, HEADERS.length).setFontSize(10).setVerticalAlignment("middle");
  sheet.getRange(2, COL.EMAIL_STATUS, n, 1).setHorizontalAlignment("center");

  applyTrackerColors_(sheet, n);
  sheet.getRange(1, 1, lastRow, HEADERS.length).createFilter();

  setupLogSheet_(ss);
  if (!ss.getSheetByName(DASHBOARD_SHEET_NAME)) { ss.insertSheet(DASHBOARD_SHEET_NAME); }
  buildSettingsSheet_(ss, false);

  ss.setActiveSheet(sheet);
  ui.alert("Tracker is ready. Fill in the Settings tab with CC recipients before sending emails.");
}

function setupSettingsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  buildSettingsSheet_(ss, false);
  ss.setActiveSheet(ss.getSheetByName(SETTINGS_SHEET_NAME));
  SpreadsheetApp.getUi().alert("Settings tab is ready. Existing contacts were kept.");
}

function setupLogSheet_(ss) {
  var logSheet = ss.getSheetByName(LOG_SHEET_NAME) || ss.insertSheet(LOG_SHEET_NAME);
  if (logSheet.getLastRow() <= 1) {
    logSheet.clear();
    logSheet.clearConditionalFormatRules();
  }
  if (logSheet.getMaxColumns() < LOG_HEADERS.length) {
    logSheet.insertColumnsAfter(logSheet.getMaxColumns(), LOG_HEADERS.length - logSheet.getMaxColumns());
  }
  styleHeader_(logSheet, LOG_HEADERS);
  setWidths_(logSheet, COL_WIDTHS);
  logSheet.setColumnWidth(LOG_HEADERS.length - 1, 280);
  logSheet.setColumnWidth(LOG_HEADERS.length, 150);
  return logSheet;
}

/* Builds the Settings tab. Keeps existing contacts unless forceReset = true. */
function buildSettingsSheet_(ss, forceReset) {
  var sh = ss.getSheetByName(SETTINGS_SHEET_NAME);
  var hasContacts = sh && sh.getLastRow() > SETTINGS_HEADER_ROW;
  if (!sh) { sh = ss.insertSheet(SETTINGS_SHEET_NAME); }

  if (!hasContacts || forceReset) {
    sh.clear();
    sh.clearConditionalFormatRules();
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();

    var rows = [];
    DIVISIONS.forEach(function (d) { rows.push([ROLE.PRINCIPAL, "", "", d, "", "Active"]); });
    if (SYS.hasVicePrincipal) {
      DIVISIONS.forEach(function (d) { rows.push([ROLE.VICE_PRINCIPAL, "", "", d, "", "Active"]); });
    }
    DIVISIONS.forEach(function (d) { rows.push([ROLE.OFFICE_ASSISTANT, "", "", d, "", "Active"]); });
    rows.push([ROLE.CAMPUS_DIRECTOR, "", "", "All", "", "Active"]);
    for (var i = 0; i < 3; i++) { rows.push([ROLE.DCI, "", "", "All", "", "Active"]); }
    SUBJECTS.forEach(function (s) { rows.push([ROLE.HOD, "", "", "All", s, "Active"]); });
    sh.getRange(SETTINGS_HEADER_ROW + 1, 1, rows.length, SETTINGS_HEADERS.length).setValues(rows);
  }

  sh.getRange(1, 1, 1, SETTINGS_HEADERS.length).merge()
    .setValue("Settings — Email CC Recipients (" + SYS.label + ")")
    .setBackground(NAVY).setFontColor("#ffffff").setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 36);

  sh.getRange(2, 1, 1, SETTINGS_HEADERS.length).merge()
    .setValue(
      "Contacts set to Active are CC'd automatically on every action email when their Division matches the teacher's " +
      SYS.levelWord + " (" + SYS.divisionHelp + ") or is set to All. " +
      "Heads of Department are CC'd only when their Subject matches the row's subject. " +
      "Limits: max 3 Principals, 1 Campus Director, 3 Directors of Curriculum & Instruction. " +
      "Each office can have more than one Office Assistant" + (SYS.hasVicePrincipal ? " or Vice Principal" : "") +
      " — just add rows at the bottom.")
    .setWrap(true).setBackground(TEAL_LIGHT).setFontSize(10).setVerticalAlignment("middle");
  sh.setRowHeight(2, 64);

  styleHeaderAt_(sh, SETTINGS_HEADER_ROW, SETTINGS_HEADERS);
  sh.setFrozenRows(SETTINGS_HEADER_ROW);
  setWidths_(sh, [260, 200, 260, 120, 190, 90]);

  var n = Math.max(sh.getMaxRows(), SETTINGS_HEADER_ROW + 60) - SETTINGS_HEADER_ROW;
  if (sh.getMaxRows() < SETTINGS_HEADER_ROW + n) {
    sh.insertRowsAfter(sh.getMaxRows(), SETTINGS_HEADER_ROW + n - sh.getMaxRows());
  }
  var start = SETTINGS_HEADER_ROW + 1;
  sh.getRange(start, S_COL.ROLE, n, 1).setDataValidation(listRule_(ROLE_OPTIONS));
  sh.getRange(start, S_COL.DIVISION, n, 1).setDataValidation(listRule_(SETTINGS_DIVISIONS));
  sh.getRange(start, S_COL.SUBJECT, n, 1).setDataValidation(listRule_(["All"].concat(SUBJECTS)));
  sh.getRange(start, S_COL.STATUS, n, 1).setDataValidation(listRule_(STATUS_OPTIONS));
  sh.getRange(start, S_COL.EMAIL, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireTextIsEmail().setAllowInvalid(false).build());
  sh.getRange(start, 1, n, SETTINGS_HEADERS.length).setFontSize(10).setVerticalAlignment("middle");
  sh.getRange(start, S_COL.DIVISION, n, 1).setHorizontalAlignment("center");
  sh.getRange(start, S_COL.STATUS, n, 1).setHorizontalAlignment("center");

  var statusRange = sh.getRange(start, S_COL.STATUS, n, 1);
  var roleRange = sh.getRange(start, S_COL.ROLE, n, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Active").setBackground("#d9ead3").setFontColor("#274e13").setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Inactive").setBackground("#eeeeee").setFontColor("#666666").setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(ROLE.PRINCIPAL).setBackground("#dbe5f1").setRanges([roleRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(ROLE.VICE_PRINCIPAL).setBackground("#e8eef7").setRanges([roleRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(ROLE.OFFICE_ASSISTANT).setBackground("#f3f3f3").setRanges([roleRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(ROLE.CAMPUS_DIRECTOR).setBackground("#fff2cc").setRanges([roleRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(ROLE.DCI).setBackground("#fce5cd").setRanges([roleRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(ROLE.HOD).setBackground(TEAL_LIGHT).setRanges([roleRange]).build()
  ]);
  return sh;
}

/* =====================================================================
 * SETTINGS HELPERS
 * ===================================================================== */
function getSettingsContacts_(ss) {
  var sh = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sh || sh.getLastRow() <= SETTINGS_HEADER_ROW) { return []; }
  var values = sh.getRange(SETTINGS_HEADER_ROW + 1, 1,
    sh.getLastRow() - SETTINGS_HEADER_ROW, SETTINGS_HEADERS.length).getValues();
  var out = [];
  values.forEach(function (r, idx) {
    var role = String(r[S_COL.ROLE - 1] || "").trim();
    var email = String(r[S_COL.EMAIL - 1] || "").trim();
    if (!role && !email) { return; }
    out.push({
      row: SETTINGS_HEADER_ROW + 1 + idx,
      role: role,
      name: String(r[S_COL.NAME - 1] || "").trim(),
      email: email,
      division: String(r[S_COL.DIVISION - 1] || "All").trim() || "All",
      subject: String(r[S_COL.SUBJECT - 1] || "").trim(),
      active: String(r[S_COL.STATUS - 1] || "").trim().toLowerCase() === "active"
    });
  });
  return out;
}

function validateSettings_(contacts) {
  var counts = {};
  contacts.forEach(function (c) {
    if (c.email) { counts[c.role] = (counts[c.role] || 0) + 1; }
  });
  var errors = [];
  Object.keys(ROLE_LIMITS).forEach(function (role) {
    if ((counts[role] || 0) > ROLE_LIMITS[role]) {
      errors.push("• " + role + ": " + counts[role] + " listed (maximum is " + ROLE_LIMITS[role] + ").");
    }
  });
  return errors;
}

function getCcRecipients_(contacts, division, subject, teacherEmail) {
  var seen = {};
  var teacher = String(teacherEmail || "").toLowerCase();
  var list = [];
  contacts.forEach(function (c) {
    if (!c.active || !c.email) { return; }
    var divOk = !c.division || c.division === "All" || c.division === division;
    if (!divOk) { return; }
    if (c.role === ROLE.HOD) {
      var subOk = !c.subject || c.subject === "All" || c.subject === subject;
      if (!subOk) { return; }
    }
    var key = c.email.toLowerCase();
    if (key === teacher || seen[key]) { return; }
    seen[key] = true;
    list.push(c.email);
  });
  return list;
}

function previewCcForSelectedRow() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  if (sheet.getName() !== TRACKER_SHEET_NAME) {
    ui.alert("Select a row in the '" + TRACKER_SHEET_NAME + "' tab first.");
    return;
  }
  var r = sheet.getActiveRange().getRow();
  if (r < 2) { ui.alert("Select a data row (row 2 or below)."); return; }
  var row = sheet.getRange(r, 1, 1, HEADERS.length).getValues()[0];
  var division = divisionForGrade_(row[COL.GRADE_LEVEL - 1]);
  var subject = String(row[COL.SUBJECT - 1] || "");
  var cc = getCcRecipients_(getSettingsContacts_(ss), division, subject, row[COL.TEACHER_EMAIL - 1]);
  ui.alert("CC preview — Row " + r,
    "Division: " + (division || "(no " + SYS.levelWord + ")") + "\nSubject: " + (subject || "(none)") + "\n\n" +
    (cc.length ? "Will be CC'd:\n" + cc.join("\n") : "No active CC recipients match this row."),
    ui.ButtonSet.OK);
}

/* =====================================================================
 * SEND EMAILS
 * ===================================================================== */
function sendActionEmails() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var sheet = ss.getSheetByName(TRACKER_SHEET_NAME) || ss.getSheets()[0];
  var logSheet = ss.getSheetByName(LOG_SHEET_NAME) || setupLogSheet_(ss);

  var contacts = getSettingsContacts_(ss);
  var errors = validateSettings_(contacts);
  if (errors.length) {
    ui.alert("Please fix the Settings tab first", errors.join("\n"), ui.ButtonSet.OK);
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert("No rows to process."); return; }
  var data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  var pending = data.filter(function (row) {
    var a = String(row[COL.ACTION_NEEDED - 1] || "").trim();
    var s = String(row[COL.EMAIL_STATUS - 1] || "").toLowerCase();
    return a && s.indexOf("sent") === -1;
  }).length;
  if (!pending) { ui.alert("There are no rows with an Action Needed waiting to be sent."); return; }

  var confirm = ui.alert("Send Action Emails",
    pending + " email(s) will be sent, with active Settings contacts CC'd. Continue?",
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) { return; }

  var sent = 0, skipped = 0, failed = 0;
  var logRows = [];
  var rowsToDelete = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var action = String(row[COL.ACTION_NEEDED - 1] || "").trim();
    var status = String(row[COL.EMAIL_STATUS - 1] || "").toLowerCase();
    var teacherEmail = String(row[COL.TEACHER_EMAIL - 1] || "").trim();
    if (!action || status.indexOf("sent") !== -1) { continue; }
    if (!teacherEmail) {
      sheet.getRange(i + 2, COL.EMAIL_STATUS).setValue("Missing email");
      skipped++;
      continue;
    }

    var division = divisionForGrade_(row[COL.GRADE_LEVEL - 1]);
    var cc = getCcRecipients_(contacts, division, String(row[COL.SUBJECT - 1] || ""), teacherEmail);

    try {
      var content = buildEmailContent_(row, action, tz);
      var msg = {
        to: teacherEmail,
        subject: content.subject,
        body: content.body,
        htmlBody: content.htmlBody,
        name: SENDER_NAME
      };
      if (cc.length) { msg.cc = cc.join(","); }
      MailApp.sendEmail(msg);
    } catch (err) {
      sheet.getRange(i + 2, COL.EMAIL_STATUS).setValue("Error: " + err.message);
      failed++;
      continue;
    }

    var logRow = row.slice();
    logRow[COL.EMAIL_STATUS - 1] = "Sent";
    logRow.push(cc.join(", "), new Date());
    logRows.push(logRow);
    rowsToDelete.push(i + 2);
    sent++;
  }

  if (logRows.length) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, LOG_HEADERS.length).setValues(logRows);
  }

  rowsToDelete.sort(function (a, b) { return b - a; });
  rowsToDelete.forEach(function (r) { sheet.deleteRow(r); });
  if (rowsToDelete.length) {
    // keep the tracker at its original size so formatting/dropdowns continue below
    sheet.insertRowsAfter(sheet.getMaxRows(), rowsToDelete.length);
  }

  refreshDashboard();

  ui.alert("Done",
    "Sent: " + sent + "\nSkipped (missing teacher email): " + skipped + "\nFailed: " + failed +
    "\n\nSent rows were moved to the Email Log.", ui.ButtonSet.OK);
}

function buildEmailContent_(row, action, tz) {
  function v(c) {
    var x = row[c - 1];
    if (x === null || x === undefined) { return ""; }
    if (x instanceof Date) { return Utilities.formatDate(x, tz, "dd MMM yyyy"); }
    return String(x).trim();
  }

  var teacherName = v(COL.TEACHER_NAME) || "Teacher";
  var subject = v(COL.SUBJECT);
  var grade = v(COL.GRADE_LEVEL);
  var week = v(COL.WEEK);
  var title = v(COL.ASSIGN_TITLE);
  var assignType = v(COL.ASSIGN_TYPE);
  var notes = v(COL.NOTES);
  var nextReview = v(COL.NEXT_REVIEW);

  var ref = [subject, grade].filter(String).join(", ") + (week ? " — Week " + week : "");
  var assignmentLabel = title ? "\"" + title + "\"" : (assignType || "this assignment");

  var details = [
    ["Term", v(COL.TERM)],
    ["Week", week],
    ["Check Date", v(COL.CHECK_DATE)],
    [SYS.levelLabel, grade],
    ["Subject", subject],
    ["Assignment Type", assignType],
    ["Assignment Title", title],
    ["Grading Quality", v(COL.GRADING_QUALITY)],
    ["Feedback Provided", v(COL.FEEDBACK)],
    ["Grading Status", v(COL.GRADING_STATUS)],
    ["Consistency / Fair Grading", v(COL.CONSISTENCY)]
  ].filter(function (d) { return d[1] !== ""; });

  var emailSubject, intro, ask, accent, urgent = false;

  if (action === ACTION.THANK_YOU) {
    emailSubject = "Thank You — Grading Check: " + ref;
    intro = "Thank you for your careful and consistent grading of " + assignmentLabel +
      ". Your recent grading check reflects the high standard we value at our school.";
    ask = "No further action is needed on your part — please keep up the excellent work.";
    accent = "#38761d";
  } else if (action === ACTION.AMENDMENT) {
    emailSubject = "Amendments Required — Grading Check: " + ref;
    intro = "Following the recent grading check of " + assignmentLabel + ", some amendments are required.";
    ask = "Please review the notes below, amend the grading of this assignment accordingly, " +
      "and reply to this email once the amendments are done.";
    accent = "#b45f06";
  } else if (action === ACTION.RECORRECTION) {
    emailSubject = "Re-Correction Required — Grading Check: " + ref;
    intro = "Following the recent grading check, " + assignmentLabel + " needs to be re-corrected.";
    ask = "Please re-correct this assignment, addressing each point in the notes below, " +
      "and reply to this email once it is ready for review.";
    accent = "#990000";
  } else if (action === ACTION.GRADEBOOK) {
    emailSubject = "URGENT: Gradebook Incomplete — " + ref;
    intro = "Our grading check found that the gradebook for " + assignmentLabel + " is incomplete.";
    ask = "Please enter the grades IMMEDIATELY for all students who have submitted this assignment, " +
      "and reply to this email once the gradebook is complete.";
    accent = "#990000";
    urgent = true;
  } else {
    emailSubject = "Grading Check Update — " + ref;
    intro = "Here is an update from your recent grading check of " + assignmentLabel + ".";
    ask = "";
    accent = NAVY;
  }

  /* ---- Plain-text version ---- */
  var body = "Dear " + teacherName + ",\n\n" + intro + "\n\n" +
    "Grading Check Details\n" +
    details.map(function (d) { return "• " + d[0] + ": " + d[1]; }).join("\n") + "\n\n" +
    (notes ? "Notes:\n" + notes + "\n\n" : "") +
    (ask ? (action === ACTION.THANK_YOU ? "" : "Action Needed: ") + ask + "\n\n" : "") +
    (nextReview && action !== ACTION.THANK_YOU ? "Next review date: " + nextReview + "\n\n" : "") +
    "Best regards,\n" + SENDER_NAME;

  /* ---- HTML version ---- */
  var rowsHtml = details.map(function (d, idx) {
    return "<tr style='background:" + (idx % 2 ? "#ffffff" : "#f5f8fb") + "'>" +
      "<td style='padding:6px 10px;font-weight:bold;color:#333;width:200px'>" + esc_(d[0]) + "</td>" +
      "<td style='padding:6px 10px;color:#333'>" + esc_(d[1]) + "</td></tr>";
  }).join("");

  var html =
    "<div style='font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;max-width:640px'>" +
    "<div style='background:" + NAVY + ";color:#fff;padding:12px 16px;font-size:16px;font-weight:bold'>" +
    (urgent ? "&#9888; " : "") + esc_(emailSubject) + "</div>" +
    "<div style='padding:16px;border:1px solid #dde3ea;border-top:none'>" +
    "<p>Dear " + esc_(teacherName) + ",</p>" +
    "<p>" + esc_(intro) + "</p>" +
    "<table style='border-collapse:collapse;width:100%;margin:12px 0;border:1px solid #dde3ea'>" + rowsHtml + "</table>" +
    (notes
      ? "<div style='margin:12px 0;padding:10px 12px;background:#fffbea;border-left:4px solid #f1c232'>" +
        "<b>Notes</b><br>" + esc_(notes).replace(/\n/g, "<br>") + "</div>"
      : "") +
    (ask
      ? "<div style='margin:12px 0;padding:10px 12px;background:#f7f7f7;border-left:4px solid " + accent + "'>" +
        (action === ACTION.THANK_YOU ? "" : "<b style='color:" + accent + "'>Action Needed:</b> ") + esc_(ask) + "</div>"
      : "") +
    (nextReview && action !== ACTION.THANK_YOU ? "<p><b>Next review date:</b> " + esc_(nextReview) + "</p>" : "") +
    "<p style='margin-top:20px'>Best regards,<br>" + esc_(SENDER_NAME) + "</p>" +
    "</div></div>";

  return { subject: emailSubject, body: body, htmlBody: html };
}

/* =====================================================================
 * AUTOMATIONS
 * ===================================================================== */
function onEdit(e) {
  if (!e || !e.range) { return; }
  var sh = e.range.getSheet();
  var name = sh.getName();

  // Mark rows as "Pending" when an action is chosen
  if (name === TRACKER_SHEET_NAME && e.range.getRow() > 1 &&
      e.range.getColumn() <= COL.ACTION_NEEDED && e.range.getLastColumn() >= COL.ACTION_NEEDED) {
    var startRow = e.range.getRow();
    var numRows = e.range.getNumRows();
    var actions = sh.getRange(startRow, COL.ACTION_NEEDED, numRows, 1).getValues();
    var statusRange = sh.getRange(startRow, COL.EMAIL_STATUS, numRows, 1);
    var statuses = statusRange.getValues();
    for (var i = 0; i < numRows; i++) {
      var a = String(actions[i][0] || "").trim();
      var s = String(statuses[i][0] || "").trim();
      if (a && !s) { statuses[i][0] = "Pending"; }
      if (!a && s === "Pending") { statuses[i][0] = ""; }
    }
    statusRange.setValues(statuses);
  }

  // Warn when Settings limits are exceeded
  if (name === SETTINGS_SHEET_NAME && e.range.getRow() > SETTINGS_HEADER_ROW) {
    var errors = validateSettings_(getSettingsContacts_(e.source));
    if (errors.length) { e.source.toast(errors.join(" "), "Settings limit exceeded", 10); }
  }
}

/* =====================================================================
 * DASHBOARD
 * ===================================================================== */
function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tracker = ss.getSheetByName(TRACKER_SHEET_NAME);
  var logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME) || ss.insertSheet(DASHBOARD_SHEET_NAME);

  var rows = [];
  [tracker, logSheet].forEach(function (s) {
    if (s && s.getLastRow() > 1) {
      rows = rows.concat(s.getRange(2, 1, s.getLastRow() - 1, HEADERS.length).getValues());
    }
  });

  var divStats = {};
  DIVISIONS.forEach(function (d) { divStats[d] = { total: 0, issues: 0 }; });
  var subStats = {};
  SUBJECTS.forEach(function (s) { subStats[s] = { total: 0, issues: 0 }; });

  var k = { checks: 0, issues: 0, incomplete: 0, noFeedback: 0, unfair: 0, dissatisfying: 0 };

  rows.forEach(function (row) {
    var quality = String(row[COL.GRADING_QUALITY - 1] || "");
    var feedback = String(row[COL.FEEDBACK - 1] || "");
    var status = String(row[COL.GRADING_STATUS - 1] || "");
    var consistency = String(row[COL.CONSISTENCY - 1] || "");
    if (!quality && !feedback && !status && !consistency) { return; }

    var isIncomplete = status === "Incomplete";
    var isNoFeedback = feedback === "No";
    var isUnfair = consistency === "No";
    var isQualityIssue = quality && quality !== "Satisfying";
    var isIssue = isIncomplete || isNoFeedback || isUnfair || isQualityIssue;

    k.checks++;
    if (isIssue) { k.issues++; }
    if (isIncomplete) { k.incomplete++; }
    if (isNoFeedback) { k.noFeedback++; }
    if (isUnfair) { k.unfair++; }
    if (quality === "Dissatisfying") { k.dissatisfying++; }

    var division = divisionForGrade_(row[COL.GRADE_LEVEL - 1]);
    var subject = row[COL.SUBJECT - 1];
    if (division && divStats[division]) {
      divStats[division].total++;
      if (isIssue) { divStats[division].issues++; }
    }
    if (subject && subStats[subject]) {
      subStats[subject].total++;
      if (isIssue) { subStats[subject].issues++; }
    }
  });

  var topDivision = topEntry_(divStats);
  var topSubject = topEntry_(subStats);

  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();

  sheet.getRange("A1:D1").merge().setValue("Grading Checks Dashboard — " + SYS.label)
    .setBackground(NAVY).setFontColor("#ffffff").setFontWeight("bold")
    .setFontSize(14).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(1, 36);

  var kpis = [
    ["Total Checks Logged", k.checks],
    ["Checks with Issues", k.issues],
    ["Gradebooks Incomplete", k.incomplete],
    ["Feedback Not Provided", k.noFeedback],
    ["Fair Grading Concerns", k.unfair],
    ["Dissatisfying Grading Quality", k.dissatisfying],
    ["Most Affected Division", topDivision ? topDivision.key + " (" + topDivision.value.issues + " issues)" : "No issues yet"],
    ["Most Affected Subject", topSubject ? topSubject.key + " (" + topSubject.value.issues + " issues)" : "No issues yet"]
  ];
  kpis.forEach(function (kp, idx) {
    var r = 3 + idx;
    sheet.getRange(r, 1).setValue(kp[0]).setFontWeight("bold").setBackground(TEAL_LIGHT);
    sheet.getRange(r, 2, 1, 2).merge().setValue(kp[1]).setBackground(TEAL_LIGHT).setHorizontalAlignment("left");
  });

  var divHeaderRow = 3 + kpis.length + 2;
  var divRows = DIVISIONS.map(function (d) {
    var st = divStats[d];
    return [d, st.total, st.issues, st.total ? st.issues / st.total : 0];
  });
  sheet.getRange(divHeaderRow, 1, 1, 4).setValues([["Division", "Total Checks", "Issues", "Issue Rate"]])
    .setBackground(NAVY).setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange(divHeaderRow + 1, 1, divRows.length, 4).setValues(divRows);
  sheet.getRange(divHeaderRow + 1, 4, divRows.length, 1).setNumberFormat("0.0%");

  var subjectRows = SUBJECTS.map(function (s) {
    var st = subStats[s];
    return [s, st.total, st.issues, st.total ? st.issues / st.total : 0];
  }).sort(function (a, b) { return b[2] - a[2]; });
  var subHeaderRow = divHeaderRow + divRows.length + 3;
  sheet.getRange(subHeaderRow, 1, 1, 4).setValues([["Subject", "Total Checks", "Issues", "Issue Rate"]])
    .setBackground(NAVY).setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange(subHeaderRow + 1, 1, subjectRows.length, 4).setValues(subjectRows);
  sheet.getRange(subHeaderRow + 1, 4, subjectRows.length, 1).setNumberFormat("0.0%");

  sheet.setColumnWidth(1, 230);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 110);

  sheet.insertChart(sheet.newChart().asColumnChart()
    .addRange(sheet.getRange(divHeaderRow + 1, 1, divRows.length, 1))
    .addRange(sheet.getRange(divHeaderRow + 1, 3, divRows.length, 1))
    .setPosition(divHeaderRow, 6, 0, 0)
    .setOption("title", "Issues by Division")
    .setOption("legend", "none")
    .setOption("colors", [NAVY])
    .setOption("vAxis.viewWindow.min", 0)
    .build());

  sheet.insertChart(sheet.newChart().asBarChart()
    .addRange(sheet.getRange(subHeaderRow + 1, 1, subjectRows.length, 1))
    .addRange(sheet.getRange(subHeaderRow + 1, 3, subjectRows.length, 1))
    .setPosition(subHeaderRow, 6, 0, 0)
    .setOption("title", "Issues by Subject")
    .setOption("legend", "none")
    .setOption("colors", ["#45818e"])
    .setOption("hAxis.viewWindow.min", 0)
    .build());
}

function topEntry_(statsObj) {
  var top = null;
  Object.keys(statsObj).forEach(function (key) {
    var val = statsObj[key];
    if (val.issues > 0 && (!top || val.issues > top.value.issues)) { top = { key: key, value: val }; }
  });
  return top;
}

/* =====================================================================
 * UTILITIES
 * ===================================================================== */
function divisionForGrade_(grade) {
  var m = String(grade || "").match(/\d+/);
  if (!m) { return ""; }
  var n = parseInt(m[0], 10);
  for (var i = 0; i < SYS.divisions.length; i++) {
    var d = SYS.divisions[i];
    if (n >= d[1] && n <= d[2]) { return d[0]; }
  }
  return "";
}

function listRule_(options) {
  return SpreadsheetApp.newDataValidation().requireValueInList(options, true).setAllowInvalid(false).build();
}

function applyList_(sheet, col, options, numRows) {
  sheet.getRange(2, col, numRows, 1).setDataValidation(listRule_(options));
}

function styleHeader_(sheet, headers) {
  styleHeaderAt_(sheet, 1, headers);
  sheet.setFrozenRows(1);
}

function styleHeaderAt_(sheet, rowNum, headers) {
  sheet.getRange(rowNum, 1, 1, headers.length).setValues([headers])
    .setBackground(NAVY).setFontColor("#ffffff").setFontWeight("bold").setFontSize(10)
    .setVerticalAlignment("middle").setHorizontalAlignment("center").setWrap(true);
  sheet.setRowHeight(rowNum, 34);
}

function setWidths_(sheet, widths) {
  for (var i = 0; i < widths.length; i++) { sheet.setColumnWidth(i + 1, widths[i]); }
}

function applyTrackerColors_(sheet, numRows) {
  function r(col) { return sheet.getRange(2, col, numRows, 1); }
  function rule(text, bg, range, font) {
    var b = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(bg).setRanges([range]);
    if (font) { b.setFontColor(font); }
    return b.build();
  }
  var q = r(COL.GRADING_QUALITY), f = r(COL.FEEDBACK), g = r(COL.GRADING_STATUS),
      c = r(COL.CONSISTENCY), a = r(COL.ACTION_NEEDED), s = r(COL.EMAIL_STATUS);
  sheet.setConditionalFormatRules([
    rule("Satisfying", "#d9ead3", q), rule("Flags Highlighted", "#fff2cc", q), rule("Dissatisfying", "#f4cccc", q),
    rule("Yes", "#d9ead3", f), rule("No", "#f4cccc", f),
    rule("Complete", "#d9ead3", g), rule("Incomplete", "#f4cccc", g),
    rule("Yes", "#d9ead3", c), rule("No", "#f4cccc", c),
    rule(ACTION.THANK_YOU, "#d9ead3", a), rule(ACTION.AMENDMENT, "#fce5cd", a),
    rule(ACTION.RECORRECTION, "#f4cccc", a), rule(ACTION.GRADEBOOK, "#ea9999", a, "#660000"),
    rule("Pending", "#fff2cc", s), rule("Sent", "#d9ead3", s)
  ]);
}

function esc_(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
