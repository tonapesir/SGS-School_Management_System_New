var MAIN_HEADERS = [
  "Timestamp","Acad Year","Reg No","Student ID","Book No","Aadhar",
  "Class","Division","PEN","Roll No","Full Name",
  "Mother Name","Gender","Religion","Caste","Sub-caste","DOB","DOB Words",
  "Nationality","Mother Tongue","Birth Place",
  "Previous School","Admission Date","Admission Class",
  "Contact","Address","Photo URL","WhatsApp Mobile No","Alternate Mobile No",
  "Category","Minority"
];
var LC_HEADERS = [
  "Timestamp","Serial No","Student ID","Full Name","Gender",
  "DOB","DOB Words","Class","Religion","Caste","Nationality","Mother Tongue",
  "Admission Date","Admission Class","Leave Date","Leave Class","Class Start Date",
  "LC No","LC Date","Conduct","Progress","Stxt75","Stxt76","Remarks",
  "Medium","Next School","LC Count","Fee Status","Stxt60","Stxt61",
  "Stxt62","Stxt63","Stxt67","Stxt68","Stxt69","Stxt70","Nationality Type"
];
var BF_HEADERS = [
  "Timestamp","Serial No","Student ID","Full Name",
  "DOB","DOB Words","Class","Tukdi","Religion","Caste","Nationality",
  "Acad Year","BF Date","Purpose","Admission Date","Admission Class",
  "Remarks","Address","Contact"
];
var AT_HEADERS = [
  "Timestamp","Serial No","Student ID","Full Name","Gender",
  "DOB","Class","Tukdi","Acad Year","AT Date","Purpose",
  "Total Days","Present Days","Percentage","From Date","To Date","Remarks"
];
var PHOTO_FOLDER_ID = "1urIudSUS7U0ClMTjvp__uZ0ZgTQ7GrbU";
var PHOTO_COL = 27;   // V19.45: Students मधील "Photo URL" स्तंभ (MAIN_HEADERS नुसार 27/AA) — लिहिणे व वाचणे दोन्ही येथेच

// ===== CLASS TEACHER MODULE (V19.32) — sheet headers =====
var DAILY_ATT_HEADERS = ["Timestamp","Date","Class","Division","RegNo","StudentId","FullName","Reason","MarkedBy"];
var NOTICE_HEADERS   = ["Timestamp","Date","Title","Message","PostedBy","TargetClass"];
var DIARY_HEADERS    = ["Timestamp","Date","RegNo","StudentId","FullName","Class","Division","Type","Remark","EnteredBy"];
var TRANSPORT_HEADERS= ["Timestamp","RegNo","StudentId","FullName","Class","Division","NativeVillage","TravelMode","TransportContact","UpdatedBy"];
// ===== FEES (V19.33) — प्रत्येक जमा installment ची स्वतंत्र row, पण त्या क्षणापर्यंतची cumulative FeePaid/PendingFee सोबत =====
var FEES_HEADERS     = ["Timestamp","RegNo","StudentId","FullName","Class","Division","AcYear","TotalFee","FeePaid","PendingFee","UpdatedBy"];
var TOTAL_FEE_PER_STUDENT_DEFAULT = 1000; // सुरुवातीचे Default — Super/Master ने Maintenance मधून बदल न केल्यास हेच वापरले जाईल
// चालू शाळेची एकूण फी — Super Master/Master ने "मेंटेनन्स → शाळेची एकूण फी ठरवा" मधून ठरवलेली, PropertiesService मध्ये जतन
function getTotalFeePerStudent() {
  try {
    var v = PropertiesService.getScriptProperties().getProperty('TOTAL_FEE_PER_STUDENT');
    var n = parseFloat(v);
    return (v !== null && !isNaN(n) && n > 0) ? n : TOTAL_FEE_PER_STUDENT_DEFAULT;
  } catch (e) {
    return TOTAL_FEE_PER_STUDENT_DEFAULT;
  }
}
var CONTACTS_HEADERS = ["Timestamp","RegNo","StudentId","FullName","Class","Division","WhatsAppMobile","OtherMobile","UpdatedBy"];
var CATEGORY_LIST_SRV = ["SC","ST","VJA","NT B","NT C","NT D","SEBC","SBC","OBC","Gen"];

function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    trimNewSheet_(sh, headers.length);
    sh.appendRow(headers);
  } else if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
  } else {
    var existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
    for (var i = 0; i < headers.length; i++) {
      if (!existing[i]) sh.getRange(1, i + 1).setValue(headers[i]);
    }
  }
  return sh;
}

function findRowByKey(sh, col, key) {
  key = (key || "").toString().trim().toLowerCase();
  if (!key || !sh || sh.getLastRow() < 2) return 0;
  var vals = sh.getRange(2, col, sh.getLastRow()-1, 1).getValues();
  for (var i=0; i<vals.length; i++) {
    if ((vals[i][0] || "").toString().trim().toLowerCase() === key) return i + 2;
  }
  return 0;
}

function writeRow(sh, rowNum, row) {
  if (rowNum) {
    sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
    return {rowIndex: rowNum, mode: "updated"};
  }
  sh.appendRow(row);
  return {rowIndex: sh.getLastRow(), mode: "created"};
}

function doPostCore_(e) {
  var p = e.parameter || {};
  var authErrP = applyAuth_(p);
  if (authErrP) {
    return ContentService.createTextOutput(JSON.stringify(authErrP)).setMimeType(ContentService.MimeType.JSON);
  }
  if (p.action === "uploadPhoto") {
    var photoResult = handlePhotoUpload(p);
    return ContentService.createTextOutput(JSON.stringify(photoResult))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return handleAction(p, ""); // "" = JSON response — serial परत frontend ला येतो
}

function doGetCore_(e) {
  var p = e.parameter || {};
  var cb = p.callback || "";
  if (!p.action && (p.regNo || p.studentId || p.firstName)) p.action = "upsert";
  var authErr = applyAuth_(p);
  if (authErr) return wrap(cb, authErr);
  if (p.action === "login") {
    return doLogin(p, cb);
  }
  if (p.action === "verify") {
    return doVerifyPage(p);
  }
  if (p.action === "ping") {
    try {
      var ssName = SpreadsheetApp.getActiveSpreadsheet().getName();
      return wrap(cb, {status:"ok", message:"Connected! Sheet: " + ssName});
    } catch(ex) {
      return wrap(cb, {status:"error", message:"Script is deployed but NOT BOUND to a Sheet! Re-deploy from inside the Sheet."});
    }
  }
  if (p.action === "getAll") {
    return doGetAllAction(p, cb);
  }
  if (p.action === "testPhotoFolder") {
    return wrap(cb, testPhotoFolderAccess());
  }
  if (p.action === "photoChunkStart") {
    return wrap(cb, startPhotoChunkUpload(p));
  }
  if (p.action === "photoChunk") {
    return wrap(cb, savePhotoChunk(p));
  }
  if (p.action === "photoChunkFinish") {
    return wrap(cb, finishPhotoChunkUpload(p));
  }
  if (p.action === "getClassTeachers") {
    return doGetClassTeachers(p, cb);
  }
  if (p.action === "getPhotoData") {
    return doGetPhotoData(p, cb);
  }
  if (p.action === "getPhotoUrl" && p.regNo) {
    return wrap(cb, getPhotoUrlByRegNo(p.regNo));
  }
  if (p.action === "search" && p.q) {
    return doSearchAction(p, cb);
  }
  if (p.action === "getDashboardStats") {
    return doGetDashboardStats(p, cb);
  }
  if (p.action === "getStudentHistory") {
    return doGetStudentHistory(p, cb);
  }
  if (p.action === "getAllCertificates") {
    return doGetAllCertificates(p, cb);
  }
  if (p.action === "getUsers") {
    return doGetUsers(p, cb);
  }
  if (p.action === "saveUser") {
    return doSaveUser(p, cb);
  }
  if (p.action === "deleteUser") {
    return doDeleteUser(p, cb);
  }
  if (p.action === "changePassword") {
    return doChangePassword(p, cb);
  }
  if (p.action === "getAnalyticsData") {
    return doGetAnalyticsData(p, cb);
  }
  if (p.action === "verify") {
    return doVerifyPage(p);
  }
  // ===== CLASS TEACHER MODULE (V19.32) =====
  if (p.action === "getClassStudents") {
    return doGetClassStudents(p, cb);
  }
  if (p.action === "getClassList") {
    return doGetClassList(p, cb);
  }
  if (p.action === "getTodayBirthdays") {
    return doGetTodayBirthdays(p, cb);
  }
  if (p.action === "getAttendance") {
    return doGetAttendance(p, cb);
  }
  if (p.action === "getNotices") {
    return doGetNotices(p, cb);
  }
  if (p.action === "getDiary") {
    return doGetDiary(p, cb);
  }
  if (p.action === "getStudentYearHistory") {   // V19.56 — एका विद्यार्थ्याचा वर्षागणिक इतिहास (फक्त वाचन)
    return doGetStudentYearHistory(p, cb);
  }
  if (p.action === "getArchiveYears") {     // V19.55 — Archive स्क्रीन (फक्त super/master, Token सह)
    return doGetArchiveYears(p, cb);
  }
  if (p.action === "getArchiveReport") {    // V19.55
    return doGetArchiveReport(p, cb);
  }
  if (p.action === "getScholarship") {      // V19.47
    return doGetScholarship(p, cb);
  }
  if (p.action === "deleteScholarship") {   // V19.47 — फक्त super/master; आधी Deleted_Scholarship मध्ये Backup
    return doDeleteScholarship(p, cb);
  }
  if (p.action === "getTransport") {
    return doGetTransport(p, cb);
  }
  if (p.action === "getFees") {
    return doGetFees(p, cb);
  }
  if (p.action === "getTeacherDashboard") {
    return doGetTeacherDashboard(p, cb);
  }
  if (p.action === "getMyLog") {
    return doGetMyLog(p, cb);
  }
  if (p.action === "getStudentContacts") {
    return doGetStudentContacts(p, cb);
  }
  if (p.action === "getStatsReport") {
    return doGetStatsReport(p, cb);
  }
  if (p.action === "deleteStudent") {
    return doDeleteStudent(p, cb);
  }
  // ===== V19.35 — नवीन Admin सुविधा =====
  if (p.action === "getPendingFeesAll") {
    return doGetPendingFeesAll(p, cb);
  }
  if (p.action === "previewPromotion") {
    return doPreviewPromotion(p, cb);
  }
  if (p.action === "getAttendancePending") {
    return doGetAttendancePending(p, cb);
  }
  if (p.action === "getBackupStatus") {
    return doGetBackupStatus(p, cb);
  }
  if (p.action === "getAttendanceAnalytics") {
    return doGetAttendanceAnalytics(p, cb);
  }
  if (p.action === "getYearStatus") {
    return doGetYearStatus(p, cb);
  }
  if (p.action === "getAttendanceOverview") {
    return doGetAttendanceOverview(p, cb);
  }
  if (p.action === "getSheetSizes") {
    return doGetSheetSizes(p, cb);
  }
  if (p.action === "getEarlyLeave") {
    return doGetEarlyLeave(p, cb);
  }
  if (p.action === "getEarlyLeavePeon") {
    return doGetEarlyLeavePeon(p, cb);
  }
  if (p.action === "getSecurityStatus") {
    return doGetSecurityStatus(p, cb);
  }
  if (p.action === "migrateLogs") {   // V19.48 — जुनी Log/Log_Archive → Log_YYYY_MM (फक्त Super, एकदाच)
    return doMigrateLogs(p, cb);
  }
  if (p.action === "photoRecall") {   // V19.46 — Photo link recall (फक्त Super)
    return doPhotoRecall(p, cb);
  }
  if (p.action === "promoteStudents" || p.action === "undoPromotion" ||
      p.action === "runBackup" || p.action === "setBackupSchedule" ||
      p.action === "archivePassedOut" || p.action === "restorePassedOut" ||
      p.action === "renumberRolls" || p.action === "archiveLogNow" ||
      p.action === "saveEarlyLeave" || p.action === "markEarlyLeaveOut" || p.action === "shrinkSheets" || p.action === "rebuildFeesSummary" ||
      p.action === "startNewYear" || p.action === "archiveYear") {
    return handleAction(p, cb);
  }
  var a = p.action || "";
  if (a.indexOf("save") === 0 || a.indexOf("upsert") === 0 ||
      a.indexOf("update") === 0) {
    return handleAction(p, cb);
  }
  return wrap(cb, {status:"ok", message:"Running!"});
}

function getNextSerial(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 1;
  var vals = sh.getRange(2, 2, lastRow - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < vals.length; i++) {
    var n = parseInt(vals[i][0], 10);
    if (!isNaN(n) && n > max) max = n;
  }
  return max + 1;
}

function handleAction(d, cb) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return wrap(cb, {status:"error", message:"Script NOT bound to a Sheet. Open Sheet → Extensions → Apps Script."});
    var action = d.action || "upsert";
    var result = {rowIndex:0, mode:"created"};

    // ===== Offline "Pending Saves" queue साठी — तोच reqId पुन्हा आल्यास (retry) दुबार नोंद होऊ नये (V19.35) =====
    var reqKey = d.reqId ? ("req_" + d.reqId) : "";
    if (reqKey) {
      try {
        var prevResp = CacheService.getScriptCache().get(reqKey);
        if (prevResp) { var pr = JSON.parse(prevResp); pr.duplicate = true; return wrap(cb, pr); }
      } catch (eReq) {}
    }

    if (action === "save" || action === "update" || action === "upsert") {
      var sh = getOrCreateSheet("Students", MAIN_HEADERS);
      var rowToWrite = 0;
      if (action === "update" && d.rowIndex) rowToWrite = parseInt(d.rowIndex,10);
      if (action === "upsert") rowToWrite = findRowByKey(sh, 3, d.regNo) || findRowByKey(sh, 4, d.studentId);
      // "संपर्क / Mobile" फील्ड फॉर्म मधून काढले आहे — जुनी नोंद असल्यास ती जागच्या जागी सुरक्षित ठेवा
      var contactVal = d.contact;
      if (!contactVal && rowToWrite) {
        contactVal = sh.getRange(rowToWrite, 25).getValue() || "";
      }
      // Category/Minority फॉर्म मधून रिकामे आल्यास (जुनी नोंद edit करताना चुकून रिकामे राहिल्यास) आधीचीच माहिती सुरक्षित ठेवा
      var categoryVal = d.category;
      var minorityVal = d.minority;
      if (!categoryVal && rowToWrite) categoryVal = sh.getRange(rowToWrite, 30).getValue() || "";
      if (!minorityVal && rowToWrite) minorityVal = sh.getRange(rowToWrite, 31).getValue() || "";
      var row = [new Date().toLocaleString("en-IN"),
        d.acYear,d.regNo,d.studentId,d.bookNo,d.aadhar,d.iyatta,d.tukdi,
        d.pen,d.rollNo,d.firstName,d.motherName,d.gender,
        d.religion,d.caste,d.subcaste,d.dob,d.dobWords,d.nationality,d.motherTongue,
        d.birthVillage,d.prevSchool,
        d.admissionDate,d.admissionClass,contactVal,d.address,d.photoUrl,
        d.whatsappMobile||"", d.alternateMobile||"",
        categoryVal||"", minorityVal||""];
      result = writeRow(sh, rowToWrite, row);
    } else if (action === "save_lc" || action === "update_lc" || action === "upsert_lc") {
      var sh = getOrCreateSheet("LC", LC_HEADERS);
      // ✅ सदैव नवीन row — प्रत्येक LC ला स्वतंत्र serial नंबर
      var rowToWrite = 0;
      var lcSerial = getNextSerial(sh);
      var row = [new Date().toLocaleString("en-IN"),
        lcSerial,d.stxt1,d.firstName,d.gender,d.dob,d.dobWords,
        d.iyatta,d.religion,d.caste,d.nationality,d.motherTongue,
        d.admissionDate,d.admissionClass,d.lcLeaveDate,d.lcLeaveClass,d.classStartDate,
        d.lcNo,d.lcDate,d.conduct,d.progress,d.stxt75,d.stxt76,d.remarks,
        d.medium,d.nextSchool,d.lcCount,d.feeStatus,
        d.stxt60,d.stxt61,d.stxt62,d.stxt63,d.stxt67,d.stxt68,d.stxt69,d.stxt70,d.nationalityType];
      result = writeRow(sh, rowToWrite, row);
      result.serial = lcSerial;
    } else if (action === "save_bf" || action === "upsert_bf") {
      var sh = getOrCreateSheet("Bonafide", BF_HEADERS);
      // ✅ सदैव नवीन row — प्रत्येक Bonafide ला स्वतंत्र serial नंबर
      var rowToWrite = 0;
      var bfSerial = getNextSerial(sh);
      var row = [new Date().toLocaleString("en-IN"),
        bfSerial,d.stxt1,d.firstName,d.dob,d.dobWords,
        d.iyatta,d.tukdi,d.religion,d.caste,d.nationality,
        d.acYear,d.bfDate,d.purpose,d.admissionDate,d.admissionClass,
        d.remarks,d.address,d.contact];
      result = writeRow(sh, rowToWrite, row);
      result.serial = bfSerial;
    } else if (action === "save_at" || action === "upsert_at") {
      var sh = getOrCreateSheet("Attendance", AT_HEADERS);
      // ✅ सदैव नवीन row — प्रत्येक Attendance certificate ला स्वतंत्र serial नंबर
      var rowToWrite = 0;
      var atSerial = getNextSerial(sh);
      var row = [new Date().toLocaleString("en-IN"),
        atSerial,d.stxt1,d.firstName,d.gender,d.dob,
        d.iyatta,d.tukdi,d.acYear,d.atDate,d.purpose,
        d.totalDays,d.presentDays,d.pct,d.fromDate,d.toDate,d.remarks];
      result = writeRow(sh, rowToWrite, row);
      result.serial = atSerial;
    } else if (action === "saveAttendance") {
      result = doSaveAttendance(d);
    } else if (action === "saveNotice") {
      result = doSaveNotice(d);
    } else if (action === "saveDiary") {
      result = doSaveDiary(d);
    } else if (action === "saveScholarship") {
      result = doSaveScholarship(d);
    } else if (action === "saveTransport") {
      result = doSaveTransport(d);
    } else if (action === "saveFees") {
      result = doSaveFees(d);
    } else if (action === "saveStudentContact") {
      result = doSaveStudentContact(d);
    } else if (action === "updateClassDivision") {
      result = doUpdateClassDivision(d);
    } else if (action === "saveTotalFee") {
      result = doSaveTotalFee(d);
    } else if (action === "startNewYear") {
      result = doStartNewYear(d);
    } else if (action === "archiveYear") {
      result = doArchiveYear(d);
    } else if (action === "rebuildFeesSummary") {
      result = doRebuildFeesSummary(d);
    } else if (action === "shrinkSheets") {
      result = doShrinkSheets(d);
    } else if (action === "saveEarlyLeave") {
      result = doSaveEarlyLeave(d);
    } else if (action === "markEarlyLeaveOut") {
      result = doMarkEarlyLeaveOut(d);
    } else if (action === "archivePassedOut") {
      result = doArchivePassedOut(d);
    } else if (action === "restorePassedOut") {
      result = doRestorePassedOut(d);
    } else if (action === "renumberRolls") {
      result = doRenumberRolls(d);
    } else if (action === "archiveLogNow") {
      result = doArchiveLogNow(d);
    } else if (action === "promoteStudents") {
      result = doPromoteStudents(d);
    } else if (action === "undoPromotion") {
      result = doUndoPromotion(d);
    } else if (action === "runBackup") {
      result = doRunBackup(d);
    } else if (action === "setBackupSchedule") {
      result = doSetBackupSchedule(d);
    } else {
      return wrap(cb, {status:"error", message:"Unknown action: "+action});
    }
    if (result && result.status === "error") return wrap(cb, result);
    // Students sheet बदलणाऱ्या कृतींनंतर Roster Cache रद्द करा
    if (action === "save" || action === "update" || action === "upsert" || action === "updateClassDivision" ||
        action === "saveStudentContact" || action === "promoteStudents" || action === "undoPromotion" ||
        action === "archivePassedOut" || action === "restorePassedOut" || action === "renumberRolls") {
      invalidateRosterCache_();
    }
    logAudit(d.auditUser || d.requesterUser, d.auditRole || d.requesterRole, action, result.serial || d.regNo || d.stxt1 || result.ref || "");
    var okResp = {status:"ok", rowIndex:result.rowIndex, action:action, mode:result.mode, serial:result.serial||"", count:result.count||0,
      oldIyatta:result.oldIyatta||"", oldTukdi:result.oldTukdi||"",
      oldTotalFee:result.oldTotalFee||0, newTotalFee:result.newTotalFee||0,
      extra: result.extra || null};
    if (reqKey) { try { CacheService.getScriptCache().put(reqKey, JSON.stringify(okResp), 21600); } catch (eRq2) {} }
    return wrap(cb, okResp);
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ===== AUDIT LOG (V19.8) — कोणत्या User ने कधी काय Save केले =====
// =====================================================================
// ===== V19.48 (टप्पा 6) — मासिक Log: Log_YYYY_MM (भारतीय वेळेनुसार महिना; लहान Sheet) =====
// जुनी एकच "Log" व "Log_Archive" एकदाच Log_YYYY_MM मध्ये स्थलांतरित (doMigrateLogs, फक्त Super). स्थलांतरापूर्वी वाचणारे भाग जुनी "Log" सुद्धा पाहतात.
// =====================================================================
var LOG_HEADERS = ["Timestamp","User","Role","Action","Reference"];
function logMonthKey_(d) { return Utilities.formatDate(d || new Date(), "Asia/Kolkata", "yyyy_MM"); }
function logPrevKey_(key) {
  var m = /^(\d{4})_(\d{2})$/.exec(key), y = parseInt(m[1], 10), mo = parseInt(m[2], 10) - 1;
  if (mo < 1) { y--; mo = 12; }
  return y + "_" + ("0" + mo).slice(-2);
}
function logMigrated_() { return PropertiesService.getScriptProperties().getProperty("LOG_MIGRATED") === "1"; }
function logSheet_(key, create) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), name = "Log_" + key, sh = ss.getSheetByName(name);
  if (sh || !create) return sh;
  try {
    sh = ss.insertSheet(name);
    trimNewSheet_(sh, LOG_HEADERS.length);
    sh.appendRow(LOG_HEADERS);
    try { sh.setFrozenRows(1); } catch (eF) {}
  } catch (eIns) {
    sh = ss.getSheetByName(name);   // एकाच वेळी दुसऱ्या विनंतीने बनवली असू शकते
    if (!sh) throw eIns;
  }
  return sh;
}
function logAudit(user, role, action, refInfo) {
  try {
    var sh = logSheet_(logMonthKey_(new Date()), true);
    sh.appendRow([new Date().toLocaleString("en-IN"), user || "unknown", role || "", action || "", refInfo || ""]);
  } catch(e) {
    // Logging failure should never block the main save operation
  }
}
// त्या महिन्याच्या Sheet ची शेवटची maxRows ओळी (५ स्तंभ)
function logTailRows_(key, maxRows) {
  var sh = logSheet_(key, false);
  if (!sh || sh.getLastRow() < 2) return [];
  var last = sh.getLastRow(), n = Math.min(last - 1, maxRows || 50000);
  return sh.getRange(last - n + 1, 1, n, 5).getValues();
}
// स्थलांतरापूर्वी जुनी एकच "Log" (फक्त शेवटच्या maxRows ओळी); स्थलांतरानंतर रिकामी
function logLegacyTail_(maxRows) {
  if (logMigrated_()) return [];
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log");
  if (!sh || sh.getLastRow() < 2) return [];
  var last = sh.getLastRow(), n = Math.min(last - 1, maxRows || 5000);
  return sh.getRange(last - n + 1, 1, n, 5).getValues();
}
// एका महिन्यासाठी: (स्थलांतरापूर्वी जुनी Log) + त्या महिन्याची Sheet — जुने आधी, नवे मागून
function logRowsForMonth_(key, maxRows) { return logLegacyTail_(maxRows).concat(logTailRows_(key, maxRows)); }

// ===== DASHBOARD LIVE STATS (V19.8) =====
function doGetDashboardStats(p, cb) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var todayStr = fmt(new Date());
    var now = new Date();
    var monthKey = now.getFullYear() + "-" + ("0"+(now.getMonth()+1)).slice(-2);

    var stSh = ss.getSheetByName("Students");
    var totalStudents = 0;
    var classCounts = {};
    if (stSh && stSh.getLastRow() > 1) {
      var stRows = stSh.getRange(2, 1, stSh.getLastRow()-1, 11).getValues();
      for (var i=0;i<stRows.length;i++) {
        if (!stRows[i][2] && !stRows[i][10]) continue;
        totalStudents++;
        var cls = (stRows[i][6]||"").toString().trim();
        if (cls) classCounts[cls] = (classCounts[cls]||0) + 1;
      }
    }

    function countTodayAndMonth(sheetName) {
      var sh = ss.getSheetByName(sheetName);
      var today = 0, month = 0;
      if (sh && sh.getLastRow() > 1) {
        var rows = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
        for (var i=0;i<rows.length;i++) {
          var ts = rows[i][0];
          if (!ts) continue;
          var d = new Date(ts);
          if (isNaN(d.getTime())) continue;
          var dKey = fmt(d);
          var mKey = d.getFullYear() + "-" + ("0"+(d.getMonth()+1)).slice(-2);
          if (dKey === todayStr) today++;
          if (mKey === monthKey) month++;
        }
      }
      return {today:today, month:month};
    }

    var lcCounts = countTodayAndMonth("LC");
    var bfCounts = countTodayAndMonth("Bonafide");
    var atCounts = countTodayAndMonth("Attendance");
    var monthCerts = lcCounts.month + bfCounts.month + atCounts.month;

    return wrap(cb, {
      status:"ok",
      totalStudents: totalStudents,
      todayLC: lcCounts.today,
      todayBF: bfCounts.today,
      monthCerts: monthCerts,
      classCounts: classCounts
    });
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ===== LC/BF/AT HISTORY FOR ONE STUDENT (V19.8) =====
function doGetStudentHistory(p, cb) {
  try {
    var regNo = (p.regNo || "").toString().trim().toLowerCase();
    if (!regNo) return wrap(cb, {status:"error", message:"regNo required"});
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var out = [];

    function scan(sheetName, type, detailFn) {
      var sh = ss.getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 2) return;
      var rows = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if ((r[2]||"").toString().trim().toLowerCase() !== regNo) continue;
        out.push({
          type: type,
          serial: r[1],
          date: fmt(r[0]),
          detail: detailFn(r)
        });
      }
    }
    scan("LC", "LC", function(r){ return "LC No: " + (r[17]||"—"); });
    scan("Bonafide", "BF", function(r){ return "Purpose: " + (r[13]||"—"); });
    scan("Attendance", "AT", function(r){ return "Purpose: " + (r[10]||"—"); });

    out.sort(function(a,b){ return (a.date < b.date) ? 1 : -1; });
    return wrap(cb, {status:"ok", data: out});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ===== BULK EXPORT — ALL LC/BF/AT (V19.8) =====
function doGetAllCertificates(p, cb) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var out = [];

    function scan(sheetName, type, detailFn) {
      var sh = ss.getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 2) return;
      var rows = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if (!r[2]) continue;
        out.push({
          type: type,
          serial: r[1],
          date: fmt(r[0]),
          regNo: r[2],
          firstName: r[3],
          detail: detailFn(r)
        });
      }
    }
    scan("LC", "LC", function(r){ return "LC No: " + (r[17]||"—"); });
    scan("Bonafide", "BF", function(r){ return "Purpose: " + (r[13]||"—"); });
    scan("Attendance", "AT", function(r){ return "Purpose: " + (r[10]||"—"); });

    return wrap(cb, {status:"ok", data: out});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// =====================================================
// 👥 USER MANAGEMENT PRO (V19.9)
// =====================================================
var USER_HEADERS = ["Username","Password","Role","Label","AssignedClass","Mobile No."];
var DEFAULT_USERS_SEED = [
  ["user_1","Pass@1234","master","Master User",""],
  ["user_2","Pass@1234","deo","DEO User",""],
  ["user_3","Pass@1234","cert","Certificate User",""]
];

function getUsersSheet() {
  var sh = getOrCreateSheet("Users", USER_HEADERS);
  if (sh.getLastRow() < 2) {
    for (var i=0;i<DEFAULT_USERS_SEED.length;i++) sh.appendRow(DEFAULT_USERS_SEED[i]);
  }
  return sh;
}

// =====================================================================
// ===== V19.52 — Password सुरक्षा: hash (HMAC-SHA256 + salt + गुप्त की), Users मधून Password बाहेर न पडणे =====
// डीफॉल्ट (PW_HASHED ≠ 1 व Token सक्ती बंद) असताना वर्तन पूर्वीसारखेच. hashAllPasswords() (Editor मधून) नंतर फक्त hash साठवला जातो;
// जुने साधे Password hash होईपर्यंत चालतात (passwordMatches_ दोन्ही ओळखतो).
// =====================================================================
function pwHashed_() { return PropertiesService.getScriptProperties().getProperty("PW_HASHED") === "1"; }
function isHashedPw_(v) { return /^h1\$[A-Za-z0-9_-]+\$[A-Za-z0-9_=-]+$/.test((v || "").toString()); }
function hashPassword_(plain) {
  var salt = Utilities.getUuid().replace(/[^A-Za-z0-9]/g, "").substring(0, 16);
  return "h1$" + salt + "$" + hmacB64_(salt + "|" + String(plain)).replace(/\./g, "_");
}
function passwordMatches_(stored, given) {
  stored = (stored === null || stored === undefined ? "" : stored).toString();
  given = (given === null || given === undefined ? "" : given).toString();
  if (!stored || !given) return false;
  if (isHashedPw_(stored)) {
    var m = /^h1\$([^$]+)\$(.+)$/.exec(stored);
    return safeEq_(hmacB64_(m[1] + "|" + given).replace(/\./g, "_"), m[2]);
  }
  return safeEq_(stored, given);   // जुना साधा Password (hash होईपर्यंत)
}
// Users मधून Password कधीही बाहेर पाठवायचे नसतील तेव्हा true (hash झाले असेल किंवा Token सक्ती चालू असेल)
function hidePasswordsInApi_() { return pwHashed_() || authEnforced_(); }
function pwPolicyError_(pw) {
  pw = String(pw || "");
  if (!(pwHashed_() || authEnforced_())) return "";   // पूर्वीसारखे — कोणतीही अट नाही
  if (pw.length < 8) return "नवा Password किमान ८ अक्षरी असावा.";
  if (pw === "Pass@1234" || pw === "Super@1234") return "हा डिफॉल्ट Password वापरता येत नाही — वेगळा निवडा.";
  return "";
}

function doGetUsers(p, cb) {
  try {
    if (p._authed && p.requesterRole !== "super") {
      return wrap(cb, {status:"error", message:"User यादी पाहण्याचा अधिकार फक्त Super Master ला आहे."});
    }
    var hide = hidePasswordsInApi_();
    var sh = getUsersSheet();
    ensureCols_(sh, 6);
    var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow()-1, 6).getValues() : [];
    var out = [];
    for (var i=0;i<rows.length;i++) {
      if (!rows[i][0]) continue;
      out.push({ rowIndex:i+2, username:rows[i][0], password: hide ? "" : rows[i][1], hasPassword: !!rows[i][1], role:normRole_(rows[i][2]), label:rows[i][3], assignedClass:rows[i][4]||"", mobile:phoneDigits_(rows[i][5]) })
    }
    return wrap(cb, {status:"ok", data: out, passwordsHidden: hide});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function doSaveUser(p, cb) {
  try {
    // फक्त Super Master role असलेल्या logged-in user कडूनच call यायला हवा — frontend कडून व्हॅलिडेशन,
    // तसेच backend मध्येही requesterRole तपासतो
    if (p.requesterRole !== "super") {
      return wrap(cb, {status:"error", message:"User Management अधिकार फक्त Super Master User ला आहे."});
    }
    var sh = getUsersSheet();
    var username = (p.username||"").toString().trim();
    if (!username) return wrap(cb, {status:"error", message:"Username आवश्यक आहे."});
    var roleIn = normRole_(p.role || "cert");
    if (VALID_ROLES.indexOf(roleIn) === -1) {
      return wrap(cb, {status:"error", message:"Role चुकीचा आहे."});
    }
    ensureCols_(sh, 6);
    var rowToWrite = findRowByKey(sh, 1, username);
    // मोबाईल: पाठवला नसेल तर जुना तसाच राहतो; रिकामा पाठवला तर काढला जातो
    var mobile = rowToWrite ? sh.getRange(rowToWrite, 6).getValue() : "";
    if (p.mobile !== undefined) {
      mobile = phoneDigits_(p.mobile);
      if (mobile && (mobile.length < 10 || mobile.length > 12)) return wrap(cb, {status:"error", message:"मोबाईल क्रमांक १० अंकी असावा."});
    }
    var pwIn = (p.password === undefined || p.password === null) ? "" : String(p.password);
    var oldPw = rowToWrite ? sh.getRange(rowToWrite, 2).getValue() : "";
    var pwOut;
    if (hidePasswordsInApi_()) {
      // V19.52: Password रिकामा = जुना कायम (Edit मध्ये बदलायचा नसेल तर). नवा User साठी Password आवश्यक. hash आधीच असेल तर पुन्हा hash नाही.
      if (!pwIn) {
        if (!rowToWrite || !oldPw) return wrap(cb, {status:"error", message:"नवीन User साठी Password आवश्यक आहे."});
        pwOut = oldPw;
      } else if (isHashedPw_(pwIn)) {
        return wrap(cb, {status:"error", message:"Password चुकीच्या स्वरूपात आहे."});
      } else {
        var perr = pwPolicyError_(pwIn);
        if (perr) return wrap(cb, {status:"error", message:perr});
        pwOut = pwHashed_() ? hashPassword_(pwIn) : pwIn;
      }
    } else {
      pwOut = pwIn;   // पूर्वीसारखे (साधे)
    }
    var row = [username, pwOut, roleIn, p.label||username, p.assignedClass||"", mobile];
    if (rowToWrite) {
      sh.getRange(rowToWrite, 1, 1, 6).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    logAudit(p.requesterUser, p.requesterRole, "saveUser", username);
    return wrap(cb, {status:"ok"});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function doDeleteUser(p, cb) {
  try {
    if (p.requesterRole !== "super") {
      return wrap(cb, {status:"error", message:"User Management अधिकार फक्त Super Master User ला आहे."});
    }
    var sh = getUsersSheet();
    var username = (p.username||"").toString().trim();
    var rowToWrite = findRowByKey(sh, 1, username);
    if (!rowToWrite) return wrap(cb, {status:"error", message:"User सापडला नाही."});
    sh.deleteRow(rowToWrite);
    logAudit(p.requesterUser, p.requesterRole, "deleteUser", username);
    return wrap(cb, {status:"ok"});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function doChangePassword(p, cb) {
  try {
    var sh = getUsersSheet();
    var username = (p.username||"").toString().trim();
    if (p._authed && p.requesterRole !== "super") username = (p.requesterUser || "").toString().trim(); // स्वतःचाच Password बदलता येतो
    var rowToWrite = findRowByKey(sh, 1, username);
    if (!rowToWrite) return wrap(cb, {status:"error", message:"User सापडला नाही."});
    var current = sh.getRange(rowToWrite, 2).getValue();
    if (!passwordMatches_(current, p.oldPassword)) {   // V19.52: साधा किंवा hash — दोन्ही
      return wrap(cb, {status:"error", message:"जुना Password चुकीचा आहे."});
    }
    var perr = pwPolicyError_(p.newPassword);
    if (perr) return wrap(cb, {status:"error", message:perr});
    sh.getRange(rowToWrite, 2).setValue(pwHashed_() ? hashPassword_(p.newPassword || "") : (p.newPassword || ""));
    logAudit(username, p.requesterRole, "changePassword", username);
    return wrap(cb, {status:"ok"});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// =====================================================
// ✅ CERTIFICATE VERIFY PAGE — QR कोड scan केल्यावर उघडणारे पान (V19.9)
// =====================================================
function doVerifyPage(p) {
  var type = (p.type||"").toString().toUpperCase();
  var serial = (p.serial||"").toString().trim();
  var sheetMap = { LC:"LC", BF:"Bonafide", AT:"Attendance" };
  var sheetName = sheetMap[type];
  var found = null;
  try {
    if (sheetName && serial) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName(sheetName);
      if (sh && sh.getLastRow() > 1) {
        var rows = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
        for (var i=0;i<rows.length;i++) {
          if ((rows[i][1]||"").toString() === serial) {
            found = { serial:rows[i][1], regNo:rows[i][2], firstName:rows[i][3], date:fmt(rows[i][0]) };
            break;
          }
        }
      }
    }
  } catch(e) {}

  var html;
  if (found) {
    html = "<div style='font-family:sans-serif;max-width:420px;margin:40px auto;padding:24px;border:2px solid #1a7a3a;border-radius:10px;text-align:center;background:#f4fff4'>"
      + "<div style='font-size:40px'>✅</div>"
      + "<h2 style='color:#1a7a3a'>Certificate Verified</h2>"
      + "<p><b>Type:</b> " + type + "</p>"
      + "<p><b>Serial No:</b> " + found.serial + "</p>"
      + "<p><b>Reg No:</b> " + found.regNo + "</p>"
      + "<p><b>Name:</b> " + found.firstName + "</p>"
      + "<p><b>Date:</b> " + found.date + "</p>"
      + "<p style='font-size:12px;color:#666;margin-top:16px'>Shri Govindram Seksaria High School, Pachora</p>"
      + "</div>";
  } else {
    html = "<div style='font-family:sans-serif;max-width:420px;margin:40px auto;padding:24px;border:2px solid #7a1a1a;border-radius:10px;text-align:center;background:#fff4f4'>"
      + "<div style='font-size:40px'>❌</div>"
      + "<h2 style='color:#7a1a1a'>Not Verified</h2>"
      + "<p>सदर Certificate आमच्या records मध्ये आढळले नाही.</p>"
      + "</div>";
  }
  return HtmlService.createHtmlOutput(html);
}

// =====================================================
// 📊 ANALYTICS & REPORTING (V19.9)
// =====================================================
function classifyLcReason(remarks) {
  var t = (remarks||"").toString().toLowerCase();
  if (t.indexOf("बदली") !== -1 || t.indexOf("transfer") !== -1) return "Transfer";
  if (t.indexOf("प्रगती") !== -1 || t.indexOf("progress") !== -1) return "Progress";
  return "Other";
}

function doGetAnalyticsData(p, cb) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var admissionsByYear = {};
    var genderCounts = {};
    var religionCounts = {};
    var casteCounts = {};
    var lcReasonCounts = { Transfer:0, Progress:0, Other:0 };

    var stSh = ss.getSheetByName("Students");
    if (stSh && stSh.getLastRow() > 1) {
      var rows = stSh.getRange(2, 1, stSh.getLastRow()-1, 26).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if (!r[2] && !r[10]) continue;
        var admDate = r[22];
        if (admDate) {
          var yr = (admDate instanceof Date) ? admDate.getFullYear() : new Date(admDate).getFullYear();
          if (!isNaN(yr)) admissionsByYear[yr] = (admissionsByYear[yr]||0) + 1;
        }
        var gender = (r[12]||"Unknown").toString();
        genderCounts[gender] = (genderCounts[gender]||0) + 1;
        var religion = (r[13]||"Unknown").toString().trim();
        if (religion) religionCounts[religion] = (religionCounts[religion]||0) + 1;
        var caste = (r[14]||"Unknown").toString().trim();
        if (caste) casteCounts[caste] = (casteCounts[caste]||0) + 1;
      }
    }

    var lcSh = ss.getSheetByName("LC");
    if (lcSh && lcSh.getLastRow() > 1) {
      var lcRows = lcSh.getRange(2, 1, lcSh.getLastRow()-1, 24).getValues();
      for (var j=0;j<lcRows.length;j++) {
        if (!lcRows[j][2]) continue;
        var reason = classifyLcReason(lcRows[j][23]);
        lcReasonCounts[reason] = (lcReasonCounts[reason]||0) + 1;
      }
    }

    return wrap(cb, {
      status:"ok",
      admissionsByYear: admissionsByYear,
      genderCounts: genderCounts,
      religionCounts: religionCounts,
      casteCounts: casteCounts,
      lcReasonCounts: lcReasonCounts
    });
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// =====================================================
// 👩‍🏫 CLASS TEACHER MODULE (V19.32)
// =====================================================
function ckey(iyatta, tukdi) {
  return (iyatta||"").toString().trim() + "|" + (tukdi||"").toString().trim();
}
function todayStr() { return fmt(new Date()); }
function safeParseJson(s) {
  try { return JSON.parse(s || "[]"); } catch(e) { return []; }
}

// ----- Attendance (Daily register — only absentees are stored) -----
// V19.35: महिन्यानुसार स्वतंत्र Sheet (Attendance_YYYY_MM) — प्रत्येक Save फक्त त्या महिन्याची छोटी Sheet वाचते/लिहिते
var ATT_SHEET_PREFIX = "Attendance_";
var ATT_LEGACY_SHEET = "DailyAttendance";
var ATT_SHEET_RE = /^Attendance_(\d{4})_(\d{2})$/;
var ISO_DATE_RE = /^(\d{4})-(\d{2})-(\d{2})$/;

function attMonthSheetName_(dateStr) {
  var m = ISO_DATE_RE.exec((dateStr || "").toString().trim());
  return m ? (ATT_SHEET_PREFIX + m[1] + "_" + m[2]) : null;
}

function getAttMonthSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  trimNewSheet_(sh, DAILY_ATT_HEADERS.length);
  sh.appendRow(DAILY_ATT_HEADERS);
  sh.getRange("B:B").setNumberFormat("@"); // Date कॉलम plain text — Sheets ने तारीख बदलू नये
  sh.setFrozenRows(1);
  return sh;
}

// दिलेल्या तारीख-मर्यादेतील महिन्यांच्या Sheets (+ अजून Migrate न झालेली जुनी DailyAttendance, असल्यास)
function attMonthSheets_(fromDate, toDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fromKey = ISO_DATE_RE.test(fromDate || "") ? fromDate.slice(0, 7).replace("-", "_") : "";
  var toKey = ISO_DATE_RE.test(toDate || "") ? toDate.slice(0, 7).replace("-", "_") : "";
  var out = [];
  ss.getSheets().forEach(function(sh) {
    var m = ATT_SHEET_RE.exec(sh.getName());
    if (!m) return;
    var k = m[1] + "_" + m[2];
    if (fromKey && k < fromKey) return;
    if (toKey && k > toKey) return;
    out.push(sh);
  });
  var legacy = ss.getSheetByName(ATT_LEGACY_SHEET);
  if (legacy) out.push(legacy);
  return out;
}

function readAttendanceRows_(fromDate, toDate) {
  var rows = [];
  attSheetsForRead_(fromDate, toDate).forEach(function(sh) {   // V19.53: Archive झालेल्या महिन्यांसह (फक्त वाचन)
    var last = sh.getLastRow();
    if (last > 1) rows = rows.concat(sh.getRange(2, 1, last - 1, 9).getValues());
  });
  return rows;
}

// जुनी DailyAttendance → महिनावार Sheets (एकदाच). Lock धरूनच call करा.
function migrateLegacyAttendance_() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var old = ss.getSheetByName(ATT_LEGACY_SHEET);
  if (!old) { props.setProperty("ATT_MIGRATED", "1"); return {moved: 0, unparsed: 0}; }
  var last = old.getLastRow();
  var groups = {}, unparsed = [], moved = 0;
  if (last > 1) {
    var vals = old.getRange(2, 1, last - 1, 9).getValues();
    vals.forEach(function(r) {
      if (!r[1] && !r[4] && !r[6]) return;
      var ds = fmt(r[1]);
      var nm = attMonthSheetName_(ds);
      r[1] = ds;
      if (!nm) { unparsed.push(r); return; }
      (groups[nm] = groups[nm] || []).push(r);
    });
  }
  Object.keys(groups).forEach(function(nm) {
    var sh = getAttMonthSheet_(nm);
    var rows = groups[nm];
    ensureRows_(sh, sh.getLastRow() + rows.length);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
    moved += rows.length;
  });
  if (unparsed.length) {
    var up = ss.getSheetByName("DailyAttendance_Unparsed");
    if (!up) { up = ss.insertSheet("DailyAttendance_Unparsed"); trimNewSheet_(up, 9); }
    if (up.getLastRow() === 0) up.appendRow(DAILY_ATT_HEADERS);
    ensureRows_(up, up.getLastRow() + unparsed.length);
    up.getRange(up.getLastRow() + 1, 1, unparsed.length, 9).setValues(unparsed);
  }
  var stamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyyMMdd_HHmm");
  old.setName("DailyAttendance_OLD_" + stamp);
  props.setProperty("ATT_MIGRATED", "1");
  return {moved: moved, unparsed: unparsed.length};
}

// Apps Script Editor मधून एकदा manually Run करता येते (Save वेळी हे आपोआपही होते)
function migrateDailyAttendanceToMonthly() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var res = migrateLegacyAttendance_();
    Logger.log("Migrated: " + JSON.stringify(res));
    return res;
  } finally { lock.releaseLock(); }
}

function doSaveAttendance(d) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var iyatta = (d.iyatta||"").toString().trim();
    var tukdi = (d.tukdi||"").toString().trim();
    var date = (d.date||todayStr()).toString().trim();
    if (!iyatta) return {status:"error", message:"Class आवश्यक आहे."};
    var archErr = archivedWriteError_(date); if (archErr) return {status:"error", message: archErr};   // V19.53
    var shName = attMonthSheetName_(date);
    if (!shName) return {status:"error", message:"तारीख चुकीच्या स्वरूपात आहे (YYYY-MM-DD हवी)."};
    lock.waitLock(25000); locked = true;
    if (PropertiesService.getScriptProperties().getProperty("ATT_MIGRATED") !== "1") migrateLegacyAttendance_();

    var sh = getAttMonthSheet_(shName);
    var absentList = safeParseJson(d.absentJson);
    var ts = new Date().toLocaleString("en-IN");
    var newRows = absentList.map(function(a) {
      return [ts, date, iyatta, tukdi, a.regNo||"", a.studentId||"", a.fullName||"", a.reason||"", d.markedBy||""];
    });

    // त्या दिवसाच्या/वर्गाच्या जुन्या नोंदी वगळून (फक्त या महिन्याच्या छोट्या Sheet मध्ये) बाकीच्या ठेवा
    var last = sh.getLastRow();
    var keep = [], removed = 0, oldRows = [];
    if (last > 1) {
      var vals = sh.getRange(2, 1, last - 1, 9).getValues();
      for (var i = 0; i < vals.length; i++) {
        var r = vals[i];
        if (fmt(r[1]) === date && (r[2]||"").toString().trim() === iyatta && (r[3]||"").toString().trim() === tukdi) { removed++; r[1] = fmt(r[1]); oldRows.push(r); continue; }
        r[1] = fmt(r[1]);
        keep.push(r);
      }
    }
    if (removed === 0) {
      // पहिली Save — फक्त शेवटी जोडा (सर्वात जलद)
      if (newRows.length) { ensureRows_(sh, last + newRows.length); sh.getRange(last + 1, 1, newRows.length, 9).setValues(newRows); }
    } else {
      var all = keep.concat(newRows);
      ensureRows_(sh, 1 + all.length);
      sh.getRange(2, 1, last - 1, 9).clearContent();
      if (all.length) sh.getRange(2, 1, all.length, 9).setValues(all);
      // V19.40: पुन्हा Save मध्ये काढल्या/बदललेल्या (नवीन यादीत नसलेल्या) जुन्या नोंदींचा Backup
      var newKeys = {};
      newRows.forEach(function(nr) { newKeys[(nr[4] || "").toString().trim().toLowerCase() + "|" + (nr[7] || "").toString().trim()] = true; });
      var gone = oldRows.filter(function(orow) { return !newKeys[(orow[4] || "").toString().trim().toLowerCase() + "|" + (orow[7] || "").toString().trim()]; });
      if (gone.length) backupDeleted_("Attendance", DAILY_ATT_HEADERS, gone, d.markedBy || "", "हजेरी पुन्हा Save — जुनी आवृत्ती");
    }
    try { attMarksUpsert_(date, iyatta, tukdi, absentList.length, d.markedBy || ""); } catch (eM) { attBump_(); }   // V19.43: कोणत्या वर्गाची हजेरी कधी भरली — Log वाचावी लागू नये म्हणून
    return {rowIndex: sh.getLastRow(), mode: removed ? "updated" : "created", count: absentList.length, ref: iyatta+"-"+tukdi+" "+date};
  } catch(err) {
    return {status:"error", message:err.toString()};
  } finally {
    if (locked) lock.releaseLock();
  }
}

function doGetAttendance(p, cb) {
  try {
    var iyatta = (p.iyatta||"").toString().trim();
    var tukdi = (p.tukdi||"").toString().trim();
    var date = (p.date||"").toString().trim();
    var dateFrom = (p.dateFrom||"").toString().trim();
    var dateTo = (p.dateTo||"").toString().trim();
    var rows = (date && !dateFrom && !dateTo) ? readAttendanceDateRows_(date) : readAttendanceRows_(date || dateFrom, date || dateTo);
    var out = [];
    for (var i=0;i<rows.length;i++) {
      var r = rows[i];
      var rDate = fmt(r[1]);
      if (iyatta && (r[2]||"").toString().trim() !== iyatta) continue;
      if (tukdi && (r[3]||"").toString().trim() !== tukdi) continue;
      if (date && rDate !== date) continue;
      if (dateFrom && rDate < dateFrom) continue;
      if (dateTo && rDate > dateTo) continue;
      out.push({date:rDate, iyatta:r[2], tukdi:r[3], regNo:r[4], studentId:r[5], fullName:r[6], reason:r[7], markedBy:r[8]});
    }
    out.sort(function(a,b){ return (a.date < b.date) ? 1 : (a.date > b.date ? -1 : 0); });
    var resAt = {status:"ok", data: out};
    if (archivedFlagDate_(date || dateFrom)) resAt.archived = true;   // V19.53
    return wrap(cb, resAt);
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- Notices (posted by Super Master / Master, visible to Class Teachers) -----
function doSaveNotice(d) {
  if (d.requesterRole !== "super" && d.requesterRole !== "master") {
    return {status:"error", message:"सूचना पोस्ट करण्याचा अधिकार फक्त Super Master / Master User ला आहे."};
  }
  try {
    var sh = getOrCreateSheet("Notices", NOTICE_HEADERS);
    var row = [new Date().toLocaleString("en-IN"), d.date||todayStr(), d.title||"", d.message||"", d.postedBy||d.requesterUser||"", d.targetClass||""];
    sh.appendRow(row);
    return {rowIndex: sh.getLastRow(), mode:"created", ref: d.title||""};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function doGetNotices(p, cb) {
  try {
    var targetClass = (p.targetClass||"").toString().trim();
    var limit = parseInt(p.limit||"20",10);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Notices");
    var out = [];
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2,1,sh.getLastRow()-1,6).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        var tc = (r[5]||"").toString().trim();
        if (tc && targetClass && tc !== targetClass) continue;
        out.push({timestamp:r[0], date:fmt(r[1])||r[1], title:r[2], message:r[3], postedBy:r[4], targetClass:tc});
      }
    }
    out.sort(function(a,b){ return new Date(b.timestamp) - new Date(a.timestamp); });
    return wrap(cb, {status:"ok", data: out.slice(0, limit)});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- Student Diary (Positive / Warning / Scholarship / वैयक्तिक नोंदी) — V19.42: वार्षिक Sheet "Diary_2026-27" -----
function doSaveDiary(d) {
  try {
    diaryEnsureMigrated_();
    var date = (d.date || todayStr()).toString();
    var archErrDi = archivedWriteError_(date); if (archErrDi) return {status:"error", message: archErrDi};   // V19.53
    var sh = diarySheet_(acYearOfDate_(date), true);
    var row = [new Date().toLocaleString("en-IN"), date, d.regNo||"", d.studentId||"", d.fullName||"",
      d.iyatta||"", d.tukdi||"", d.type||"", d.remark||"", d.enteredBy||""];
    sh.appendRow(row);
    return {rowIndex: sh.getLastRow(), mode:"created", ref: d.fullName||d.regNo||""};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

var DIARY_TAIL_ROWS = 3000;   // वर्ग-निहाय यादीत फक्त अलीकडच्या इतक्या ओळी वाचल्या जातात ("full=1" असल्यास सर्व)
function doGetDiary(p, cb) {
  try {
    diaryEnsureMigrated_();
    var regNo = (p.regNo||"").toString().trim().toLowerCase();
    var iyatta = (p.iyatta||"").toString().trim();
    var tukdi = (p.tukdi||"").toString().trim();
    var type = (p.type||"").toString().trim();
    var out = [], truncated = false;
    function add(r) {
      if (type && (r[7]||"").toString().trim() !== type) return;
      out.push({timestamp:r[0], date:fmt(r[1])||r[1], regNo:r[2], studentId:r[3], fullName:r[4], iyatta:r[5], tukdi:r[6], type:r[7], remark:r[8], enteredBy:r[9]});
    }
    if (regNo) {
      // एका विद्यार्थ्याचा इतिहास: सर्व वर्षांच्या Sheets मध्ये (TextFinder ने फक्त त्याच्याच ओळी)
      listYearSheets_("Diary").forEach(function(sh) {
        readRowsAt_(sh, findRowsByValue_(sh, 3, regNo), 10).forEach(add);
      });
      if (p.archive === "1") {     // Archive फाइल्समधील जुनी वर्षे (प्रत्येकी ~१ सेकंद जास्त)
        archivedYears_().forEach(function(y) {
          var ar = getArchiveSS_(y, false), ash = ar ? ar.getSheetByName("Diary_" + y) : null;
          if (ash) readRowsAt_(ash, findRowsByValue_(ash, 3, regNo), 10).forEach(add);
        });
      }
    } else {
      var year = p.year ? normAcYear_(p.year) : currentAcYear_();
      var sh2 = diarySheet_(year, false);
      if (sh2 && sh2.getLastRow() > 1) {
        var last = sh2.getLastRow();
        var startRow = (p.full === "1") ? 2 : Math.max(2, last - DIARY_TAIL_ROWS + 1);
        truncated = startRow > 2;
        sh2.getRange(startRow, 1, last - startRow + 1, 10).getValues().forEach(function(r) {
          if (iyatta && (r[5]||"").toString().trim() !== iyatta) return;
          if (tukdi && (r[6]||"").toString().trim() !== tukdi) return;
          add(r);
        });
      }
    }
    out.sort(function(a,b){ return new Date(b.timestamp) - new Date(a.timestamp); });
    return wrap(cb, {status:"ok", data: out, truncated: truncated, archivedYears: archivedYears_()});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- Transport / Updown info (प्रत्येक विद्यार्थ्याची एक नोंद, upsert) — V19.42: वार्षिक Sheet "Transport_2026-27"; नवीन वर्षात मागील वर्षाची प्रत -----
function transportSheetForRead_(year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Transport_" + year);
  if (sh) return sh;
  var all = listYearSheets_("Transport");                       // चालू वर्षाची Sheet नसेल तर सर्वात अलीकडची आधीच्या वर्षाची (फक्त वाचण्यासाठी)
  for (var i = 0; i < all.length; i++) if (all[i].getName() < "Transport_" + year) return all[i];
  return null;
}
function transportSheetForWrite_(year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Transport_" + year);
  if (sh) return sh;
  var prev = null, all = listYearSheets_("Transport");
  for (var i = 0; i < all.length; i++) if (all[i].getName() < "Transport_" + year) { prev = all[i]; break; }
  sh = getYearSheet_("Transport", year, TRANSPORT_HEADERS, true);
  if (prev && prev.getLastRow() > 1) {                            // वर्षारंभी मागील वर्षाची प्रत
    var rows = prev.getRange(2, 1, prev.getLastRow() - 1, 10).getValues();
    ensureRows_(sh, 1 + rows.length);
    sh.getRange(2, 1, rows.length, 10).setValues(rows);
  }
  return sh;
}
function doSaveTransport(d) {
  try {
    transportEnsureMigrated_();
    var year = normAcYear_(d.acYear);
    var archErrTr = archivedWriteError_(year); if (archErrTr) return {status:"error", message: archErrTr};   // V19.53
    var sh = transportSheetForWrite_(year);
    var rowToWrite = findRowByKey(sh, 2, d.regNo);
    var row = [new Date().toLocaleString("en-IN"), d.regNo||"", d.studentId||"", d.fullName||"",
      d.iyatta||"", d.tukdi||"", d.nativeVillage||"", d.travelMode||"", d.transportContact||"", d.updatedBy||""];
    if (rowToWrite) {   // बदल होत असल्यास जुनी आवृत्ती Deleted_Transport मध्ये
      var old = sh.getRange(rowToWrite, 1, 1, 10).getValues()[0];
      if ([6, 7, 8].some(function(ix) { return (old[ix] || "").toString().trim() !== (row[ix] || "").toString().trim(); })) {
        backupDeleted_("Transport", TRANSPORT_HEADERS, [old], d.updatedBy || "", "प्रवास नोंद बदल — जुनी आवृत्ती");
      }
    }
    var res = writeRow(sh, rowToWrite, row);
    res.ref = d.fullName||d.regNo||"";
    return res;
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

// ===== V19.54: फी / प्रवास (पैशाची व वैयक्तिक माहिती) वाचन-अधिकार =====
// Token असल्यास: super/master = सर्व; teacher = फक्त स्वतःच्या वर्गाचे विद्यार्थी (वर्ग Token मधून); deo/cert/peon = नाही.
// Token नसल्यास (जुना Login, AUTH_ENFORCE बंद असताना) = पूर्वीसारखेच चालते — ही मर्यादा सुरक्षा-टप्प्यात (Token सक्ती) संपते.
function feeAccessErr_(p, regNo) {
  if (!p._authed) return "";
  var r = p.requesterRole;
  if (r === "super" || r === "master") return "";
  if (r === "teacher") {
    var tc = scholarClassOf_(p._tClass);
    if (!tc.iyatta) return "तुमच्या User ला वर्ग नेमलेला नाही — Master कडून Users मध्ये वर्ग भरून घ्या.";
    if (!regNo) return "";   // वर्ग-यादी मोड: वर्ग applyAuth_ ने Token मधून आधीच सक्त केला आहे
    var st = scholarFindStudent_(regNo);
    if (!st) return "विद्यार्थी सापडला नाही.";
    if (String(st.iyatta || "").trim() !== tc.iyatta || (tc.tukdi && String(st.tukdi || "").trim() !== tc.tukdi)) return "हा विद्यार्थी आपल्या वर्गाचा नाही.";
    return "";
  }
  return "फी/प्रवास माहिती पाहण्याचा अधिकार Super / Master / वर्ग शिक्षकाला आहे.";
}
// Profile वर्ष-dropdown साठी: चालू + Archive + मुख्य फाइलमधील वर्षे (फक्त नावे — कोणतीही Sheet/Archive फाइल उघडत नाही)
function feeYearList_() {
  var cur = currentAcYear_(), seen = {}, list = [];
  function add(y) { if (/^\d{4}-\d{2}$/.test(y) && !seen[y]) { seen[y] = 1; list.push(y); } }
  add(cur);
  ["Fees", "Transport"].forEach(function(b) { listYearSheets_(b).forEach(function(sh) { add(sh.getName().slice(b.length + 1)); }); });
  archivedYears_().forEach(add);
  list.sort().reverse();
  return list.map(function(y) { return {year: y, archived: isArchivedYear_(y), current: y === cur}; });
}
function doGetTransport(p, cb) {
  try {
    transportEnsureMigrated_();
    var regNo = (p.regNo||"").toString().trim().toLowerCase();
    var iyatta = (p.iyatta||"").toString().trim();
    var tukdi = (p.tukdi||"").toString().trim();
    var accErr = feeAccessErr_(p, regNo); if (accErr) return wrap(cb, {status:"error", code:"forbidden", message: accErr});   // V19.54
    var tYear = p.year ? normAcYear_(p.year) : currentAcYear_();
    var tArch = isArchivedYear_(tYear);
    var sh = tArch ? getYearSheet_("Transport", tYear, null, false) : transportSheetForRead_(tYear);   // V19.54: Archive वर्षासाठी नेमकी त्याच वर्षाची Sheet (दुसऱ्या वर्षाची नव्हे)
    if (tArch && !sh && !archiveSS_(tYear)) return wrap(cb, {status:"error", code:"archive_missing", message:"📦 " + tYear + " ची Archive फाइल उघडता आली नाही — रिकामे दाखवत नाही. Archive फाइल/परवानगी तपासा."});
    var out = [];
    if (sh && sh.getLastRow() > 1) {
      var classSet = null;
      if (!regNo && (iyatta || tukdi)) {      // वर्ग-बढती/तुकडी बदलानंतरही अचूक: विद्यार्थ्यांची सध्याची वर्ग-यादी वापरा
        classSet = {};
        getRosterCached_(iyatta, tukdi, false).data.forEach(function(st) { classSet[(st.regNo || "").toString().trim().toLowerCase()] = st; });
      }
      var rows = sh.getRange(2,1,sh.getLastRow()-1,10).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i], key = (r[1]||"").toString().trim().toLowerCase();
        if (regNo && key !== regNo) continue;
        if (classSet && !classSet[key]) continue;
        var cur = classSet ? classSet[key] : null;
        out.push({regNo:r[1], studentId:r[2], fullName:r[3], iyatta: cur ? cur.iyatta : r[4], tukdi: cur ? cur.tukdi : r[5], nativeVillage:r[6], travelMode:r[7], transportContact:r[8], updatedBy:r[9]});
      }
    }
    var tRes = {status:"ok", data: out};
    if (tArch) tRes.archived = true;                                           // V19.54: फक्त Archive असल्यास
    if (sh) { try { var sy = sheetYearOf_(sh.getName()); if (sy && sy !== tYear) tRes.sourceYear = sy; } catch (eSy) {} }   // मागील वर्षाची नोंद दाखवली असल्यास
    return wrap(cb, tRes);
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}


// =====================================================================
// ===== V19.55 — "📦 Archive" स्क्रीन: जुन्या वर्षांचे अहवाल (फक्त वाचन; कोणतीही Sheet तयार/बदल/मिटवली जात नाही) =====
// • अधिकार: फक्त super/master, Token सह (p._authed) — शिक्षक/deo/cert/peon नाही.
// • फी: त्या वर्षाची "Fees_<वर्ष>" जमा-नोंदवही एकदा वाचून; वर्ग = त्या नोंदीत साठवलेला (आजचा नाही); एकूण फी = त्या वर्षाच्या नोंदीतील TotalFee.
// • हजेरी / शिष्यवृत्ती: विद्यमान getAttendanceOverview / getScholarship (Archive वर आधीच चालतात) — क्लायंट वापरतो.
// =====================================================================
function archAuth_(p) {
  if (!p._authed) return "📦 Archive स्क्रीनसाठी Token सह Login आवश्यक आहे — कृपया पुन्हा Login करा.";
  if (p.requesterRole !== "super" && p.requesterRole !== "master") return "📦 Archive स्क्रीनचा अधिकार फक्त Super / Master ला आहे.";
  return "";
}
// फक्त वाचन: कधीही Sheet तयार करत नाही. परत: {sh, from, err}
function archReadSheet_(base, year) {
  var name = base + "_" + year;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (sh) return {sh: sh, from: "main"};
  if (isArchivedYear_(year)) {
    var ar = archiveSS_(year);
    if (!ar) return {sh: null, err: "📦 " + year + " ची Archive फाइल उघडता आली नाही — रिकामे दाखवत नाही. Archive फाइल/परवानगी तपासा."};
    return {sh: ar.getSheetByName(name), from: "archive"};
  }
  return {sh: null, from: "none"};
}
function doGetArchiveYears(p, cb) {
  try {
    var e = archAuth_(p); if (e) return wrap(cb, {status: "error", code: "forbidden", message: e});
    var cur = currentAcYear_(), seen = {}, list = [], props = PropertiesService.getScriptProperties();
    function add(y) { if (/^\d{4}-\d{2}$/.test(y) && !seen[y]) { seen[y] = 1; list.push(y); } }
    add(cur);
    ["Fees", "Transport", "Scholarship"].forEach(function(b) { listYearSheets_(b).forEach(function(sh) { add(sh.getName().slice(b.length + 1)); }); });
    archivedYears_().forEach(add);
    list.sort().reverse();
    // Archive फाइल उघडत नाही (प्रत्येकी ~१ से.) — फक्त फाइलची नोंद (Script Property) आहे का ते पाहतो
    var years = list.map(function(y) { return {year: y, current: y === cur, archived: isArchivedYear_(y), hasFile: !!props.getProperty("ARCHIVE_ID_" + y)}; });
    return wrap(cb, {status: "ok", years: years, current: cur});
  } catch (err) {
    return wrap(cb, {status: "error", message: err.toString()});
  }
}
function arClassCmp_(a, b) { return classSortKey_(a) - classSortKey_(b) || (String(a) < String(b) ? -1 : (String(a) > String(b) ? 1 : 0)); }
function doGetArchiveReport(p, cb) {
  try {
    var e = archAuth_(p); if (e) return wrap(cb, {status: "error", code: "forbidden", message: e});
    var year = (p.year || "").toString().trim();
    if (!/^\d{4}-\d{2}$/.test(year)) return wrap(cb, {status: "error", message: "वर्ष YYYY-YY स्वरूपात हवे."});
    var section = (p.section || "").toString().trim();
    var archived = isArchivedYear_(year);
    if (section === "fees") {
      feesEnsureMigrated_();
      var rd = archReadSheet_("Fees", year);
      if (rd.err) return wrap(cb, {status: "error", code: "archive_missing", message: rd.err});
      var ev = rd.sh, assumedTotal = getTotalFeePerStudent();
      var vals = (ev && ev.getLastRow() > 1) ? ev.getRange(2, 1, ev.getLastRow() - 1, 11).getValues() : [];
      var map = {}, order = [];
      vals.forEach(function(r) {
        var key = (r[1] || "").toString().trim().toLowerCase(); if (!key) return;
        var x = map[key];
        if (!x) { x = map[key] = {reg: r[1], name: r[3], cls: r[4], div: r[5], paid: 0, total: null, opening: 0, n: 0}; order.push(key); }
        if ((r[10] || "").toString() === "OPENING") { x.opening += parseFloat(r[9]) || 0; if (!x.n) { x.cls = r[4]; x.div = r[5]; x.name = r[3] || x.name; } return; }
        x.n++; x.paid = parseFloat(r[8]) || 0;
        var t = parseFloat(r[7]); if (t > 0) x.total = t;
        x.cls = r[4]; x.div = r[5]; x.name = r[3] || x.name;
      });
      var rows = [], byClass = {}, tot = {students: 0, payers: 0, paid: 0, opening: 0, balance: 0, assumed: 0};
      order.forEach(function(k) {
        var x = map[k], assumed = x.total === null, total = assumed ? assumedTotal : x.total;
        var bal = Math.max(0, x.opening + total - x.paid);
        rows.push([x.reg, x.name, x.cls, x.div, total, x.opening, x.paid, bal, x.n, assumed ? 1 : 0]);
        var c = byClass[x.cls] || (byClass[x.cls] = {cls: x.cls, students: 0, payers: 0, paid: 0, opening: 0, balance: 0});
        [c, tot].forEach(function(a) { a.students++; if (x.paid > 0) a.payers++; a.paid += x.paid; a.opening += x.opening; a.balance += bal; });
        if (assumed) tot.assumed++;
      });
      rows.sort(function(a, b) { return arClassCmp_(a[2], b[2]) || (String(a[3]) < String(b[3]) ? -1 : (String(a[3]) > String(b[3]) ? 1 : 0)) || (String(a[1]) < String(b[1]) ? -1 : 1); });
      var classes = Object.keys(byClass).map(function(k) { return byClass[k]; }).sort(function(a, b) { return arClassCmp_(a.cls, b.cls); });
      var res = {status: "ok", section: "fees", year: year, classes: classes, total: tot, rows: rows, hasSheet: !!ev};
      if (archived) res.archived = true;
      return wrap(cb, res);
    }
    if (section === "transport") {
      transportEnsureMigrated_();
      var rt = archReadSheet_("Transport", year);
      if (rt.err) return wrap(cb, {status: "error", code: "archive_missing", message: rt.err});
      var ts = rt.sh, tv = (ts && ts.getLastRow() > 1) ? ts.getRange(2, 1, ts.getLastRow() - 1, 10).getValues() : [];
      var trows = tv.filter(function(r) { return (r[1] || "").toString().trim(); }).map(function(r) { return [r[1], r[3], r[4], r[5], r[6], r[7], r[8]]; });
      trows.sort(function(a, b) { return arClassCmp_(a[2], b[2]) || (String(a[3]) < String(b[3]) ? -1 : (String(a[3]) > String(b[3]) ? 1 : 0)) || (String(a[1]) < String(b[1]) ? -1 : 1); });
      var tres = {status: "ok", section: "transport", year: year, rows: trows, hasSheet: !!ts};
      if (archived) tres.archived = true;
      return wrap(cb, tres);
    }
    return wrap(cb, {status: "error", message: "अज्ञात विभाग (fees / transport). हजेरी व शिष्यवृत्तीसाठी विद्यमान action वापरा."});
  } catch (err) {
    return wrap(cb, {status: "error", message: err.toString()});
  }
}

// =====================================================================
// ===== V19.56 — एका विद्यार्थ्याचा सर्व वर्षांचा इतिहास (फक्त वाचन) =====
// • meta=1 → विद्यार्थी-ओळख + वर्ष-यादी (कोणतीही Archive फाइल उघडत नाही). नंतर क्लायंट वर्षागणिक विनंत्या करतो:
//   part=core (वर्ग, हजेरी, फी, शिष्यवृत्ती, Diary, प्रवास) आणि part=el (घरी गेलेले — आठवडानिहाय EL_ Sheets).
// • सर्व शोध findRowsByValue_ (TextFinder) ने — पूर्ण Sheet वाचत नाही (AttMarks फक्त त्या महिन्याची छोटी Sheet).
// • अधिकार: Token सह; super/master सर्व; teacher फक्त आज वर्ग-यादीत असलेल्या स्वतःच्या वर्गाचे; इतर नाही. पूर्वी-उत्तीर्ण (PassedOut_) विद्यार्थी = फक्त super/master.
// =====================================================================
function hsAuth_(p, regNo) {
  if (!p._authed) return "📚 इतिहासासाठी Token सह Login आवश्यक आहे — कृपया पुन्हा Login करा.";
  var r = p.requesterRole;
  if (r === "super" || r === "master") return "";
  if (r === "teacher") {
    var tc = scholarClassOf_(p._tClass);
    if (!tc.iyatta) return "तुमच्या User ला वर्ग नेमलेला नाही — Master कडून Users मध्ये वर्ग भरून घ्या.";
    var st = scholarFindStudent_(regNo);
    if (!st || String(st.iyatta || "").trim() !== tc.iyatta || (tc.tukdi && String(st.tukdi || "").trim() !== tc.tukdi)) return "हा विद्यार्थी आपल्या वर्गाचा नाही.";
    return "";
  }
  return "सर्व वर्षांचा इतिहास पाहण्याचा अधिकार Super / Master / वर्ग शिक्षकाला आहे.";
}
function hsP2_(n) { return ("0" + n).slice(-2); }
function hsPassedOut_(regNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), found = null;
  ss.getSheets().forEach(function(sh) {
    if (found || sh.getName().indexOf(ARCHIVE_PREFIX) !== 0) return;
    var rows = findRowsByValue_(sh, 3, regNo);
    if (!rows.length) return;
    var r = readRowsAt_(sh, [rows[rows.length - 1]], 11)[0];
    found = {fullName: r[10], iyatta: (r[6] || "").toString(), tukdi: (r[7] || "").toString(), acYear: normAcYear_(r[1]), sheet: sh.getName()};
  });
  return found;
}
function hsYearList_() {
  var seen = {}, list = [];
  function add(y) { if (/^\d{4}-\d{2}$/.test(y) && !seen[y]) { seen[y] = 1; list.push(y); } }
  feeYearList_().forEach(function(o) { add(o.year); });
  ["Diary", "Scholarship"].forEach(function(b) { listYearSheets_(b).forEach(function(sh) { add(sh.getName().slice(b.length + 1)); }); });
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function(sh) { var y = sheetYearOf_(sh.getName()); if (y) add(y); });
  list.sort().reverse();
  return list;
}
function hsClassOfYear_(regNo, year, passed) {
  var st = scholarFindStudent_(regNo);
  if (st && st.acYear && normAcYear_(st.acYear) === year) return {iyatta: String(st.iyatta || ""), tukdi: String(st.tukdi || ""), roll: st.rollNo || "", source: "Students"};
  var pl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PromotionLog");
  if (pl && pl.getLastRow() > 1) {
    var hit = null;
    readRowsAt_(pl, findRowsByValue_(pl, 3, regNo), 11).forEach(function(r) {
      if ((r[10] || "").toString() === "Y") return;
      if (normAcYear_(r[6]) === year) hit = r;
    });
    if (hit) return {iyatta: String(hit[4] || ""), tukdi: String(hit[5] || ""), roll: "", source: "PromotionLog"};
  }
  if (passed && passed.acYear === year) return {iyatta: passed.iyatta, tukdi: passed.tukdi, roll: "", source: "PassedOut"};
  return {iyatta: "", tukdi: "", roll: "", source: ""};
}
function hsRows_(base, year, col, regNo, ncols) {
  var sh = getYearSheet_(base, year, null, false);   // चालू फाइल, नसेल तर Archive (फक्त वाचन; Sheet तयार करत नाही)
  if (!sh) return {has: false, rows: []};
  return {has: true, rows: readRowsAt_(sh, findRowsByValue_(sh, col, regNo), ncols)};
}
function hsYearCore_(regNo, year, passed) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  feesEnsureMigrated_(); transportEnsureMigrated_(); diaryEnsureMigrated_();   // विद्यमान action प्रमाणेच (flag-नियंत्रित, एकदाच)
  var out = {year: year, archived: isArchivedYear_(year)};
  var cls = hsClassOfYear_(regNo, year, passed);

  // फी
  var f = hsRows_("Fees", year, 2, regNo, 11), inst = [], prev = 0, opening = 0, total = null, paid = 0;
  f.rows.forEach(function(r) {
    if ((r[10] || "").toString() === "OPENING") { opening += parseFloat(r[9]) || 0; if (!cls.iyatta) cls = {iyatta: String(r[4] || ""), tukdi: String(r[5] || ""), roll: "", source: "फी-नोंद"}; return; }
    var cum = parseFloat(r[8]) || 0;
    inst.push({date: fmt(r[0]) || r[0], amount: cum - prev, by: r[10]});
    prev = cum; paid = cum;
    var t = parseFloat(r[7]); if (t > 0) total = t;
    if (!cls.iyatta) cls = {iyatta: String(r[4] || ""), tukdi: String(r[5] || ""), roll: "", source: "फी-नोंद"};
  });
  var assumed = (f.rows.length > 0 && total === null) || (f.rows.length > 0 && !inst.length);
  var curTotal = getTotalFeePerStudent();
  var useTotal = total === null ? curTotal : total;
  out.fees = {has: f.rows.length > 0, total: useTotal, totalAssumed: f.rows.length > 0 && total === null, opening: opening, paid: paid,
    balance: Math.max(0, opening + useTotal - paid), installments: inst};

  // शिष्यवृत्ती
  var sc = hsRows_("Scholarship", year, 4, regNo, SCHOLAR_HEADERS.length);
  out.scholarship = sc.rows.map(function(r) { return scholarRowToObj_(r, null); });
  if (!cls.iyatta && out.scholarship.length) cls = {iyatta: String(out.scholarship[0].iyatta || ""), tukdi: String(out.scholarship[0].tukdi || ""), roll: "", source: "शिष्यवृत्ती-नोंद"};

  // Diary
  var dy = hsRows_("Diary", year, 3, regNo, 10);
  out.diary = dy.rows.map(function(r) { return {date: fmt(r[1]), type: (r[7] || "").toString(), remark: r[8], by: r[9]}; });
  if (!cls.iyatta && dy.rows.length) cls = {iyatta: String(dy.rows[0][5] || ""), tukdi: String(dy.rows[0][6] || ""), roll: "", source: "Diary-नोंद"};

  // प्रवास
  var tr = hsRows_("Transport", year, 2, regNo, 10);
  var tl = tr.rows.length ? tr.rows[tr.rows.length - 1] : null;
  out.transport = tl ? {nativeVillage: tl[6], travelMode: tl[7], transportContact: tl[8]} : null;
  if (!cls.iyatta && tl) cls = {iyatta: String(tl[4] || ""), tukdi: String(tl[5] || ""), roll: "", source: "प्रवास-नोंद"};

  // हजेरी — गैरहजर नोंदी (Attendance_YYYY_MM) + त्या वर्गाची हजेरी भरलेले दिवस (AttMarks_YYYY_MM)
  var ys = acYearStart_(year), months = [], totAbs = 0, totMarked = 0, anyMarked = false, absDates = [];
  for (var i = 0; i < 12; i++) {
    var m = 6 + i, yy = ys; if (m > 12) { m -= 12; yy = ys + 1; }
    var key = yy + "_" + hsP2_(m), nm = "Attendance_" + key;
    var ash = ss.getSheetByName(nm) || archiveSheet_(nm);
    var arows = ash ? readRowsAt_(ash, findRowsByValue_(ash, 5, regNo), 9) : [];
    var seenD = {}, list = [];
    arows.forEach(function(r) { var d = fmt(r[1]); if (seenD[d]) return; seenD[d] = 1; list.push({date: d, reason: (r[7] || "").toString()}); });
    var marked = null, mk = attMarksMonthRows_(key);
    if (mk && cls.iyatta) {
      var ds = {};
      mk.forEach(function(r) {
        if ((r[1] || "").toString().trim() !== cls.iyatta) return;
        if (cls.tukdi && (r[2] || "").toString().trim() !== cls.tukdi) return;
        ds[fmt(r[0])] = true;
      });
      marked = Object.keys(ds).length; anyMarked = true; totMarked += marked;
    }
    totAbs += list.length;
    list.forEach(function(x) { absDates.push(x); });
    if (list.length || marked) months.push({month: yy + "-" + hsP2_(m), absent: list.length, markedDays: marked});
  }
  out.attendance = {absentDays: totAbs, markedDays: anyMarked ? totMarked : null,
    pct: (anyMarked && totMarked > 0) ? Math.round(Math.max(0, totMarked - totAbs) * 1000 / totMarked) / 10 : null,
    months: months, absentDates: absDates};
  out.cls = cls;
  return out;
}
function hsYearEL_(regNo, year) {
  elEnsureMigrated_();
  var y = acYearStart_(year), lo = y + "-06-01", hi = (y + 1) + "-05-31", sheets = [];
  function collect(book) {
    book.getSheets().forEach(function(sh) {
      var m = /^EL_(\d{4}-\d{2}-\d{2})$/.exec(sh.getName());
      if (m && m[1] >= lo && m[1] <= hi) sheets.push(sh);
    });
  }
  collect(SpreadsheetApp.getActiveSpreadsheet());
  if (isArchivedYear_(year)) { var ar = archiveSS_(year); if (ar) collect(ar); }
  var recs = [];
  sheets.forEach(function(sh) {
    readRowsAt_(sh, findRowsByValue_(sh, 5, regNo), 14).forEach(function(r) {
      recs.push({date: fmt(r[1]), type: (r[7] || "").toString(), period: parseInt(r[8], 10) || 0, status: (r[11] || "").toString(), recordedBy: r[9], markedBy: r[12] || ""});
    });
  });
  recs.sort(function(a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : a.period - b.period); });
  var gone = recs.filter(function(x) { return x.status === "gone"; }).length;
  return {year: year, archived: isArchivedYear_(year), sheetsScanned: sheets.length, records: recs, gone: gone};
}
function doGetStudentYearHistory(p, cb) {
  try {
    var regNo = (p.regNo || "").toString().trim();
    if (!regNo) return wrap(cb, {status: "error", message: "Reg No आवश्यक आहे."});
    var e = hsAuth_(p, regNo); if (e) return wrap(cb, {status: "error", code: "forbidden", message: e});
    if (p.meta === "1") {
      var st = scholarFindStudent_(regNo), student;
      if (st) student = {status: "active", regNo: st.regNo, fullName: st.fullName, iyatta: st.iyatta, tukdi: st.tukdi, acYear: st.acYear, rollNo: st.rollNo || ""};
      else {
        var po = hsPassedOut_(regNo);
        student = po ? {status: "passedout", regNo: regNo, fullName: po.fullName, iyatta: po.iyatta, tukdi: po.tukdi, acYear: po.acYear} : {status: "unknown", regNo: regNo};
      }
      return wrap(cb, {status: "ok", student: student, years: hsYearList_(), current: currentAcYear_()});
    }
    var year = (p.year || "").toString().trim();
    if (!/^\d{4}-\d{2}$/.test(year)) return wrap(cb, {status: "error", message: "वर्ष YYYY-YY स्वरूपात हवे."});
    if (isArchivedYear_(year) && !archiveSS_(year)) return wrap(cb, {status: "error", code: "archive_missing", message: "📦 " + year + " ची Archive फाइल उघडता आली नाही — रिकामे दाखवत नाही. Archive फाइल/परवानगी तपासा."});
    var res;
    if ((p.part || "core") === "el") res = hsYearEL_(regNo, year);
    else {
      var passed = scholarFindStudent_(regNo) ? null : hsPassedOut_(regNo);
      res = hsYearCore_(regNo, year, passed);
    }
    res.status = "ok"; res.part = (p.part === "el") ? "el" : "core";
    return wrap(cb, res);
  } catch (err) {
    return wrap(cb, {status: "error", message: err.toString()});
  }
}

// ----- Fees (टप्प्याटप्प्याने जमा होणारी रक्कम — प्रत्येक एन्ट्री ही एक जमा नोंद, त्या क्षणापर्यंतच्या
//        cumulative TotalFee/FeePaid/PendingFee सह लॉग केली जाते — V19.33) -----
function doSaveFees(d) {
  try {
    if (d.requesterRole !== "super" && d.requesterRole !== "master" && d.requesterRole !== "teacher") {
      return {status:"error", message:"Fees माहिती Save करण्याचा अधिकार Super Master / Master / Class Teacher ला आहे."};
    }
    // Class Teacher फक्त स्वतःच्या वर्गाच्याच विद्यार्थ्यांची Fees नोंद करू शकतो
    if (d.requesterRole === "teacher") {
      var tIyatta = (d.teacherIyatta||"").toString().trim();
      var tTukdi = (d.teacherTukdi||"").toString().trim();
      if ((d.iyatta||"").toString().trim() !== tIyatta || (d.tukdi||"").toString().trim() !== tTukdi) {
        return {status:"error", message:"आपण फक्त स्वतःच्या वर्गाच्याच विद्यार्थ्यांची Fees नोंद करू शकता."};
      }
    }
    var amount = parseFloat(d.amountPaid)||0;
    if (amount <= 0) return {status:"error", message:"जमा रक्कम बरोबर टाका."};
    var archErrFee = archivedWriteError_(normAcYear_(d.acYear)); if (archErrFee) return {status:"error", message: archErrFee};   // V19.53
    feesEnsureMigrated_();
    var feeLock = LockService.getScriptLock();
    feeLock.waitLock(20000);
    try {
      // V19.42: वार्षिक "Fees_<वर्ष>" (प्रत्येक जमा नोंद) + "FeesSummary_<वर्ष>" (प्रत्येक विद्यार्थ्याची एक ओळ) — जमा रक्कम सारांशातून मिळते, मोठी Sheet वाचावी लागत नाही
      var year = normAcYear_(d.acYear);
      var ev = feesEventSheet_(year, true);
      var sum = feesSummaryEnsure_(year) || feesSummarySheet_(year, true);
      var sumRow = findRowByKey(sum, 1, (d.regNo||"").toString().trim());
      ensureCols_(sum, 12);
      var prev = sumRow ? sum.getRange(sumRow, 1, 1, 12).getValues()[0] : null;
      var prevPaid = prev ? (parseFloat(prev[6]) || 0) : 0;
      var openingDue = prev ? (parseFloat(prev[11]) || 0) : 0;
      var totalFee = getTotalFeePerStudent();
      var feePaid = prevPaid + amount;
      var pendingFee = Math.max(0, openingDue + totalFee - feePaid);
      var nowStr = new Date().toLocaleString("en-IN");
      ev.appendRow([nowStr, d.regNo||"", d.studentId||"", d.fullName||"", d.iyatta||"", d.tukdi||"", year, totalFee, feePaid, pendingFee, d.enteredBy||""]);
      var sRow = [d.regNo||"", d.studentId||"", d.fullName||"", d.iyatta||"", d.tukdi||"", year, feePaid, nowStr, amount,
        (prev ? (parseInt(prev[9], 10) || 0) : 0) + 1, d.enteredBy||"", openingDue];
      if (sumRow) sum.getRange(sumRow, 1, 1, 12).setValues([sRow]); else sum.appendRow(sRow);
      return {rowIndex: ev.getLastRow(), mode:"created", ref: d.fullName||d.regNo||""};
    } finally { feeLock.releaseLock(); }
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function doGetFees(p, cb) {
  try {
    feesEnsureMigrated_();
    var total = getTotalFeePerStudent();
    if (p.totalOnly === "1") return wrap(cb, {status:"ok", data: [], totalFeePerStudent: total});   // फक्त "एकूण फी" हवी असल्यास काहीही वाचत नाही
    var regNo = (p.regNo||"").toString().trim().toLowerCase();
    var iyatta = (p.iyatta||"").toString().trim();
    var tukdi = (p.tukdi||"").toString().trim();
    var accErr = feeAccessErr_(p, regNo); if (accErr) return wrap(cb, {status:"error", code:"forbidden", message: accErr});   // V19.54
    if (p.yearsOnly === "1") return wrap(cb, {status:"ok", data: [], years: feeYearList_(), currentYear: currentAcYear_()});   // V19.54: Profile चे वर्ष-dropdown
    var year = normAcYear_(p.acYear);
    var out = [];
    if (regNo) {
      // एका विद्यार्थ्याचा जमा-इतिहास (प्रत्येक installment) — फक्त त्याच्या ओळी वाचल्या जातात
      var ev = findYearSheet_("Fees", year);          // चालू फाइल, नसेल तर त्या वर्षाची Archive फाइल
      var fArch = isArchivedYear_(year);
      if (fArch && !ev && !archiveSS_(year)) return wrap(cb, {status:"error", code:"archive_missing", message:"📦 " + year + " ची Archive फाइल उघडता आली नाही — रिकामे दाखवत नाही. Archive फाइल/परवानगी तपासा."});   // V19.54
      var openingH = 0;
      if (ev) {
        readRowsAt_(ev, findRowsByValue_(ev, 2, regNo), 11).forEach(function(r) {
          if ((r[10] || "").toString() === "OPENING") { openingH += parseFloat(r[9]) || 0; return; }
          out.push({date:fmt(r[0])||r[0], regNo:r[1], studentId:r[2], fullName:r[3], iyatta:r[4], tukdi:r[5],
            acYear:r[6], totalFee:r[7], feePaid:r[8], pendingFee:r[9], updatedBy:r[10]});
        });
      }
      // V19.54: त्या वर्षाची एकूण फी = त्या वर्षातील सर्वात अलीकडच्या जमा-नोंदीतील TotalFee (सध्याचे Setting नव्हे). नोंदच नसेल तर सध्याचे Setting — तसे स्पष्ट सांगतो.
      var haveRow = out.length > 0;
      var yearTotal = haveRow ? (parseFloat(out[out.length - 1].totalFee) || total) : total;
      var paidNow = haveRow ? (parseFloat(out[out.length - 1].feePaid) || 0) : 0;
      var fRes = {status:"ok", data: [], totalFeePerStudent: total, openingDue: openingH,
        acYear: year, yearTotalFee: yearTotal, totalFeeAssumed: !haveRow && year !== currentAcYear_(),
        balance: Math.max(0, openingH + yearTotal - paidNow)};
      out.reverse();   // नवीन-ते-जुने
      fRes.data = out;
      if (fArch) fRes.archived = true;
      return wrap(cb, fRes);
    } else {
      // वर्ग / संपूर्ण शाळा: प्रत्येक विद्यार्थ्याची सर्वात अलीकडची (cumulative) नोंद — सारांश Sheet मधून
      var classSet = null;
      if (iyatta || tukdi) {
        classSet = {};
        getRosterCached_(iyatta, tukdi, false).data.forEach(function(st) { classSet[(st.regNo || "").toString().trim().toLowerCase()] = st; });
      }
      feesSummaryRows_(year).forEach(function(r) {
        var key = (r[0]||"").toString().trim().toLowerCase();
        if (classSet && !classSet[key]) return;
        var cur = classSet ? classSet[key] : null;
        var paid = parseFloat(r[6]) || 0, opening = parseFloat(r[11]) || 0;
        out.push({date:(r[7]||"").toString(), regNo:r[0], studentId:r[1], fullName:r[2], iyatta: cur ? cur.iyatta : r[3], tukdi: cur ? cur.tukdi : r[4],
          acYear:r[5], totalFee:total, feePaid:paid, openingDue:opening, pendingFee:Math.max(0, opening + total - paid), updatedBy:r[10]});
      });
    }
    return wrap(cb, {status:"ok", data: out, totalFeePerStudent: total});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- विद्यार्थी Delete करणे — पूर्ण माहिती आधी DeletedStudents Sheet मध्ये जतन करूनच (V19.32) -----
function doDeleteStudent(p, cb) {
  try {
    if (p.requesterRole !== "super" && p.requesterRole !== "master") {
      return wrap(cb, {status:"error", message:"विद्यार्थी Delete करण्याचा अधिकार फक्त Super Master / Master User ला आहे."});
    }
    var regNo = (p.regNo||"").toString().trim();
    if (!regNo) return wrap(cb, {status:"error", message:"Reg No आवश्यक आहे."});
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    if (!sh) return wrap(cb, {status:"error", message:"Students sheet सापडली नाही."});
    var rowToWrite = findRowByKey(sh, 3, regNo); // कॉलम C = Reg No
    if (!rowToWrite) return wrap(cb, {status:"error", message:"हा Reg No असलेला विद्यार्थी सापडला नाही."});

    var rowData = sh.getRange(rowToWrite, 1, 1, 31).getValues()[0];
    backupDeleted_("Students", MAIN_HEADERS, [rowData], p.requesterUser || "", "विद्यार्थी Delete");   // V19.40: SGS_Deleted_Backup मध्ये (अपयश आल्यास मुख्य फाइलमध्ये)

    sh.deleteRow(rowToWrite);
    invalidateRosterCache_();
    logAudit(p.requesterUser, p.requesterRole, "deleteStudent", regNo);
    return wrap(cb, {status:"ok", regNo: regNo});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- वर्ग/तुकडी बदल — मेंटेनन्स (V19.20) -----
function doUpdateClassDivision(d) {
  try {
    if (d.requesterRole !== "super" && d.requesterRole !== "master" && d.requesterRole !== "teacher") {
      return {status:"error", message:"अधिकार नाही."};
    }
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    if (!sh) return {status:"error", message:"Students sheet सापडली नाही."};
    var rowToWrite = findRowByKey(sh, 3, d.regNo); // कॉलम C = Reg No
    if (!rowToWrite) return {status:"error", message:"विद्यार्थी सापडला नाही."};
    var curIyatta = sh.getRange(rowToWrite, 7).getValue().toString().trim();
    var curTukdi = sh.getRange(rowToWrite, 8).getValue().toString().trim();
    if (d.requesterRole === "teacher") {
      var tIyatta = (d.teacherIyatta||"").toString().trim();
      if (curIyatta !== tIyatta) {
        return {status:"error", message:"हा विद्यार्थी आपल्या वर्गाचा नाही."};
      }
      if ((d.iyatta||"").toString().trim() !== curIyatta) {
        return {status:"error", message:"Class Teacher फक्त तुकडी बदलू शकतो, वर्ग नाही."};
      }
    }
    sh.getRange(rowToWrite, 7, 1, 2).setValues([[d.iyatta||"", d.tukdi||""]]);
    return {rowIndex: rowToWrite, mode:"updated", ref: d.regNo||"",
      oldIyatta: curIyatta, oldTukdi: curTukdi};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

// ===== शाळेची एकूण फी ठरवणे — फक्त Super Master / Master (Maintenance पान) =====
function doSaveTotalFee(d) {
  try {
    if (d.requesterRole !== "super" && d.requesterRole !== "master") {
      return {status:"error", message:"एकूण फी ठरवण्याचा अधिकार फक्त Super Master / Master ला आहे."};
    }
    var amount = parseFloat(d.totalFee);
    if (isNaN(amount) || amount <= 0) {
      return {status:"error", message:"एकूण फीची रक्कम बरोबर टाका (0 पेक्षा जास्त)."};
    }
    var oldAmount = getTotalFeePerStudent();
    PropertiesService.getScriptProperties().setProperty('TOTAL_FEE_PER_STUDENT', String(amount));
    return {rowIndex: 0, mode: "updated", ref: "TotalFee:" + amount, oldTotalFee: oldAmount, newTotalFee: amount};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

// ----- सांख्यिकी अहवाल — एकाच call मध्ये सर्व माहिती एकत्र (Performance — V19.21) -----
function statsAge(dobStr, asOfStr) {
  if (!dobStr) return null;
  var d = new Date(dobStr);
  if (isNaN(d.getTime())) return null;
  var asOf = new Date(asOfStr);
  var age = asOf.getFullYear() - d.getFullYear();
  var mo = asOf.getMonth() - d.getMonth();
  if (mo < 0 || (mo === 0 && asOf.getDate() < d.getDate())) age--;
  return age;
}

function doGetStatsReport(p, cb) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var today = fmt(new Date());
    var asOf = (p.asOfDate || "2026-09-30").toString();
    var classMap = {};

    var stSh = ss.getSheetByName("Students");
    if (stSh && stSh.getLastRow() > 1) {
      var rows = stSh.getRange(2, 1, stSh.getLastRow()-1, 31).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if (!r[10] && !r[2]) continue;
        var iyatta = (r[6]||"").toString().trim();
        var tukdi = (r[7]||"").toString().trim();
        var key = iyatta + "|" + tukdi;
        if (!classMap[key]) classMap[key] = {
          iyatta: iyatta, tukdi: tukdi, total: 0, absent: 0, fees: 0,
          female: {}, male: {}, ageCounts: {}
        };
        var c = classMap[key];
        c.total++;
        var gender = r[12];
        var cat = (r[29]||"").toString().trim();
        var minority = (r[30]||"").toString().trim().toLowerCase();
        var bucket = gender === "Female" ? c.female : (gender === "Male" ? c.male : null);
        if (bucket) {
          if (CATEGORY_LIST_SRV.indexOf(cat) !== -1) bucket[cat] = (bucket[cat]||0) + 1;
          if (minority === "yes") bucket.Minority = (bucket.Minority||0) + 1;
        }
        var age = statsAge(fmt(r[16]), asOf);
        if (age !== null) c.ageCounts[age] = (c.ageCounts[age]||0) + 1;
      }
    }

    var attRows = readAttendanceDateRows_(today);
    for (var j=0;j<attRows.length;j++) {
      var ar = attRows[j];
      if (fmt(ar[1]) !== today) continue;
      var key2 = (ar[2]||"").toString().trim() + "|" + (ar[3]||"").toString().trim();
      if (classMap[key2]) classMap[key2].absent++;
    }

    // Fees sheet मध्ये आता प्रत्येक installment ला cumulative FeePaid असतो, त्यामुळे सरळ बेरीज न करता
    // प्रत्येक विद्यार्थ्याची फक्त सर्वात अलीकडची (शेवटची) नोंद घ्यावी लागते, मगच वर्गनिहाय बेरीज करायची
    feesEnsureMigrated_();
    feesSummaryRows_(currentAcYear_()).forEach(function(fr) {      // V19.42: चालू वर्षाच्या FeesSummary मधून (प्रत्येक विद्यार्थ्याची एक ओळ)
      var key3 = (fr[3]||"").toString().trim() + "|" + (fr[4]||"").toString().trim();
      if (classMap[key3]) classMap[key3].fees += parseFloat(fr[6]) || 0;
    });

    var classes = Object.keys(classMap).map(function(k){ return classMap[k]; });
    classes.sort(function(a,b){
      var ka = a.iyatta+"|"+a.tukdi, kb = b.iyatta+"|"+b.tukdi;
      return ka < kb ? -1 : (ka > kb ? 1 : 0);
    });
    return wrap(cb, {status:"ok", classes: classes, today: today, asOfDate: asOf, categoryList: CATEGORY_LIST_SRV});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- विद्यार्थ्याचे WhatsApp Mobile + इतर Mobile (Class Teacher entry) -----
function doSaveStudentContact(d) {
  try {
    var sh = getOrCreateSheet("StudentContacts", CONTACTS_HEADERS);
    var rowToWrite = findRowByKey(sh, 2, d.regNo);
    var row = [new Date().toLocaleString("en-IN"), d.regNo||"", d.studentId||"", d.fullName||"",
      d.iyatta||"", d.tukdi||"", d.whatsappMobile||"", d.otherMobile||"", d.updatedBy||""];
    var res = writeRow(sh, rowToWrite, row);
    res.ref = d.fullName||d.regNo||"";
    return res;
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function doGetStudentContacts(p, cb) {
  try {
    var regNo = (p.regNo||"").toString().trim().toLowerCase();
    var iyatta = (p.iyatta||"").toString().trim();
    var tukdi = (p.tukdi||"").toString().trim();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("StudentContacts");
    var out = [];
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2,1,sh.getLastRow()-1,9).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if (regNo && (r[1]||"").toString().trim().toLowerCase() !== regNo) continue;
        if (!regNo && iyatta && (r[4]||"").toString().trim() !== iyatta) continue;
        if (!regNo && tukdi && (r[5]||"").toString().trim() !== tukdi) continue;
        out.push({regNo:r[1], studentId:r[2], fullName:r[3], iyatta:r[4], tukdi:r[5], whatsappMobile:r[6], otherMobile:r[7], updatedBy:r[8]});
      }
    }
    return wrap(cb, {status:"ok", data: out});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- Class Teacher's student list (filtered by their assigned class) — V19.35: CacheService सह -----
var ROSTER_CACHE_TTL = 600; // सेकंद (१० मिनिटे) — कोणतीही Student-बदल Save झाल्यास Cache आपोआप रद्द होतो

function rosterVer_() {
  try {
    var c = CacheService.getScriptCache();
    var v = c.get("roster_ver");
    if (!v) { v = String(Date.now()); c.put("roster_ver", v, 21600); }
    return v;
  } catch (e) { return "0"; }
}
function invalidateRosterCache_() {
  try { CacheService.getScriptCache().put("roster_ver", String(Date.now()), 21600); } catch (e) {}
}
function cacheGetBig_(key) {
  try {
    var c = CacheService.getScriptCache();
    var meta = c.get(key);
    if (!meta) return null;
    var n = parseInt(meta, 10);
    if (!n || n < 1) return null;
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(key + "_" + i);
    var parts = c.getAll(keys);
    var s = "";
    for (var j = 0; j < n; j++) {
      var part = parts[key + "_" + j];
      if (part === null || part === undefined) return null;
      s += part;
    }
    return JSON.parse(s);
  } catch (e) { return null; }
}
function cachePutBig_(key, obj, ttl) {
  try {
    var s = JSON.stringify(obj);
    var CH = 30000; // देवनागरी ३ bytes/अक्षर — एका value साठी 100KB मर्यादेत राहण्यासाठी
    var n = Math.ceil(s.length / CH) || 1;
    if (n > 12) return; // खूपच मोठे — Cache करू नका
    var bag = {};
    for (var i = 0; i < n; i++) bag[key + "_" + i] = s.substr(i * CH, CH);
    bag[key] = String(n);
    CacheService.getScriptCache().putAll(bag, ttl);
  } catch (e) {}
}

function buildRoster_(iyatta, tukdi) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  var out = [];
  if (sh && sh.getLastRow() > 1) {
    var rows = sh.getRange(2,1,sh.getLastRow()-1,31).getValues();
    for (var i=0;i<rows.length;i++) {
      var r = rows[i];
      if (!r[10] && !r[2]) continue;
      if (iyatta && (r[6]||"").toString().trim() !== iyatta) continue;
      if (tukdi && (r[7]||"").toString().trim() !== tukdi) continue;
      out.push({
        regNo:r[2], studentId:r[3], acYear:r[1], pen:r[8], rollNo:r[9], fullName:r[10],
        motherName:r[11], gender:r[12], dob:fmt(r[16]), contact:r[24], address:r[25], photoUrl:r[26],
        whatsappMobile:r[27]||"", alternateMobile:r[28]||"",
        category:(r[29]||"").toString().trim(), minority:(r[30]||"").toString().trim(),
        iyatta:r[6], tukdi:r[7]
      });
    }
  }
  out.sort(function(a,b){ return (parseInt(a.rollNo,10)||0) - (parseInt(b.rollNo,10)||0); });
  return out;
}

function getRosterCached_(iyatta, tukdi, noCache) {
  var key = "roster_" + rosterVer_() + "_" + encodeURIComponent(iyatta + "|" + tukdi);
  if (!noCache) {
    var hit = cacheGetBig_(key);
    if (hit) return {data: hit, cached: true};
  }
  var data = buildRoster_(iyatta, tukdi);
  cachePutBig_(key, data, ROSTER_CACHE_TTL);
  return {data: data, cached: false};
}

function doGetClassStudents(p, cb) {
  try {
    var iyatta = (p.iyatta||"").toString().trim();
    var tukdi = (p.tukdi||"").toString().trim();
    var res = getRosterCached_(iyatta, tukdi, p.nocache === "1");
    return wrap(cb, {status:"ok", data: res.data, cached: res.cached});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- Notice Board साठी वर्ग-तुकडी यादी — Students sheet मध्ये प्रत्यक्षात असलेल्या इयत्ता-तुकडी जोड्या (V19.34) -----
function doGetClassList(p, cb) {
  try {
    var clKey = "classlist_" + rosterVer_();
    var clHit = cacheGetBig_(clKey);
    if (clHit) return wrap(cb, {status:"ok", data: clHit, cached: true});
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    var out = [];
    var seen = {};
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2, 7, sh.getLastRow()-1, 2).getValues(); // कॉलम G=Class, H=Division
      for (var i=0;i<rows.length;i++) {
        var iyatta = (rows[i][0]||"").toString().trim();
        var tukdi = (rows[i][1]||"").toString().trim();
        if (!iyatta) continue;
        var key = iyatta + "|" + tukdi;
        if (seen[key]) continue;
        seen[key] = true;
        out.push({iyatta: iyatta, tukdi: tukdi});
      }
    }
    out.sort(function(a,b){
      var ka = a.iyatta+"|"+a.tukdi, kb = b.iyatta+"|"+b.tukdi;
      return ka < kb ? -1 : (ka > kb ? 1 : 0);
    });
    cachePutBig_(clKey, out, ROSTER_CACHE_TTL);
    return wrap(cb, {status:"ok", data: out});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- आज वाढदिवस असलेले विद्यार्थी — संपूर्ण शाळेतील (Super Master / Master Dashboard — V19.34) -----
function doGetTodayBirthdays(p, cb) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var today = fmt(new Date());
    var todayMD = today.slice(5); // MM-DD
    var todayYear = new Date().getFullYear();
    var sh = ss.getSheetByName("Students");
    var out = [];
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2, 1, sh.getLastRow()-1, 31).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if (!r[10] && !r[2]) continue; // रिकामी row वगळा
        var dobStr = fmt(r[16]);
        if (!dobStr || dobStr.slice(5) !== todayMD) continue;
        var birthYear = parseInt(dobStr.slice(0,4), 10);
        var age = !isNaN(birthYear) ? (todayYear - birthYear) : null;
        out.push({
          iyatta: r[6], tukdi: r[7], regNo: r[2], studentId: r[3], rollNo: r[9],
          fullName: r[10], dob: dobStr, age: age,
          whatsappMobile: r[27]||"", alternateMobile: r[28]||"", contact: r[24]||"",
          photoUrl: r[26]||""
        });
      }
    }
    out.sort(function(a,b){
      var ka = (a.iyatta||"").toString()+"|"+(a.tukdi||"").toString();
      var kb = (b.iyatta||"").toString()+"|"+(b.tukdi||"").toString();
      return ka < kb ? -1 : (ka > kb ? 1 : 0);
    });
    return wrap(cb, {status:"ok", data: out, today: today});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- Combined Teacher Dashboard (V19.35: Roster Cache + महिनावार Attendance) -----
function doGetTeacherDashboard(p, cb) {
  try {
    var iyatta = (p.iyatta||"").toString().trim();
    var tukdi = (p.tukdi||"").toString().trim();
    var date = (p.date||todayStr()).toString().trim();
    // V19.43: थोड्या वेळासाठी Cache (हजेरी Save / विद्यार्थी बदल झाल्यास आपोआप रद्द)
    var ckey = "tdash_" + encodeURIComponent(iyatta + "|" + tukdi) + "_" + date + "_" + attVer_() + "_" + rosterVer_();
    if (p.nocache !== "1") { var hit = cacheGetBig_(ckey); if (hit) { hit.cached = true; return wrap(cb, hit); } }
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var roster = getRosterCached_(iyatta, tukdi, p.nocache === "1").data;
    var totalStudents = roster.length;
    var birthdayStudents = [];
    var todayMD = date.slice(5); // MM-DD
    for (var i=0;i<roster.length;i++) {
      var st = roster[i];
      if (st.dob && st.dob.slice(5) === todayMD) {
        birthdayStudents.push({regNo:st.regNo, fullName:st.fullName, dob:st.dob, contact:st.contact});
      }
    }

    var absentList = [];
    var attRows = readAttendanceDateRows_(date);
    for (var j=0;j<attRows.length;j++) {
      var ar = attRows[j];
      if (fmt(ar[1]) !== date) continue;
      if (iyatta && (ar[2]||"").toString().trim() !== iyatta) continue;
      if (tukdi && (ar[3]||"").toString().trim() !== tukdi) continue;
      absentList.push({regNo:ar[4], studentId:ar[5], fullName:ar[6], reason:ar[7]});
    }

    var notices = [];
    var noticeSh = ss.getSheetByName("Notices");
    if (noticeSh && noticeSh.getLastRow() > 1) {
      var nLast = noticeSh.getLastRow();
      var nStart = Math.max(2, nLast - 199);               // फक्त शेवटच्या २०० सूचना — Sheet वाढली तरी वेग तोच
      var nRows = noticeSh.getRange(nStart, 1, nLast - nStart + 1, 6).getValues();
      for (var k=0;k<nRows.length;k++) {
        var nr = nRows[k];
        var tc = (nr[5]||"").toString().trim();
        if (tc && tc !== iyatta+"|"+tukdi && tc !== iyatta) continue;
        notices.push({timestamp:nr[0], date:fmt(nr[1])||nr[1], title:nr[2], message:nr[3], postedBy:nr[4]});
      }
      notices.sort(function(a,b){ return new Date(b.timestamp) - new Date(a.timestamp); });
      notices = notices.slice(0,5);
    }

    var res = {
      status:"ok",
      totalStudents: totalStudents,
      absentCount: absentList.length,
      presentCount: totalStudents - absentList.length,
      absentList: absentList,
      birthdayStudents: birthdayStudents,
      notices: notices
    };
    cachePutBig_(ckey, res, 90);
    return wrap(cb, res);
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function doGetMyLog(p, cb) {
  try {
    var username = (p.username||"").toString().trim().toLowerCase();
    var limit = parseInt(p.limit||"100",10);
    // V19.48: एकावेळी फक्त चालू + मागील महिना (Log_YYYY_MM); स्थलांतरापूर्वी जुनी "Log" सुद्धा
    var curKey = logMonthKey_(new Date());
    var rows = logLegacyTail_(5000).concat(logTailRows_(logPrevKey_(curKey), 20000), logTailRows_(curKey, 20000));
    var tmp = [];
    for (var i=0;i<rows.length;i++) {
      var r = rows[i];
      if (username && (r[1]||"").toString().trim().toLowerCase() !== username) continue;
      var dt = parseLogDate_(r[0]);
      tmp.push({seq: i, t: dt ? dt.getTime() : 0, o: {timestamp:r[0], user:r[1], role:r[2], action:r[3], reference:r[4]}});
    }
    tmp.sort(function(a,b){ return (b.t - a.t) || (b.seq - a.seq); });   // नवे आधी (तारीख, नंतर नोंदीचा क्रम)
    return wrap(cb, {status:"ok", data: tmp.slice(0, limit).map(function(x){ return x.o; })});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function wrap(cb, obj) {
  var json = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(cb + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doGetAllAction(p, cb) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    if (!sh || sh.getLastRow() < 2) return wrap(cb, {status:"ok", data:[]});
    var rows = sh.getRange(2, 1, sh.getLastRow()-1, 31).getValues();
    var all = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!r[10] && !r[2]) continue; // रिकामी row वगळा
      all.push({
        _rowIndex: i + 2,
        acYear:r[1], regNo:r[2], studentId:r[3], bookNo:r[4], aadhar:r[5],
        iyatta:r[6], tukdi:r[7], pen:r[8], rollNo:r[9],
        firstName:r[10], motherName:r[11], gender:r[12],
        religion:r[13], caste:r[14], subcaste:r[15], dob:fmt(r[16]), dobWords:r[17],
        nationality:r[18], motherTongue:r[19], birthVillage:r[20],
        prevSchool:r[21], admissionDate:fmt(r[22]), admissionClass:r[23],
        contact:r[24], address:r[25], photoUrl:r[26],
        whatsappMobile:r[27]||"", alternateMobile:r[28]||"",
        category:r[29]||"", minority:r[30]||""
      });
    }
    return wrap(cb, {status:"ok", data:all});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function doSearchAction(p, cb) {
  try {
    var q = (p.q || "").toString().trim().toLowerCase();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    if (!sh || sh.getLastRow() < 2) return wrap(cb,{status:"notfound"});
    var rows = sh.getRange(2,1,sh.getLastRow()-1,31).getValues();
    for (var i=0;i<rows.length;i++) {
      var r=rows[i];
      var cands=[
        (r[2]||"").toString().trim().toLowerCase(),
        (r[3]||"").toString().trim().toLowerCase(),
        (r[5]||"").toString().trim().toLowerCase(),
        (r[10]||"").toString().trim().toLowerCase()
      ];
      if (cands.indexOf(q)!==-1) {
        return wrap(cb,{status:"found",rowIndex:i+2,data:{
          acYear:r[1],regNo:r[2],studentId:r[3],bookNo:r[4],aadhar:r[5],
          iyatta:r[6],tukdi:r[7],pen:r[8],rollNo:r[9],
          firstName:r[10],motherName:r[11],gender:r[12],
          religion:r[13],caste:r[14],subcaste:r[15],dob:fmt(r[16]),dobWords:r[17],nationality:r[18],
          motherTongue:r[19],birthVillage:r[20],
          prevSchool:r[21],admissionDate:fmt(r[22]),admissionClass:r[23],
          contact:r[24],address:r[25],photoUrl:r[26],
          whatsappMobile:r[27]||"",alternateMobile:r[28]||"",
          category:r[29]||"",minority:r[30]||""
        }});
      }
    }
    return wrap(cb,{status:"notfound"});
  } catch(err) {
    return wrap(cb,{status:"error",message:err.toString()});
  }
}
function fmt(v){
  if(!v) return "";
  if(v instanceof Date){
    return v.getFullYear()+"-"+("0"+(v.getMonth()+1)).slice(-2)+"-"+("0"+v.getDate()).slice(-2);
  }
  return v.toString();
}

function cleanRegFileName(regNo) {
  var text = (regNo || "student").toString().trim();
  var bad = '\\/:*?"<>|#%&{}$!\'@+=~';
  var out = "";
  for (var i = 0; i < text.length; i++) {
    var ch = text.charAt(i);
    out += bad.indexOf(ch) >= 0 ? "_" : ch;
  }
  return out || "student";
}

function publicDriveUrl(fileId) {
  return "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1000";
}

function testPhotoFolderAccess() {
  try {
    var folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
    var blob = Utilities.newBlob("ok", "text/plain", "_sgs_upload_test.txt");
    var file = folder.createFile(blob);
    file.setTrashed(true);
    return {status:"ok", folderName:folder.getName(), folderId:PHOTO_FOLDER_ID};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

// V19.45: एकाच नावाच्या अनेक फाइल्स असल्यास सर्वात अलीकडे बदललेली फाइल निवडतो.
function latestFileByName_(folder, name) {
  var best = null, bestT = -1;
  var it = folder.getFilesByName(name);
  while (it.hasNext()) {
    var f = it.next();
    var t = 0;
    try { t = f.getLastUpdated().getTime(); } catch (e) {}
    if (!best || t > bestT) { best = f; bestT = t; }
  }
  return best;
}

function savePhotoData(regNo, photoData, mimeType) {
  var stage = "सुरुवात";   // V19.46.1: त्रुटी-संदेशात कोणता Drive/Sheet कॉल अडला ते दिसावे म्हणून
  try {
    regNo = (regNo || "").toString().trim();
    if (!regNo) return {status:"error", message:"regNo required"};
    if (!photoData) return {status:"error", message:"photoData required"};

    // V19.45: आकार आधी तपासा — पूर्वी जुना फोटो आधी Trash होऊन मग आकार-त्रुटी यायची (जुना फोटो वाया).
    var bytes = Utilities.base64Decode(photoData);
    if (bytes.length > 10240) return {status:"error", message:"Photo size is " + bytes.length + " bytes. Please compress below 10 KB."};

    stage = "getFolderById (Photo फोल्डर)";
    var folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
    var fileName = cleanRegFileName(regNo) + ".jpg";
    stage = "getFilesByName (जुनी फाइल शोध)";
    var oldFiles = [];
    var existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) oldFiles.push(existing.next());

    stage = "createFile (नवी फाइल)";
    var blob = Utilities.newBlob(bytes, mimeType || "image/jpeg", fileName);
    var file = folder.createFile(blob);
    stage = "setSharing (Anyone with link — View)";
    var shareWarning = "";
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (eShare) {
      // V19.47.1: फाइल-स्तरावर sharing नाकारले (उदा. Workspace धोरण) तरी फोल्डरच "Anyone" असेल तर फाइलला तोच प्रवेश मिळतो — चालू ठेवा.
      // फोल्डर खाजगी असल्यास फोटो कोणालाच दिसणार नाही, म्हणून तेव्हा पूर्वीप्रमाणेच अपयश (नवी फाइल Trash, जुनी तशीच).
      var folderPublic = false;
      try {
        var fa = folder.getSharingAccess();
        folderPublic = (fa === DriveApp.Access.ANYONE || fa === DriveApp.Access.ANYONE_WITH_LINK);
      } catch (eFa) {}
      if (!folderPublic) {
        try { file.setTrashed(true); } catch (eT) {}
        throw eShare;
      }
      shareWarning = "फाइल-स्तरावर sharing नाकारले; फोल्डरच्या sharing वर अवलंबून.";
    }
    stage = "setTrashed (जुनी फाइल)";
    oldFiles.forEach(function(f) { try { f.setTrashed(true); } catch (eOld) {} });   // नवी फाइल यशस्वी झाल्यावरच जुनी काढा

    stage = "Students मध्ये लिंक लिहिणे";
    var photoUrl = publicDriveUrl(file.getId());
    var upd = updateStudentPhotoUrl(regNo, photoUrl);
    var okRes = {status:"ok", regNo:regNo, photoUrl:photoUrl, fileId:file.getId(), studentFound:!!upd.found, sheetUpdated:!!upd.updated};
    if (shareWarning) okRes.shareWarning = shareWarning;
    return okRes;
  } catch(err) {
    return {status:"error", message:err.toString() + " [चरण: " + stage + "]"};
  }
}

function handlePhotoUpload(p) {
  return savePhotoData(p.regNo, p.photoData, p.mimeType);
}

function startPhotoChunkUpload(p) {
  try {
    if (!p.uploadId) return {status:"error", message:"uploadId required"};
    var cache = CacheService.getScriptCache();
    cache.put("photo_" + p.uploadId + "_meta", JSON.stringify({
      regNo:p.regNo || "",
      mimeType:p.mimeType || "image/jpeg",
      count:parseInt(p.count || "0", 10)
    }), 600);
    return {status:"ok", uploadId:p.uploadId};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function savePhotoChunk(p) {
  try {
    if (!p.uploadId) return {status:"error", message:"uploadId required"};
    if (p.idx === undefined) return {status:"error", message:"idx required"};
    if (p.chunk === undefined) return {status:"error", message:"chunk required"};
    CacheService.getScriptCache().put("photo_" + p.uploadId + "_" + p.idx, p.chunk, 600);
    return {status:"ok", idx:p.idx};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function finishPhotoChunkUpload(p) {
  try {
    var uploadId = p.uploadId || "";
    var count = parseInt(p.count || "0", 10);
    if (!uploadId) return {status:"error", message:"uploadId required"};
    if (!count || count < 1) return {status:"error", message:"count required"};
    var cache = CacheService.getScriptCache();
    var keys = [];
    for (var i = 0; i < count; i++) keys.push("photo_" + uploadId + "_" + i);
    var got = cache.getAll(keys);
    var parts = [];
    for (var j = 0; j < count; j++) {
      var part = got["photo_" + uploadId + "_" + j];
      if (part === null || part === undefined) return {status:"error", message:"Missing photo chunk " + (j + 1) + "/" + count};
      parts.push(part);
    }
    return savePhotoData(p.regNo, parts.join(""), p.mimeType || "image/jpeg");
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

// V19.45: Students मधील Photo URL (PHOTO_COL) अद्ययावत. फक्त बदल असेल तेव्हाच लिहितो व roster cache रद्द करतो
// (पूर्वी प्रत्येक वेळी लिहून cache रद्द व्हायचा). परत: {found, updated}
function updateStudentPhotoUrl(regNo, photoUrl) {
  var sh = getOrCreateSheet("Students", MAIN_HEADERS);
  var row = findRowByKey(sh, 3, regNo) || findRowByKey(sh, 4, regNo);
  if (!row) return {found:false, updated:false};
  var cur = (sh.getRange(row, PHOTO_COL).getValue() || "").toString();
  if (cur === String(photoUrl)) return {found:true, updated:false, same:true};
  sh.getRange(row, PHOTO_COL).setValue(photoUrl);
  SpreadsheetApp.flush();
  invalidateRosterCache_();
  return {found:true, updated:true};
}

function getPhotoUrlByRegNo(regNo) {
  try {
    regNo = (regNo || "").toString().trim();
    if (!regNo) return {status:"error", message:"regNo required"};
    var sh = getOrCreateSheet("Students", MAIN_HEADERS);
    var row = findRowByKey(sh, 3, regNo) || findRowByKey(sh, 4, regNo);
    var folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
    var file = latestFileByName_(folder, cleanRegFileName(regNo) + ".jpg");
    if (file) {
      var photoUrl = publicDriveUrl(file.getId());
      updateStudentPhotoUrl(regNo, photoUrl);
      return {status:"ok", photoUrl:photoUrl};
    }
    if (row) {
      var savedUrl = sh.getRange(row, PHOTO_COL).getValue();   // V19.45: पूर्वी स्तंभ 33 (अस्तित्वात नसलेला) वाचला जायचा
      if (savedUrl) return {status:"ok", photoUrl:savedUrl.toString()};
    }
    return {status:"notfound", photoUrl:""};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}


// =====================================================================
// ===== V19.47 — शिष्यवृत्ती नोंदवही (Diary मधील "Scholarship" प्रकारापासून स्वतंत्र) =====
// वार्षिक Sheet "Scholarship_<शै.वर्ष>". नोंद/अद्ययावत: super, master, वर्ग शिक्षक (फक्त स्वतःच्या वर्गाचे — Token मधून सक्ती).
// मिटवणे: फक्त super/master, आधी Deleted_Scholarship मध्ये Backup. नाव/वर्ग server वर विद्यार्थी-यादीतून भरले जाते.
// एकाच विद्यार्थ्याची, एकाच वर्षातील, एकाच नावाची नोंद दुबार होत नाही — अद्ययावत होते.
// =====================================================================
var SCHOLAR_HEADERS = ["Id","Timestamp","AcYear","RegNo","StudentId","FullName","Class","Division","ScholarshipName","Scheme","Amount","Status","Remark","UpdatedBy","UpdatedAt"];
var SCHOLAR_STATUSES = ["प्रलंबित","मंजूर","वितरित"];

function scholarSheet_(year, create) { return getYearSheet_("Scholarship", year, SCHOLAR_HEADERS, create, 4); }
function scholarNorm_(v) { return (v === null || v === undefined ? "" : v).toString().trim().replace(/\s+/g, " ").toLowerCase(); }
function scholarClassOf_(s) { var c = String(s || "").split("|"); return {iyatta: (c[0] || "").trim(), tukdi: (c[1] || "").trim()}; }

// भूमिका तपासणी. परत: {err} किंवा {role, teacherClass:{iyatta,tukdi}|null}
function scholarAuth_(p, forWrite) {
  if (!p._authed) return {err: "शिष्यवृत्ती नोंदवहीसाठी Token सह Login आवश्यक आहे — कृपया पुन्हा Login करा."};
  var r = p.requesterRole;
  if (r !== "super" && r !== "master" && r !== "teacher") return {err: "शिष्यवृत्ती नोंदवहीचा अधिकार Super / Master / वर्ग शिक्षकाला आहे."};
  if (r === "teacher") {
    var tc = scholarClassOf_(p._tClass);
    if (!tc.iyatta) return {err: "तुमच्या User ला वर्ग नेमलेला नाही — Master कडून Users मध्ये वर्ग भरून घ्या."};
    return {role: r, teacherClass: tc};
  }
  return {role: r, teacherClass: null};
}
function scholarFindStudent_(regNo) {
  var key = scholarNorm_(regNo);
  if (!key) return null;
  var list = getRosterCached_("", "", false).data || [];
  for (var i = 0; i < list.length; i++) if (scholarNorm_(list[i].regNo) === key) return list[i];
  return null;
}
function scholarRowToObj_(r, cur) {
  return {id: r[0], timestamp: r[1], acYear: r[2], regNo: r[3], studentId: r[4], fullName: r[5],
    iyatta: cur ? cur.iyatta : r[6], tukdi: cur ? cur.tukdi : r[7], scholarshipName: r[8], scheme: r[9],
    amount: parseFloat(r[10]) || 0, status: r[11], remark: r[12], updatedBy: r[13], updatedAt: r[14]};
}
function scholarNewId_() { return "SCH" + Date.now().toString(36) + Math.floor(Math.random() * 46656).toString(36); }

function doSaveScholarship(d) {
  var lock = null;
  try {
    var a = scholarAuth_(d, true);
    if (a.err) return {status: "error", message: a.err};
    var name = (d.scholarshipName || "").toString().trim();
    if (!name) return {status: "error", message: "शिष्यवृत्तीचे नाव आवश्यक आहे."};
    if (name.length > 120) return {status: "error", message: "शिष्यवृत्तीचे नाव खूप मोठे आहे."};
    var status = (d.status || "प्रलंबित").toString().trim();
    if (SCHOLAR_STATUSES.indexOf(status) === -1) return {status: "error", message: "स्थिती: " + SCHOLAR_STATUSES.join(" / ") + " यापैकी हवी."};
    var amtRaw = (d.amount === undefined || d.amount === null || d.amount === "") ? 0 : parseFloat(d.amount);
    if (isNaN(amtRaw) || amtRaw < 0 || amtRaw > 100000000) return {status: "error", message: "रक्कम बरोबर टाका (० किंवा जास्त)."};
    var amount = Math.round(amtRaw * 100) / 100;
    var year = normAcYear_(d.acYear);
    if (isArchivedYear_(year)) return {status: "error", message: archivedMsg_(year)};   // V19.53
    var by = d.auditUser || d.requesterUser || "";
    var ts = new Date().toLocaleString("en-IN");

    lock = LockService.getScriptLock();
    lock.waitLock(20000);
    var sh = scholarSheet_(year, !!!d.id);
    var id = (d.id || "").toString().trim();
    var rowNo = 0, old = null, regNo = (d.regNo || "").toString().trim();

    if (id) {   // Edit: Id ने ओळ शोधा; विद्यार्थी त्या ओळीतूनच ठरतो (क्लायंटच्या regNo वर विश्वास नाही)
      if (!sh) return {status: "error", message: "नोंद सापडली नाही (" + year + ")."};
      var hits = findRowsByValue_(sh, 1, id);
      if (!hits.length) return {status: "error", message: "नोंद सापडली नाही — कदाचित मिटवली गेली आहे."};
      rowNo = hits[0]; old = sh.getRange(rowNo, 1, 1, SCHOLAR_HEADERS.length).getValues()[0];
      regNo = (old[3] || "").toString().trim();
    }
    var stu = scholarFindStudent_(regNo);
    if (!stu) return {status: "error", message: "विद्यार्थी (Reg No: " + regNo + ") सापडला नाही."};
    if (a.teacherClass) {   // शिक्षक: विद्यार्थ्याचा खरा (server वरील) वर्ग == Token मधील वर्ग
      if ((stu.iyatta || "").toString().trim() !== a.teacherClass.iyatta || (stu.tukdi || "").toString().trim() !== a.teacherClass.tukdi) {
        return {status: "error", message: "आपण फक्त स्वतःच्या वर्गाच्या विद्यार्थ्यांची शिष्यवृत्ती नोंद करू शकता."};
      }
    }
    if (!sh) sh = scholarSheet_(year, true);

    // दुबार तपासणी: एकाच विद्यार्थ्याची एकाच वर्षातील एकाच नावाची नोंद (Edit मध्ये स्वतःची ओळ वगळून)
    var sameRows = readRowsAt_(sh, findRowsByValue_(sh, 4, regNo), SCHOLAR_HEADERS.length);
    var rowsList = findRowsByValue_(sh, 4, regNo);
    var dupRow = 0;
    for (var i = 0; i < sameRows.length; i++) {
      if (rowsList[i] === rowNo) continue;
      if (scholarNorm_(sameRows[i][8]) === scholarNorm_(name)) { dupRow = rowsList[i]; break; }
    }
    var mode = "created";
    if (dupRow && !id) { rowNo = dupRow; old = sh.getRange(rowNo, 1, 1, SCHOLAR_HEADERS.length).getValues()[0]; id = (old[0] || "").toString(); mode = "updated"; }
    else if (dupRow && id) return {status: "error", message: "या विद्यार्थ्याची '" + name + "' नावाची नोंद या वर्षात आधीच आहे."};
    else if (id) mode = "updated";

    var row = [id || scholarNewId_(), old ? old[1] : ts, year, regNo, stu.studentId || "", stu.fullName || "",
      (stu.iyatta || "").toString(), (stu.tukdi || "").toString(), name, (d.scheme || "").toString().trim().slice(0, 150), amount, status,
      (d.remark || "").toString().trim().slice(0, 500), by, ts];
    if (rowNo) {
      sh.getRange(rowNo, 1, 1, row.length).setValues([row]);
    } else {
      rowNo = sh.getLastRow() + 1;
      ensureRows_(sh, rowNo);
      sh.getRange(rowNo, 1, 1, row.length).setValues([row]);
    }
    return {rowIndex: rowNo, mode: mode, ref: regNo + " " + name, extra: {id: row[0], acYear: year, fullName: row[5], status: status, amount: amount}};
  } catch (err) {
    return {status: "error", message: err.toString()};
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (eL) {} }
  }
}

function doGetScholarship(p, cb) {
  try {
    var a = scholarAuth_(p, false);
    if (a.err) return wrap(cb, {status: "error", message: a.err});
    var year = p.year ? normAcYear_(p.year) : currentAcYear_();
    var iyatta = (p.iyatta || "").toString().trim(), tukdi = (p.tukdi || "").toString().trim();
    if (a.teacherClass) { iyatta = a.teacherClass.iyatta; tukdi = a.teacherClass.tukdi; }   // शिक्षकाचा वर्ग सक्तीचा
    var regNo = scholarNorm_(p.regNo);
    var sh = scholarSheet_(year, false);
    var out = [];
    if (sh && sh.getLastRow() > 1) {
      var classSet = null;
      if (iyatta || tukdi) {   // वर्ग-बढती/तुकडी बदलानंतरही अचूक: विद्यार्थ्यांची सध्याची वर्ग-यादी वापरा
        classSet = {};
        getRosterCached_(iyatta, tukdi, false).data.forEach(function(st) { classSet[scholarNorm_(st.regNo)] = st; });
      }
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, SCHOLAR_HEADERS.length).getValues();
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i], key = scholarNorm_(r[3]);
        if (!key) continue;
        if (regNo && key !== regNo) continue;
        if (classSet && !classSet[key]) continue;
        out.push(scholarRowToObj_(r, classSet ? classSet[key] : null));
      }
    }
    out.sort(function(x, y) { return (x.iyatta + x.tukdi + x.fullName) < (y.iyatta + y.tukdi + y.fullName) ? -1 : 1; });
    var years = listYearSheets_("Scholarship").map(function(s2) { return s2.getName().replace("Scholarship_", ""); });
    var cur = currentAcYear_(); if (years.indexOf(cur) === -1) years.unshift(cur);
    var archYears = archivedYears_();   // V19.53: Archive झालेल्या वर्षांच्या नोंदीही पाहता येतात (फक्त वाचन)
    archYears.forEach(function(ay) { if (years.indexOf(ay) === -1) years.push(ay); });
    years.sort().reverse();
    var isArch = isArchivedYear_(year);
    var resS = {status: "ok", year: year, data: out, years: years, statuses: SCHOLAR_STATUSES, canDelete: !isArch && (a.role === "super" || a.role === "master")};
    if (archYears.length) resS.archivedYears = archYears;
    if (isArch) resS.archived = true;
    return wrap(cb, resS);
  } catch (err) {
    return wrap(cb, {status: "error", message: err.toString()});
  }
}

function doDeleteScholarship(p, cb) {
  var lock = null;
  try {
    if (!p._authed || (p.requesterRole !== "super" && p.requesterRole !== "master")) {
      return wrap(cb, {status: "error", message: "शिष्यवृत्ती नोंद मिटवण्याचा अधिकार फक्त Super / Master ला आहे (Token सह Login आवश्यक)."});
    }
    var id = (p.id || "").toString().trim();
    if (!id) return wrap(cb, {status: "error", message: "Id आवश्यक आहे."});
    var year = p.year ? normAcYear_(p.year) : currentAcYear_();
    if (isArchivedYear_(year)) return wrap(cb, {status: "error", message: archivedMsg_(year)});   // V19.53
    lock = LockService.getScriptLock();
    lock.waitLock(20000);
    var sh = scholarSheet_(year, false);
    if (!sh) return wrap(cb, {status: "error", message: "नोंद सापडली नाही (" + year + ")."});
    var hits = findRowsByValue_(sh, 1, id);
    if (!hits.length) return wrap(cb, {status: "error", message: "नोंद सापडली नाही — कदाचित आधीच मिटवली आहे."});
    var rowNo = hits[0];
    var old = sh.getRange(rowNo, 1, 1, SCHOLAR_HEADERS.length).getValues()[0];
    var by = p.auditUser || p.requesterUser || "";
    var okBk = backupDeleted_("Scholarship", SCHOLAR_HEADERS, [old], by, "शिष्यवृत्ती नोंद मिटवली");
    if (!okBk) {   // Backup फाइल न जमल्यास मुख्य फाइलमधील Deleted_Scholarship मध्ये प्रत असल्याची खात्री करा; नसल्यास मिटवू नका
      var main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Deleted_Scholarship");
      var found = main && main.getLastRow() > 1 && findRowsByValue_(main, 4, id).length > 0;
      if (!found) return wrap(cb, {status: "error", message: "Backup जतन झाला नाही म्हणून नोंद मिटवली नाही. पुन्हा प्रयत्न करा."});
    }
    sh.deleteRows(rowNo, 1);
    logAudit(by, p.auditRole || p.requesterRole, "deleteScholarship", (old[3] || "") + " " + (old[8] || ""));
    return wrap(cb, {status: "ok", deleted: id});
  } catch (err) {
    return wrap(cb, {status: "error", message: err.toString()});
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (eL) {} }
  }
}


// ===== V19.48 — एकदाचे स्थलांतर: Log + Log_Archive → Log_YYYY_MM (+ Log_Unparsed) =====
// • फक्त Super, पुष्टी-शब्द "LOG"; Script Lock (स्वतः घेतो — दुसऱ्या Lock आत बोलावू नका).
// • पुन्हा चालवणे सुरक्षित: Sheet मध्ये आधी असलेल्या सारख्या ओळी (multiset) पुन्हा लिहिल्या जात नाहीत; वेळ-बजेट संपल्यास done=false.
// • सर्व ओळी तपासल्यावरच जुन्या Sheets "Log_OLD_..." नावाने लपवल्या जातात (मिटवत नाही) व flag LOG_MIGRATED=1 बसतो.
function logSig_(r) { return [r[0], r[1], r[2], r[3], r[4]].join("\u0001"); }
function logNormRow_(r) {
  var ts = r[0];
  var raw = ts;
  if (ts instanceof Date) ts = ts.toLocaleString("en-IN");
  return {row: [ts, r[1], r[2], r[3], r[4]], raw: raw};
}
function doMigrateLogs(p, cb) {
  var lock = null;
  try {
    if (!p._authed || p.requesterRole !== "super") return wrap(cb, {status:"error", message:"Log स्थलांतराचा अधिकार फक्त Super Master ला आहे (Token सह Login आवश्यक)."});
    if (String(p.confirm || "") !== "LOG") return wrap(cb, {status:"error", message:"पुष्टीसाठी LOG टाइप करा."});
    if (logMigrated_()) return wrap(cb, {status:"ok", done:true, already:true, message:"स्थलांतर आधीच झाले आहे."});
    lock = LockService.getScriptLock();
    lock.waitLock(30000);
    if (logMigrated_()) return wrap(cb, {status:"ok", done:true, already:true, message:"स्थलांतर आधीच झाले आहे."});
    var t0 = Date.now();
    var budget = parseInt(p.budgetMs || "", 10);
    if (isNaN(budget)) budget = 240000;
    budget = Math.max(3000, Math.min(budget, 270000));
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sources = [];
    ["Log", "Log_Archive"].forEach(function(n) { var sh = ss.getSheetByName(n); if (sh) sources.push(sh); });

    // 1) वाचा व महिन्यानुसार गट करा
    var byKey = {}, unparsed = [], total = 0;
    sources.forEach(function(sh) {
      var n = sh.getLastRow();
      if (n < 2) return;
      sh.getRange(2, 1, n - 1, 5).getValues().forEach(function(r) {
        if (!r[0] && !r[1] && !r[2] && !r[3] && !r[4]) return;
        var nr = logNormRow_(r), dt = parseLogDate_(nr.raw);
        total++;
        if (!dt || isNaN(dt.getTime())) { unparsed.push(nr.row); return; }
        var key = Utilities.formatDate(dt, "Asia/Kolkata", "yyyy_MM");
        (byKey[key] = byKey[key] || []).push(nr.row);
      });
    });

    // 2) लिहा (आधीच असलेल्या सारख्या ओळी वगळून)
    function writeInto(sh, rows) {
      var have = {}, last = sh.getLastRow();
      if (last > 1) sh.getRange(2, 1, last - 1, 5).getValues().forEach(function(r) { var g = logSig_(logNormRow_(r).row); have[g] = (have[g] || 0) + 1; });
      var todo = [];
      rows.forEach(function(r) { var g = logSig_(r); if (have[g] > 0) have[g]--; else todo.push(r); });
      if (!todo.length) return 0;
      var start = Math.max(sh.getLastRow(), 1) + 1;
      ensureRows_(sh, start + todo.length - 1);
      sh.getRange(start, 1, todo.length, 5).setValues(todo);
      return todo.length;
    }
    var keys = Object.keys(byKey).sort(), written = 0, doneKeys = 0;
    for (var i = 0; i < keys.length; i++) {
      if (Date.now() - t0 > budget) break;
      written += writeInto(logSheet_(keys[i], true), byKey[keys[i]]);
      doneKeys++;
    }
    if (doneKeys < keys.length) {
      return wrap(cb, {status:"ok", done:false, remainingMonths:keys.length - doneKeys, message:"वेळ-बजेट संपले — पुन्हा चालवा (सुरक्षित)."});
    }
    var unSh = null;
    if (unparsed.length) {
      unSh = ss.getSheetByName("Log_Unparsed");
      if (!unSh) { unSh = ss.insertSheet("Log_Unparsed"); trimNewSheet_(unSh, 5); unSh.appendRow(LOG_HEADERS); }
      written += writeInto(unSh, unparsed);
    }

    // 3) पडताळणी: स्त्रोतातील प्रत्येक ओळ लक्ष्य Sheet मध्ये आहे का
    function covered(sh, rows) {
      var have = {}, last = sh.getLastRow();
      if (last > 1) sh.getRange(2, 1, last - 1, 5).getValues().forEach(function(r) { var g = logSig_(logNormRow_(r).row); have[g] = (have[g] || 0) + 1; });
      for (var j = 0; j < rows.length; j++) { var g2 = logSig_(rows[j]); if (!(have[g2] > 0)) return false; have[g2]--; }
      return true;
    }
    for (var k = 0; k < keys.length; k++) {
      if (!covered(logSheet_(keys[k], false), byKey[keys[k]])) return wrap(cb, {status:"error", message:"पडताळणी अयशस्वी (" + keys[k] + ") — जुने Sheets तसेच ठेवले; पुन्हा चालवा."});
    }
    if (unparsed.length && !covered(unSh, unparsed)) return wrap(cb, {status:"error", message:"पडताळणी अयशस्वी (Log_Unparsed) — जुने Sheets तसेच ठेवले."});

    // 4) जुन्या Sheets लपवा (मिटवू नका) व flag बसवा
    var stamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyyMMdd"), hidden = [];
    sources.forEach(function(sh) {
      var base = "Log_OLD_" + (sh.getName() === "Log_Archive" ? "Archive_" : "") + stamp, nm = base, c = 1;
      while (ss.getSheetByName(nm)) { c++; nm = base + "_" + c; }
      sh.setName(nm);
      try { sh.hideSheet(); } catch (eH) {}
      hidden.push(nm);
    });
    PropertiesService.getScriptProperties().setProperty("LOG_MIGRATED", "1");
    logAudit(p.auditUser || p.requesterUser, p.auditRole || p.requesterRole, "migrateLogs", "rows=" + total + ", months=" + keys.length + ", unparsed=" + unparsed.length);
    return wrap(cb, {status:"ok", done:true, report:{rows: total, written: written, months: keys, unparsed: unparsed.length, hiddenSheets: hidden}});
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (eR) {} }
  }
}


// =====================================================================
// ===== V19.46 — Photo link recall (फक्त Super) =====
// PHOTO_FOLDER मधील फोटो फाइल्सच्या नावावरून (विस्तारशिवाय == cleanRegFileName(RegNo)) Students स्तंभ PHOTO_COL (27) मध्ये योग्य लिंक भरतो.
// • गटांमध्ये चालतो: प्रत्येक विनंती ~४ मिनिटांचा वेळ-बजेट वापरते व continuation token परत देते; क्लायंट "पुढील गट" आपोआप मागवतो.
// • पुन्हा चालवणे सुरक्षित (idempotent): आधीच बरोबर लिंक पुन्हा लिहिली जात नाही.
// • dry=1 (डीफॉल्ट): काहीही लिहीत नाही, फक्त अहवाल. फाइल sharing/Trash/मूळ डेटा बदलत नाही; फक्त Photo URL स्तंभ.
// • सत्र-स्थिती (token, आकडे, दावे) Script Cache मध्ये ६ तास (runId नुसार); cache गेल्यास "पुन्हा सुरू करा" संदेश — डेटाला धोका नाही.
// =====================================================================
var PHOTO_RECALL_BUDGET_MS = 240000;
var PHOTO_RECALL_SHARDS = 16;
var PHOTO_RECALL_TTL = 21600;

function photoRecallShard_(key) {
  var h = 0;
  for (var i = 0; i < key.length; i++) h = (h * 31 + key.charCodeAt(i)) % 1000003;
  return h % PHOTO_RECALL_SHARDS;
}
function photoIdFromUrl_(url) {
  var m = /[?&]id=([\w-]+)|\/d\/([\w-]+)/.exec((url || "").toString());
  return m ? (m[1] || m[2]) : "";
}
function photoBaseName_(name) {
  name = (name || "").toString();
  var i = name.lastIndexOf(".");
  return i > 0 ? name.substring(0, i) : name;
}

function doPhotoRecall(p, cb) {
  var lock = null;
  try {
    if (!p._authed || p.requesterRole !== "super") {
      return wrap(cb, {status:"error", message:"Photo link recall चा अधिकार फक्त Super Master ला आहे (Token सह Login आवश्यक)."});
    }
    var dry = (p.dry === undefined || p.dry === null || p.dry === "") ? true : (String(p.dry) !== "0");
    var budget = parseInt(p.budgetMs || "", 10);
    if (isNaN(budget)) budget = PHOTO_RECALL_BUDGET_MS;
    budget = Math.max(3000, Math.min(budget, 270000));
    var t0 = Date.now();
    var cache = CacheService.getScriptCache();

    // ---- सत्र-स्थिती ----
    var runId = (p.runId || "").toString().replace(/[^\w-]/g, "");
    var st;
    if (!runId) {
      runId = Utilities.getUuid().replace(/[^\w]/g, "").substring(0, 12);
      st = {dry: dry, token: "", files: 0, nonImage: 0, orphans: 0, orphanSample: [], batches: 0};
    } else {
      var raw = cache.get("prs_" + runId);
      if (!raw) return wrap(cb, {status:"error", code:"runlost", message:"सत्र-स्थिती संपली (cache). कृपया पुन्हा सुरू करा — आधी झालेले बदल सुरक्षित आहेत व पुन्हा चालवणे सुरक्षित आहे."});
      st = JSON.parse(raw);
      dry = !!st.dry;   // सत्राचा मोड बदलता येत नाही
    }

    if (!dry) {
      lock = LockService.getScriptLock();
      try { lock.waitLock(20000); } catch (eL) { lock = null; return wrap(cb, {status:"error", message:"दुसरी प्रक्रिया चालू आहे — थोड्या वेळाने पुन्हा प्रयत्न करा."}); }
    }

    // ---- विद्यार्थी: key (cleanRegFileName(RegNo)) → ओळ, सध्याची लिंक ----
    var sh = getOrCreateSheet("Students", MAIN_HEADERS);
    var last = sh.getLastRow();
    var regVals = last >= 2 ? sh.getRange(2, 3, last - 1, 1).getValues() : [];
    var urlVals = last >= 2 ? sh.getRange(2, PHOTO_COL, last - 1, 1).getValues() : [];
    var byKey = {}, dupRegs = 0;
    for (var i = 0; i < regVals.length; i++) {
      var rg = (regVals[i][0] || "").toString().trim();
      if (!rg) continue;
      var k = cleanRegFileName(rg);
      if (byKey[k]) { dupRegs++; continue; }
      byKey[k] = {row: i + 2, reg: rg, url: (urlVals[i][0] || "").toString()};
    }

    // ---- दावे (claims) shards: key → {i: fileId, t: modified ms, o: मूळ लिंकचा id ("" रिकामी, "x" अनोळखी स्वरूप)} ----
    var shardKeys = [];
    for (var sI = 0; sI < PHOTO_RECALL_SHARDS; sI++) shardKeys.push("prc_" + runId + "_" + sI);
    var got = cache.getAll(shardKeys);
    var shards = [], dirty = {};
    for (var sJ = 0; sJ < PHOTO_RECALL_SHARDS; sJ++) shards.push(got[shardKeys[sJ]] ? JSON.parse(got[shardKeys[sJ]]) : {});

    var folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
    var it = st.token ? DriveApp.continueFileIterator(st.token) : folder.getFiles();
    var changed = {};    // row → url (या गटात लिहायचे)
    var processed = 0, finished = false;
    while (true) {
      if (!it.hasNext()) { finished = true; break; }
      if (Date.now() - t0 > budget) break;
      var f = it.next();
      processed++; st.files++;
      var mt = "image/jpeg";
      try { mt = f.getMimeType ? f.getMimeType() : mt; } catch (eM) {}
      if ((mt || "").toString().indexOf("image/") !== 0) { st.nonImage++; continue; }
      var key = photoBaseName_(f.getName());
      var stu = byKey[key];
      if (!stu) {
        st.orphans++;
        if (st.orphanSample.length < 30) st.orphanSample.push(f.getName());
        continue;
      }
      var fid = f.getId(), ft = 0;
      try { ft = f.getLastUpdated().getTime(); } catch (eT) {}
      var sd = shards[photoRecallShard_(key)];
      var c = sd[key];
      if (c && ft <= c.t) continue;           // एकाच नावाची जुनी/समान फाइल — सर्वात अलीकडील आधीच निवडली
      var origId = c ? c.o : (stu.url ? (photoIdFromUrl_(stu.url) || "x") : "");
      sd[key] = {i: fid, t: ft, o: origId};
      dirty[photoRecallShard_(key)] = 1;
      var want = publicDriveUrl(fid);
      if (stu.url !== want) changed[stu.row] = want;
      else delete changed[stu.row];
    }

    // ---- लेखन: एकाच setValues (फक्त बदललेल्या ओळींच्या पल्ल्यात; लेखनापूर्वी ताजे वाचन) ----
    var wrote = 0;
    var rows = Object.keys(changed).map(function(x) { return parseInt(x, 10); });
    if (!dry && rows.length) {
      var minR = Math.min.apply(null, rows), maxR = Math.max.apply(null, rows);
      var fresh = sh.getRange(minR, PHOTO_COL, maxR - minR + 1, 1).getValues();
      rows.forEach(function(r) { fresh[r - minR][0] = changed[r]; wrote++; });
      sh.getRange(minR, PHOTO_COL, maxR - minR + 1, 1).setValues(fresh);
      SpreadsheetApp.flush();
      invalidateRosterCache_();
    }

    // ---- स्थिती जतन ----
    st.batches++;
    st.token = finished ? "" : it.getContinuationToken();
    var putAll = {};
    Object.keys(dirty).forEach(function(d) { putAll[shardKeys[d]] = JSON.stringify(shards[d]); });
    putAll["prs_" + runId] = JSON.stringify(st);
    Object.keys(putAll).forEach(function(k2) { cache.put(k2, putAll[k2], PHOTO_RECALL_TTL); });

    var res = {status:"ok", runId: runId, dry: dry, done: finished, batchFiles: processed, batchWrote: wrote,
      progress: {files: st.files, batches: st.batches}};

    if (finished) {
      var matched = 0, updated = 0, same = 0, claimed = {};
      shards.forEach(function(sd2) {
        Object.keys(sd2).forEach(function(k3) { matched++; claimed[k3] = 1; if (sd2[k3].i === sd2[k3].o) same++; else updated++; });
      });
      var noPhoto = [], noPhotoCount = 0;
      Object.keys(byKey).forEach(function(k4) {
        if (!claimed[k4]) { noPhotoCount++; if (noPhoto.length < 30) noPhoto.push(byKey[k4].reg); }
      });
      res.report = {files: st.files, nonImage: st.nonImage, matched: matched, updated: updated, alreadyCorrect: same,
        studentsWithoutPhoto: noPhotoCount, studentsWithoutPhotoSample: noPhoto,
        filesWithoutStudent: st.orphans, filesWithoutStudentSample: st.orphanSample, duplicateRegNos: dupRegs,
        dry: dry, wouldUpdate: dry ? updated : 0};
      // cache स्वच्छ करा
      shardKeys.forEach(function(k5) { cache.remove(k5); }); cache.remove("prs_" + runId);
      if (!dry) logAudit(p.auditUser || p.requesterUser, p.auditRole || p.requesterRole, "photoRecall", "files=" + st.files + ", updated=" + updated + ", ok=" + same);
    }
    return wrap(cb, res);
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (eR) {} }
  }
}


// =====================================================================
// ===== V19.35 — नवीन Admin सुविधा (Backup / थकबाकी यादी / वर्ग-बढती / हजेरी देखरेख) =====
// =====================================================================

// ---------- 1. आपोआप Backup (Time-driven Trigger) ----------
var BACKUP_KEEP_COPIES = 14;          // प्रत्येक Spreadsheet च्या शेवटच्या किती प्रती ठेवायच्या
var BACKUP_FOLDER_NAME = "SGS_Auto_Backups";

function getBackupFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("BACKUP_FOLDER_ID");
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* Folder delete झाला असल्यास नवीन बनवा */ }
  }
  var folder = DriveApp.createFolder(BACKUP_FOLDER_NAME);
  props.setProperty("BACKUP_FOLDER_ID", folder.getId());
  return folder;
}

// Spreadsheet IDs: चालू (Students/Users इ.) + Script Property "EXTRA_BACKUP_SS_IDS" मधील (comma-separated) इतर Spreadsheets
function getBackupSpreadsheetIds_() {
  var ids = [];
  try { ids.push(SpreadsheetApp.getActiveSpreadsheet().getId()); } catch (e) {}
  var extra = PropertiesService.getScriptProperties().getProperty("EXTRA_BACKUP_SS_IDS") || "";
  extra.split(",").forEach(function(x) {
    x = x.trim();
    if (x && ids.indexOf(x) === -1) ids.push(x);
  });
  return ids;
}

// Trigger या function ला call करतो. Editor मधून manually Run केले तरी चालते.
function backupSpreadsheets() {
  var props = PropertiesService.getScriptProperties();
  var stamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd_HHmm");
  var result = {time: new Date().toISOString(), stamp: stamp, files: [], errors: []};
  try {
    var folder = getBackupFolder_();
    getBackupSpreadsheetIds_().forEach(function(id) {
      try {
        var file = DriveApp.getFileById(id);
        var baseName = file.getName();
        var copy = file.makeCopy(baseName + "_BACKUP_" + stamp, folder);
        result.files.push({name: copy.getName(), url: copy.getUrl()});
        // जुन्या प्रती काढा — फक्त शेवटच्या BACKUP_KEEP_COPIES ठेवा
        var prefix = baseName + "_BACKUP_";
        var all = [];
        var it = folder.getFiles();
        while (it.hasNext()) {
          var f = it.next();
          if (f.getName().indexOf(prefix) === 0) all.push(f);
        }
        all.sort(function(a, b) { return a.getName() < b.getName() ? 1 : (a.getName() > b.getName() ? -1 : 0); });
        for (var i = BACKUP_KEEP_COPIES; i < all.length; i++) all[i].setTrashed(true);
      } catch (e1) {
        result.errors.push(id + ": " + e1.toString());
      }
    });
  } catch (e) {
    result.errors.push(e.toString());
  }
  result.ok = result.errors.length === 0 && result.files.length > 0;
  if (result.ok) {
    // Backup यशस्वी झाल्यावरच जुन्या Log नोंदी (९० दिवसांपेक्षा जुन्या) Archive करा
    try { result.logArchived = archiveOldLog_(90); } catch (eLg) { result.logArchived = -1; }
  } else {
    notifyBackupFailure_(result);
  }
  props.setProperty("LAST_BACKUP", JSON.stringify(result));
  logAudit("system", "trigger", "autoBackup", result.ok ? (result.files.length + " प्रती") : ("त्रुटी: " + result.errors.join(" | ")));
  return result;
}

function removeBackupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "backupSpreadsheets") ScriptApp.deleteTrigger(t);
  });
  PropertiesService.getScriptProperties().setProperty("BACKUP_SCHEDULE", "off");
}
// Editor मधून एकदा Run करा → रोज पहाटे ~२ वा. Backup
function setupDailyBackupTrigger() {
  removeBackupTriggers();
  ScriptApp.newTrigger("backupSpreadsheets").timeBased().everyDays(1).atHour(2).create();
  PropertiesService.getScriptProperties().setProperty("BACKUP_SCHEDULE", "daily");
}
// किंवा दर रविवारी पहाटे ~२ वा.
function setupWeeklyBackupTrigger() {
  removeBackupTriggers();
  ScriptApp.newTrigger("backupSpreadsheets").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(2).create();
  PropertiesService.getScriptProperties().setProperty("BACKUP_SCHEDULE", "weekly");
}
// दुसऱ्या Spreadsheet(s) चे ID Backup यादीत जोडण्यासाठी Editor मधून Run करा: setExtraBackupSpreadsheetIds("ID1,ID2")
function setExtraBackupSpreadsheetIds(csv) {
  PropertiesService.getScriptProperties().setProperty("EXTRA_BACKUP_SS_IDS", (csv || "").toString());
}

function doRunBackup(d) {
  if (d.requesterRole !== "super") return {status:"error", message:"Backup चा अधिकार फक्त Super Master ला आहे."};
  var res = backupSpreadsheets();
  if (!res.ok) return {status:"error", message:"Backup अपूर्ण: " + res.errors.join(" | ")};
  return {rowIndex:0, mode:"created", count: res.files.length, ref:"backup", extra: res};
}

function notifyBackupFailure_(result) {
  try {
    var to = PropertiesService.getScriptProperties().getProperty("BACKUP_ALERT_EMAIL") || "";
    if (!to) { try { to = Session.getEffectiveUser().getEmail(); } catch (e0) {} }
    if (!to) return;
    MailApp.sendEmail(to, "⚠️ SGS School System — Backup अयशस्वी",
      "स्वयंचलित Backup अयशस्वी झाला.\n\nवेळ: " + result.time + "\nत्रुटी:\n" + (result.errors || []).join("\n") +
      "\n\nकृपया Apps Script Editor मध्ये backupSpreadsheets एकदा Run करून परवानग्या/जागा तपासा.");
  } catch (e) { /* Mail पाठवता न आल्यास Backup status मध्ये त्रुटी दिसतेच */ }
}

function doSetBackupSchedule(d) {
  if (d.requesterRole !== "super") return {status:"error", message:"Backup चा अधिकार फक्त Super Master ला आहे."};
  try {
    if (typeof d.alertEmail === "string") {
      var em = d.alertEmail.trim();
      if (em && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em)) return {status:"error", message:"Email चुकीचा आहे."};
      PropertiesService.getScriptProperties().setProperty("BACKUP_ALERT_EMAIL", em);
    }
    if (typeof d.extraIds === "string") {
      var cleaned = d.extraIds.split(",").map(function(x){ return x.trim(); }).filter(function(x){ return !!x; }).join(",");
      setExtraBackupSpreadsheetIds(cleaned);
    }
    var mode = (d.mode || "").toString();
    if (mode === "daily") setupDailyBackupTrigger();
    else if (mode === "weekly") setupWeeklyBackupTrigger();
    else if (mode === "off") removeBackupTriggers();
    else if (mode) return {status:"error", message:"अपेक्षित mode: daily / weekly / off"};
    return {rowIndex:0, mode:"updated", ref:"backupSchedule:" + mode};
  } catch (err) {
    return {status:"error", message:"Trigger सेट करता आला नाही — Apps Script Editor मध्ये setupDailyBackupTrigger एकदा Run करून परवानगी द्या. (" + err.toString() + ")"};
  }
}

function doGetBackupStatus(p, cb) {
  try {
    if (p.requesterRole !== "super") return wrap(cb, {status:"error", message:"अधिकार फक्त Super Master ला आहे."});
    var props = PropertiesService.getScriptProperties();
    var last = null;
    try { last = JSON.parse(props.getProperty("LAST_BACKUP") || "null"); } catch (e) {}
    return wrap(cb, {status:"ok", last: last, schedule: props.getProperty("BACKUP_SCHEDULE") || "off",
      extraIds: props.getProperty("EXTRA_BACKUP_SS_IDS") || "", keep: BACKUP_KEEP_COPIES,
      alertEmail: props.getProperty("BACKUP_ALERT_EMAIL") || ""});
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ---------- 2. शाळा-व्यापी थकबाकी (बाकी-फी) यादी — फक्त Super Master ----------
var CLASS_ORDER_SRV = ["5th","6th","7th","8th","9th","10th"];
function classSortKey_(c) { var i = CLASS_ORDER_SRV.indexOf((c || "").toString().trim()); return i === -1 ? 99 : i; }

function doGetPendingFeesAll(p, cb) {
  try {
    if (p.requesterRole !== "super") return wrap(cb, {status:"error", message:"ही यादी पाहण्याचा अधिकार फक्त Super Master ला आहे."});
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var totalFee = getTotalFeePerStudent();
    var students = getRosterCached_("", "", p.nocache === "1").data;

    // V19.42: चालू (किंवा निवडलेल्या) वर्षाच्या FeesSummary मधून — प्रत्येक विद्यार्थ्याची एक ओळ
    feesEnsureMigrated_();
    var feeMap = {};
    feesSummaryRows_(normAcYear_(p.acYear)).forEach(function(fr) {
      var rg = (fr[0] || "").toString().trim().toLowerCase();
      if (rg) feeMap[rg] = {paid: parseFloat(fr[6]) || 0, date: (fr[7] || "").toString(), opening: parseFloat(fr[11]) || 0};
    });

    var out = [], collectedAll = 0, pendingTotal = 0;
    for (var k = 0; k < students.length; k++) {
      var st = students[k];
      var key = (st.regNo || "").toString().trim().toLowerCase();
      var chosen = feeMap[key] || null;
      var paid = chosen ? chosen.paid : 0, opening = chosen ? chosen.opening : 0;
      collectedAll += paid;
      var pending = Math.max(0, opening + totalFee - paid);
      if (pending <= 0) continue;
      pendingTotal += pending;
      out.push({
        regNo: st.regNo, fullName: st.fullName, iyatta: st.iyatta, tukdi: st.tukdi, rollNo: st.rollNo,
        acYear: st.acYear, paid: paid, opening: opening, pending: pending, hasEntry: !!chosen,
        lastPaidDate: chosen ? chosen.date : "",
        mobile: (st.whatsappMobile || st.alternateMobile || st.contact || "").toString()
      });
    }
    out.sort(function(a, b) {
      if (b.pending !== a.pending) return b.pending - a.pending;
      var ca = classSortKey_(a.iyatta), cb2 = classSortKey_(b.iyatta);
      if (ca !== cb2) return ca - cb2;
      return (a.fullName || "").toString().localeCompare((b.fullName || "").toString());
    });
    return wrap(cb, {status:"ok", data: out, totalFee: totalFee,
      summary: {totalStudents: students.length, pendingCount: out.length, pendingTotal: pendingTotal, collectedTotal: collectedAll}});
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ---------- 3. वर्ग-बढती (Promotion) Tool — फक्त Super Master ----------
var PROMO_LOG_HEADERS = ["Timestamp","BatchId","RegNo","FullName","OldClass","OldDivision","OldAcYear","NewClass","NewAcYear","By","Undone"];

function doPreviewPromotion(p, cb) {
  try {
    if (p.requesterRole !== "super") return wrap(cb, {status:"error", message:"वर्ग-बढतीचा अधिकार फक्त Super Master ला आहे."});
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("Students");
    var counts = {};
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (!r[10] && !r[2]) continue;
        var c = (r[6] || "").toString().trim();
        if (!c) continue;
        counts[c] = (counts[c] || 0) + 1;
      }
    }
    var classes = Object.keys(counts).sort(function(a, b) { return classSortKey_(a) - classSortKey_(b) || (a < b ? -1 : 1); })
      .map(function(c) {
        var idx = CLASS_ORDER_SRV.indexOf(c);
        var next = (idx !== -1 && idx < CLASS_ORDER_SRV.length - 1) ? CLASS_ORDER_SRV[idx + 1] : "";
        return {iyatta: c, count: counts[c], suggestedNext: next};
      });

    var batches = {};
    var logSh = ss.getSheetByName("PromotionLog");
    if (logSh && logSh.getLastRow() > 1) {
      var lr = logSh.getRange(2, 1, logSh.getLastRow() - 1, 11).getValues();
      for (var j = 0; j < lr.length; j++) {
        if ((lr[j][10] || "").toString() === "Y") continue;
        var b = lr[j][1];
        if (!b) continue;
        if (!batches[b]) batches[b] = {batchId: b, time: lr[j][0], count: 0, by: lr[j][9]};
        batches[b].count++;
      }
    }
    var recent = Object.keys(batches).map(function(k) { return batches[k]; }).slice(-5).reverse();
    return wrap(cb, {status:"ok", classes: classes, recentBatches: recent, classOrder: CLASS_ORDER_SRV, archiveBatches: listArchiveBatches_()});
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function doPromoteStudents(d) {
  if (d.requesterRole !== "super") return {status:"error", message:"वर्ग-बढतीचा अधिकार फक्त Super Master ला आहे."};
  if ((d.confirm || "") !== "PROMOTE") return {status:"error", message:"खात्रीसाठी PROMOTE टाइप करणे आवश्यक आहे."};
  var map = {};
  try { map = JSON.parse(d.mapJson || "{}"); } catch (e) { return {status:"error", message:"वर्ग-नकाशा वाचता आला नाही."}; }
  var newAcYear = (d.newAcYear || "").toString().trim();
  if (!newAcYear) return {status:"error", message:"नवीन शैक्षणिक वर्ष (उदा. 2026-27) आवश्यक आहे."};
  var keys = Object.keys(map).filter(function(k) { return map[k] && map[k] !== k; });
  if (!keys.length) return {status:"error", message:"बढती द्यायचा किमान एक वर्ग निवडा."};

  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(30000); locked = true;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("Students");
    if (!sh || sh.getLastRow() < 2) return {status:"error", message:"Students sheet रिकामी आहे."};
    var n = sh.getLastRow() - 1;

    // सुरक्षेसाठी बढतीपूर्वीची Students sheet ची प्रत (लपवलेली) याच Spreadsheet मध्ये ठेवा
    var stamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyyMMdd_HHmm");
    var bkName = "Students_BeforePromotion_" + stamp;
    var bk = makeStudentsBackup_(sh, bkName, "Students_BeforePromotion_");   // V19.40: SGS_Deleted_Backup फाइलमध्ये (मुख्य फाइल हलकी)

    var base = sh.getRange(2, 1, n, 11).getValues();             // A..K
    var accYearCol = base.map(function(r) { return [r[1]]; });    // B
    var classDiv = base.map(function(r) { return [r[6], r[7]]; }); // G:H
    var batchId = "P" + stamp;
    var ts = new Date().toLocaleString("en-IN");
    var logRows = [], perClass = {};

    for (var i = 0; i < n; i++) {
      var r = base[i];
      if (!r[10] && !r[2]) continue;
      var oldCls = (r[6] || "").toString().trim();
      var newCls = map[oldCls];
      if (!newCls || newCls === oldCls) continue;
      logRows.push([ts, batchId, r[2], r[10], oldCls, r[7], r[1], newCls, newAcYear, d.requesterUser || "", ""]);
      classDiv[i][0] = newCls;
      accYearCol[i][0] = newAcYear;
      perClass[oldCls + " → " + newCls] = (perClass[oldCls + " → " + newCls] || 0) + 1;
    }
    if (!logRows.length) { dropBackup_(bk); return {status:"error", message:"निवडलेल्या वर्गांमध्ये बढती देण्यासारखे विद्यार्थी सापडले नाहीत."}; }

    sh.getRange(2, 2, n, 1).setValues(accYearCol);
    sh.getRange(2, 7, n, 2).setValues(classDiv);

    var logSh = getOrCreateSheet("PromotionLog", PROMO_LOG_HEADERS);
    ensureRows_(logSh, logSh.getLastRow() + logRows.length);
    logSh.getRange(logSh.getLastRow() + 1, 1, logRows.length, PROMO_LOG_HEADERS.length).setValues(logRows);
    invalidateRosterCache_();
    return {rowIndex: 0, mode: "updated", count: logRows.length, ref: "promotion " + batchId,
      extra: {batchId: batchId, backupSheet: bkName, perClass: perClass}};
  } catch (err) {
    return {status:"error", message:err.toString()};
  } finally {
    if (locked) lock.releaseLock();
  }
}

function doUndoPromotion(d) {
  if (d.requesterRole !== "super") return {status:"error", message:"अधिकार फक्त Super Master ला आहे."};
  var batchId = (d.batchId || "").toString().trim();
  if (!batchId) return {status:"error", message:"Batch ID आवश्यक आहे."};
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(30000); locked = true;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("Students");
    var logSh = ss.getSheetByName("PromotionLog");
    if (!sh || !logSh || logSh.getLastRow() < 2) return {status:"error", message:"PromotionLog सापडला नाही."};
    var n = sh.getLastRow() - 1;
    var base = sh.getRange(2, 1, n, 11).getValues();
    var accYearCol = base.map(function(r) { return [r[1]]; });
    var classDiv = base.map(function(r) { return [r[6], r[7]]; });
    var rowByReg = {};
    for (var i = 0; i < n; i++) {
      var rg = (base[i][2] || "").toString().trim().toLowerCase();
      if (rg) rowByReg[rg] = i;
    }
    var lm = logSh.getLastRow() - 1;
    var lr = logSh.getRange(2, 1, lm, 11).getValues();
    var undone = [], restored = 0, skipped = 0, found = 0;
    for (var j = 0; j < lm; j++) {
      undone.push([lr[j][10]]);
      if ((lr[j][1] || "").toString() !== batchId || (lr[j][10] || "").toString() === "Y") continue;
      found++;
      var idx = rowByReg[(lr[j][2] || "").toString().trim().toLowerCase()];
      if (idx === undefined) { skipped++; continue; }
      // विद्यार्थी बढतीनंतर पुन्हा हलवला/बदलला असल्यास त्याला हात लावू नका
      if ((classDiv[idx][0] || "").toString().trim() !== (lr[j][7] || "").toString().trim()) { skipped++; continue; }
      classDiv[idx][0] = lr[j][4];
      if (lr[j][8]) accYearCol[idx][0] = lr[j][6];
      undone[j][0] = "Y";
      restored++;
    }
    if (!found) return {status:"error", message:"हा Batch सापडला नाही किंवा आधीच Undo झाला आहे."};
    sh.getRange(2, 2, n, 1).setValues(accYearCol);
    sh.getRange(2, 7, n, 2).setValues(classDiv);
    logSh.getRange(2, 11, lm, 1).setValues(undone);
    invalidateRosterCache_();
    return {rowIndex: 0, mode: "updated", count: restored, ref: "undo " + batchId, extra: {restored: restored, skipped: skipped}};
  } catch (err) {
    return {status:"error", message:err.toString()};
  } finally {
    if (locked) lock.releaseLock();
  }
}

// ---------- 4. "आज कोणी हजेरी भरली नाही" — प्रशासकीय देखरेख ----------
function doGetAttendancePending(p, cb) {
  try {
    if (p.requesterRole !== "super" && p.requesterRole !== "master") {
      return wrap(cb, {status:"error", message:"ही माहिती पाहण्याचा अधिकार Super Master / Master ला आहे."});
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var date = ISO_DATE_RE.test((p.date || "").toString().trim()) ? p.date.toString().trim() : todayStr();
    var dm = ISO_DATE_RE.exec(date);
    var isSunday = new Date(parseInt(dm[1], 10), parseInt(dm[2], 10) - 1, parseInt(dm[3], 10)).getDay() === 0;

    // अपेक्षित वर्ग-शिक्षक: Users sheet मधील role=teacher + AssignedClass + मोबाईल (स्तंभ F)
    var teachers = [];
    var uSh = ss.getSheetByName("Users");
    if (uSh && uSh.getLastRow() > 1) {
      var ucols = Math.min(6, uSh.getMaxColumns());
      var uRows = uSh.getRange(2, 1, uSh.getLastRow() - 1, ucols).getValues();
      for (var i = 0; i < uRows.length; i++) {
        var u = uRows[i];
        if (normRole_(u[2]) !== "teacher") continue;
        var parts = (u[4] || "").toString().split("|");
        if (!parts[0]) continue;
        teachers.push({username: u[0], label: u[3] || u[0], iyatta: parts[0].trim(), tukdi: (parts[1] || "").trim(), mobile: ucols >= 6 ? phoneDigits_(u[5]) : ""});
      }
    }

    // हजेरी भरल्याची नोंद: AttMarks (प्रत्येक Save ची एक ओळ — शून्य-अनुपस्थित Save सुद्धा). AttMarks नसेल तरच Log चा शेवटचा भाग.
    var marked = {};
    var marks = attMarksRowsForDate_(date);
    if (marks !== null) {
      marks.forEach(function(m) {
        marked[(m[1] || "").toString().trim() + "-" + (m[2] || "").toString().trim()] = {by: m[6], time: (m[7] || "").toString().split(",").pop().trim()};
      });
    } else {
      {   // V19.48: त्या तारखेच्या महिन्याची Log_YYYY_MM (स्थलांतरापूर्वी जुनी Log सुद्धा)
        var lRows = logRowsForMonth_(date.slice(0, 7).replace("-", "_"), 4000);
        for (var j = 0; j < lRows.length; j++) {
          var lr = lRows[j];
          if ((lr[3] || "").toString() !== "saveAttendance") continue;
          var ref = (lr[4] || "").toString();
          var sp = ref.lastIndexOf(" ");
          if (sp < 1 || ref.slice(sp + 1) !== date) continue;
          var tm = lr[0] instanceof Date ? Utilities.formatDate(lr[0], "Asia/Kolkata", "HH:mm") : (lr[0] || "").toString();
          marked[ref.slice(0, sp)] = {by: lr[1], time: tm};
        }
      }
    }
    // त्या दिवशी अनुपस्थित विद्यार्थी नोंदवलेले वर्गही "भरले" समजा
    readAttendanceDateRows_(date).forEach(function(r) {
      if (fmt(r[1]) !== date) return;
      var k2 = (r[2] || "").toString().trim() + "-" + (r[3] || "").toString().trim();
      if (!marked[k2]) marked[k2] = {by: r[8], time: ""};
    });

    var pending = [], done = [], covered = {};
    teachers.forEach(function(t) {
      var key = t.iyatta + "-" + t.tukdi;
      covered[key] = true;
      var m = marked[key];
      if (m) done.push({username: t.username, label: t.label, iyatta: t.iyatta, tukdi: t.tukdi, mobile: t.mobile, by: m.by, time: m.time});
      else pending.push({username: t.username, label: t.label, iyatta: t.iyatta, tukdi: t.tukdi, mobile: t.mobile});
    });
    pending.sort(function(a, b) { return classSortKey_(a.iyatta) - classSortKey_(b.iyatta) || (a.tukdi < b.tukdi ? -1 : 1); });

    // Students मध्ये आहेत पण कोणत्याही वर्ग-शिक्षकाला नेमलेले नाहीत असे वर्ग (वर्ग-यादी Cache मधून)
    var unassigned = [], seen = {};
    getRosterCached_("", "", false).data.forEach(function(st) {
      var ci = (st.iyatta || "").toString().trim(), ct = (st.tukdi || "").toString().trim();
      if (!ci) return;
      var ck = ci + "-" + ct;
      if (seen[ck]) return;
      seen[ck] = true;
      if (!covered[ck]) unassigned.push({iyatta: ci, tukdi: ct});
    });
    return wrap(cb, {status:"ok", date: date, isSunday: isSunday, totalExpected: teachers.length,
      pending: pending, done: done, unassigned: unassigned});
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}


// =====================================================================
// ===== V19.36 — सुरक्षा: Server-side Login + स्वाक्षरी केलेला Token =====
// =====================================================================
// पूर्वी Role फक्त browser पाठवत होता (कोणीही requesterRole=super पाठवू शकत होता). आता Login ला server Username/Password तपासतो
// व HMAC ने स्वाक्षरी केलेला Token देतो; पुढील प्रत्येक विनंतीत Token तपासून Role server स्वतः ठरवतो.
// स्थलांतर: (१) setSuperCredentials('user_s','पासवर्ड') (२) नवीन Deploy (३) Login चाचणी (४) enableAuthEnforcement()
var TOKEN_TTL_MS = 12 * 60 * 60 * 1000;
var VALID_ROLES = ["super","master","deo","cert","teacher","peon"];
// Users sheet मध्ये "Peon" / " Teacher " असे मोठ्या अक्षरात किंवा space सह लिहिले तरी लहान अक्षरात बदलून तपासतो.
function normRole_(r) { return (r || "").toString().trim().toLowerCase(); }
var AUTH_PUBLIC_ACTIONS = {ping:1, login:1, verify:1};
var TEACHER_CLASS_ACTIONS = {getClassStudents:1, getTeacherDashboard:1, saveAttendance:1, getAttendance:1, getAttendanceAnalytics:1, getEarlyLeave:1, saveEarlyLeave:1, getScholarship:1, saveScholarship:1, getFees:1, getTransport:1};   // V19.54: getFees/getTransport जोडले — शिक्षक वर्ग-यादी मोडमध्ये दुसऱ्या वर्गाची/संपूर्ण शाळेची फी मागवू शकत नाही
// Peon (शिपाई) फक्त हेच करू शकतो: विद्यार्थी शोध/प्रोफाइल (getAll), घरी जाणाऱ्यांची यादी, मार्क-आउट, स्वतःचा Password बदल
var PEON_ACTIONS = {getAll:1, getEarlyLeavePeon:1, markEarlyLeaveOut:1, changePassword:1, getPhotoData:1, getClassTeachers:1};
var _authSecretMem = null;

function authSecret_() {
  if (_authSecretMem) return _authSecretMem;
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty("AUTH_SECRET");
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty("AUTH_SECRET", s); }
  _authSecretMem = s;
  return s;
}
function hmacB64_(msg) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(msg, authSecret_()));
}
function safeEq_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var d = 0;
  for (var i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}
function signToken_(payload) {
  var b64 = Utilities.base64EncodeWebSafe(JSON.stringify(payload), Utilities.Charset.UTF_8);
  return b64 + "." + hmacB64_(b64);
}
function verifyToken_(tok) {
  try {
    if (!tok) return null;
    var parts = String(tok).split(".");
    if (parts.length !== 2) return null;
    if (!safeEq_(hmacB64_(parts[0]), parts[1])) return null;
    var payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString("UTF-8"));
    if (!payload || !payload.exp || payload.exp < Date.now()) return null;
    return payload;
  } catch (e) { return null; }
}
function authEnforced_() {
  return PropertiesService.getScriptProperties().getProperty("AUTH_ENFORCE") === "1";
}

// प्रत्येक doGet/doPost च्या सुरुवातीला. null = पुढे जा; object = त्रुटी परत करा.
function applyAuth_(p) {
  var act = (p.action || "").toString();
  // बाहेरून आलेल्या (browser ने पाठवलेल्या) अंतर्गत चिन्हांना मान्यता नाही
  delete p._authed;
  delete p._tClass;   // V19.47: शिक्षकाचा वर्ग फक्त Token मधून (browser ने पाठवलेला नाही)
  if (AUTH_PUBLIC_ACTIONS[act]) return null;
  var payload = verifyToken_(p.token);
  if (payload) {
    p._authed = true;
    p.requesterRole = payload.r;
    p.requesterUser = payload.u;
    p.auditUser = payload.u;
    p.auditRole = payload.r;
    p._tClass = (payload.r === "teacher") ? String(payload.c || "") : "";   // V19.47
    if (payload.r === "peon" && !PEON_ACTIONS[act]) {
      return {status:"error", code:"forbidden", message:"या User ला हे करण्याचा अधिकार नाही."};
    }
    if (payload.r === "teacher" && TEACHER_CLASS_ACTIONS[act] && payload.c) {
      var cp = String(payload.c).split("|");
      p.iyatta = (cp[0] || "").trim();
      p.tukdi = (cp[1] || "").trim();
    }
    return null;
  }
  if (authEnforced_()) {
    return {status:"error", code:"auth", message:"Session संपले किंवा Login झाले नाही — कृपया पुन्हा Login करा."};
  }
  return null; // जुने (अजून Enforce न केलेले) व्यवहार जसे आहेत तसे चालतात
}

function doLogin(p, cb) {
  try {
    var username = (p.username || "").toString().trim();
    var password = (p.password || "").toString();
    if (!username || !password) return wrap(cb, {status:"error", code:"badcred", message:"Username व Password आवश्यक आहे."});
    var cache = CacheService.getScriptCache();
    var fk = "lf_" + username.toLowerCase();
    var fails = parseInt(cache.get(fk) || "0", 10) || 0;
    if (fails >= 8) return wrap(cb, {status:"error", code:"locked", message:"खूप चुकीचे प्रयत्न झाले — १५ मिनिटांनी पुन्हा प्रयत्न करा."});

    var user = null;
    var sh = getUsersSheet();
    if (sh.getLastRow() > 1) {
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
      for (var i = 0; i < rows.length; i++) {
        if ((rows[i][0] || "").toString().trim() === username && passwordMatches_(rows[i][1], password)) {
          var rl = normRole_(rows[i][2]);
          if (VALID_ROLES.indexOf(rl) === -1) {
            return wrap(cb, {status:"error", code:"badrole", message:"या User चा Role (" + (rows[i][2] || "रिकामा") + ") अमान्य आहे. Users sheet मध्ये Role: super / master / deo / cert / teacher / peon यापैकी एक हवा."});
          }
          user = {role: rl, label: rows[i][3] || username, assignedClass: (rows[i][4] || "").toString()};
          break;
        }
      }
    }
    if (!user) {
      var props = PropertiesService.getScriptProperties();
      var su = props.getProperty("SUPER_USER"), sp = props.getProperty("SUPER_PASS");
      if (su && sp && username === su && safeEq_(password, sp)) user = {role: "super", label: "Super Master User", assignedClass: ""};
    }
    if (!user) {
      cache.put(fk, String(fails + 1), 900);
      logAudit(username, "", "loginFailed", "");
      return wrap(cb, {status:"error", code:"badcred", message:"Username किंवा Password चुकीचा आहे."});
    }
    cache.remove(fk);
    try { PropertiesService.getScriptProperties().setProperty("LOGIN_SEEN_" + user.role, String(Date.now())); } catch (eSeen) {}   // V19.52: या भूमिकेचा Token-Login चालल्याची नोंद (सक्तीपूर्वी पूर्व-तपासणीसाठी)
    var exp = Date.now() + TOKEN_TTL_MS;
    var token = signToken_({u: username, r: user.role, c: user.assignedClass, exp: exp});
    logAudit(username, user.role, "login", "");
    return wrap(cb, {status:"ok", token: token, exp: exp,
      user: {username: username, role: user.role, label: user.label, assignedClass: user.assignedClass}});
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ---- Apps Script Editor मधून Run करण्याचे functions ----
function setSuperCredentials(username, password) {
  if (!username || !password || String(password).length < 8) throw new Error("Username द्या व Password किमान ८ अक्षरी असावा.");
  var props = PropertiesService.getScriptProperties();
  props.setProperty("SUPER_USER", String(username));
  props.setProperty("SUPER_PASS", String(password));
  Logger.log("Super credentials सेट झाले. आता नवीन Deploy करून Login तपासा, मग enableAuthEnforcement() Run करा.");
}
// V19.52: Token सक्ती चालू करण्यापूर्वी — Users मधील प्रत्येक भूमिकेने Token-Login यशस्वीपणे केल्याची खात्री (अन्यथा त्यांना बाहेर पडावे लागेल)
function securityPreflight() {
  var props = PropertiesService.getScriptProperties();
  var roles = {}, defaults = 0, plain = 0, hashed = 0;
  var sh = getUsersSheet();
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function(r) {
      if (!r[0]) return;
      var rl = normRole_(r[2]); roles[rl] = (roles[rl] || 0) + 1;
      var pw = (r[1] || "").toString();
      if (isHashedPw_(pw)) hashed++; else if (pw) plain++;
      if (pw === "Pass@1234") defaults++;
    });
  }
  var seen = {}, missing = [];
  Object.keys(roles).forEach(function(rl) {
    var t = props.getProperty("LOGIN_SEEN_" + rl);
    seen[rl] = t ? new Date(parseInt(t, 10)).toISOString() : "";
    if (!t) missing.push(rl);
  });
  var superOk = !!(props.getProperty("SUPER_USER") && props.getProperty("SUPER_PASS"));
  var rep = {superConfigured: superOk, rolesInUsers: roles, tokenLoginSeen: seen, rolesNeverLoggedInWithToken: missing,
    enforced: authEnforced_(), passwordsHashed: pwHashed_(), plainPasswords: plain, hashedPasswords: hashed, defaultPasswordUsers: defaults,
    readyForEnforcement: superOk && !missing.length};
  Logger.log(JSON.stringify(rep, null, 1));
  return rep;
}
function enableAuthEnforcement() {
  var rep = securityPreflight();
  if (!rep.superConfigured) {
    throw new Error("आधी setSuperCredentials('user_s','पासवर्ड') Run करा, नाहीतर Super Master ला बाहेर पडावे लागेल.");
  }
  if (rep.rolesNeverLoggedInWithToken.length) {
    throw new Error("या भूमिकांनी नवीन (Token) Login अजून एकदाही केलेला नाही: " + rep.rolesNeverLoggedInWithToken.join(", ") +
      ". आधी प्रत्येक भूमिकेच्या एका User ने ॲपमध्ये Login करून पाहा. नंतर पुन्हा Run करा — किंवा तरीही सक्ती करायची असेल तर enableAuthEnforcementForce() Run करा.");
  }
  authSecret_();
  PropertiesService.getScriptProperties().setProperty("AUTH_ENFORCE", "1");
  Logger.log("Token सक्ती चालू. आता Token शिवाय कोणतीही विनंती नाकारली जाईल. (परत जाण्यासाठी: disableAuthEnforcement)");
}
function enableAuthEnforcementForce() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty("SUPER_USER") || !props.getProperty("SUPER_PASS")) throw new Error("आधी setSuperCredentials Run करा.");
  authSecret_();
  props.setProperty("AUTH_ENFORCE", "1");
  Logger.log("Token सक्ती (Force) चालू. पूर्व-तपासणी वगळली.");
}
// सर्व साधे Password hash करा. पूर्वअट: Token सक्ती चालू + Super Password सर्व्हरवर. आधी Users ची पूर्ण प्रत "Users_BACKUP_<वेळ>" (लपवलेली) बनते.
function hashAllPasswords() {
  var props = PropertiesService.getScriptProperties();
  if (!authEnforced_()) throw new Error("आधी Token सक्ती चालू करा (enableAuthEnforcement) — नाहीतर जुना Login (साधे Password वाचणारा) बंद पडेल.");
  if (!props.getProperty("SUPER_USER") || !props.getProperty("SUPER_PASS")) throw new Error("आधी setSuperCredentials Run करा.");
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = getUsersSheet();
    var n = sh.getLastRow();
    if (n < 2) { props.setProperty("PW_HASHED", "1"); return {hashed: 0, backup: ""}; }
    var stamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyyMMdd_HHmm");
    var bk = sh.copyTo(ss); bk.setName("Users_BACKUP_" + stamp);
    try { bk.hideSheet(); } catch (eH) {}
    if (bk.getLastRow() !== n) { ss.deleteSheet(bk); throw new Error("Backup पडताळणी अयशस्वी — काहीही बदलले नाही."); }
    var rng = sh.getRange(2, 2, n - 1, 1), vals = rng.getValues(), cnt = 0;
    for (var i = 0; i < vals.length; i++) {
      var v = (vals[i][0] || "").toString();
      if (v && !isHashedPw_(v)) { vals[i][0] = hashPassword_(v); cnt++; }
    }
    rng.setValues(vals);
    props.setProperty("PW_HASHED", "1");
    Logger.log("Password hash झाले: " + cnt + ". Backup Sheet: Users_BACKUP_" + stamp + " (लपवलेली — पूर्ण साधे Password तिथे आहेत; पडताळणी झाल्यावर ती मिटवा).");
    return {hashed: cnt, backup: "Users_BACKUP_" + stamp};
  } finally { try { lock.releaseLock(); } catch (e) {} }
}
// आणीबाणी: सर्वात ताज्या Users_BACKUP_* मधून Password स्तंभ परत आणतो व PW_HASHED=0 करतो
function restoreUsersFromBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var names = ss.getSheets().map(function(x) { return x.getName(); }).filter(function(n) { return /^Users_BACKUP_/.test(n); }).sort();
  if (!names.length) throw new Error("Users_BACKUP_* Sheet सापडली नाही.");
  var bk = ss.getSheetByName(names[names.length - 1]), sh = getUsersSheet();
  var src = bk.getRange(2, 1, bk.getLastRow() - 1, 2).getValues(), cur = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var map = {}; src.forEach(function(r) { map[(r[0] || "").toString().trim()] = r[1]; });
  var restored = 0;
  for (var i = 0; i < cur.length; i++) { var k = (cur[i][0] || "").toString().trim(); if (map.hasOwnProperty(k)) { cur[i][1] = map[k]; restored++; } }
  sh.getRange(2, 2, cur.length, 1).setValues(cur.map(function(r) { return [r[1]]; }));
  PropertiesService.getScriptProperties().setProperty("PW_HASHED", "0");
  Logger.log("Password परत आणले: " + restored + " (स्त्रोत: " + names[names.length - 1] + ").");
  return {restored: restored, from: names[names.length - 1]};
}
function disableAuthEnforcement() {
  PropertiesService.getScriptProperties().setProperty("AUTH_ENFORCE", "0");
}
// सर्व चालू Login Session तात्काळ रद्द करण्यासाठी (गुप्त की बदलते)
function revokeAllSessions() {
  PropertiesService.getScriptProperties().setProperty("AUTH_SECRET", Utilities.getUuid() + Utilities.getUuid());
  _authSecretMem = null;
}

function doGetSecurityStatus(p, cb) {
  try {
    if (p.requesterRole !== "super") return wrap(cb, {status:"error", message:"अधिकार फक्त Super Master ला आहे."});
    var props = PropertiesService.getScriptProperties();
    var seedPass = false;
    var sh = getUsersSheet();
    if (sh.getLastRow() > 1) {
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
      for (var i = 0; i < rows.length; i++) if ((rows[i][1] || "").toString() === "Pass@1234") seedPass = true;
    }
    var pre = securityPreflight();
    return wrap(cb, {status:"ok", enforced: authEnforced_(), tokenActive: !!p._authed,
      superConfigured: !!(props.getProperty("SUPER_USER") && props.getProperty("SUPER_PASS")), defaultPasswordsPresent: seedPass,
      passwordsHashed: pre.passwordsHashed, plainPasswords: pre.plainPasswords, hashedPasswords: pre.hashedPasswords,
      rolesNeverLoggedInWithToken: pre.rolesNeverLoggedInWithToken, readyForEnforcement: pre.readyForEnforcement});
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// =====================================================================
// ===== V19.36 — हजेरी विश्लेषण (उपस्थिती % + सलग गैरहजेरी) =====
// =====================================================================
function doGetAttendanceAnalytics(p, cb) {
  try {
    var role = p.requesterRole;
    if (role !== "super" && role !== "master" && role !== "teacher") {
      return wrap(cb, {status:"error", message:"हे विश्लेषण पाहण्याचा अधिकार नाही."});
    }
    var iyatta = (p.iyatta || "").toString().trim();
    var tukdi = (p.tukdi || "").toString().trim();
    var month = (p.month || "").toString().trim();
    var minRun = Math.max(2, parseInt(p.minRun, 10) || 3);
    if (!iyatta) return wrap(cb, {status:"error", message:"वर्ग निवडा."});
    if (!/^\d{4}-\d{2}$/.test(month)) return wrap(cb, {status:"error", message:"महिना YYYY-MM स्वरूपात हवा."});
    var from = month + "-01", to = month + "-31";
    var ckey = "ana_" + month + "_" + encodeURIComponent(iyatta + "|" + tukdi) + "_" + minRun + "_" + attVer_() + "_" + rosterVer_();
    var hitA = cacheGetBig_(ckey);
    if (hitA) { hitA.cached = true; return wrap(cb, hitA); }

    var roster = getRosterCached_(iyatta, tukdi, false).data;

    // कार्यदिवस = त्या वर्ग-तुकडीसाठी ज्या तारखांना हजेरी Save झाली (शून्य अनुपस्थित असलेल्या Save सह)
    var workByDiv = {};
    function addWork(div, date) {
      if (date < from || date > to) return;
      (workByDiv[div] = workByDiv[div] || {})[date] = true;
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // कार्यदिवस: AttMarks (प्रत्येक वर्गाची प्रत्येक दिवशीची Save) — नसेल तरच जुनी पद्धत (Log)
    var marksRows = attMarksMonthRows_(month.replace("-", "_"));
    if (marksRows !== null) {
      marksRows.forEach(function(m) {
        var div = (m[1] || "").toString().trim() + "-" + (m[2] || "").toString().trim();
        if (div.indexOf(iyatta + "-") !== 0) return;
        if (tukdi && div !== iyatta + "-" + tukdi) return;
        addWork(div, fmt(m[0]));
      });
    } else {
      {   // V19.48: त्या महिन्याची Log_YYYY_MM (स्थलांतरापूर्वी जुनी Log सुद्धा)
        var lRows = logRowsForMonth_(month.replace("-", "_"), 20000);   // [3]=Action, [4]=Reference
        for (var i = 0; i < lRows.length; i++) {
          if ((lRows[i][3] || "").toString() !== "saveAttendance") continue;
          var ref = (lRows[i][4] || "").toString();
          var sp = ref.lastIndexOf(" ");
          if (sp < 1) continue;
          var div2 = ref.slice(0, sp), dt = ref.slice(sp + 1);
          if (div2.indexOf(iyatta + "-") !== 0) continue;
          if (tukdi && div2 !== iyatta + "-" + tukdi) continue;
          addWork(div2, dt);
        }
      }
    }
    var attRows = readAttendanceRows_(from, to);
    var absByReg = {};
    for (var a = 0; a < attRows.length; a++) {
      var r = attRows[a];
      var d0 = fmt(r[1]);
      if (d0 < from || d0 > to) continue;
      if ((r[2] || "").toString().trim() !== iyatta) continue;
      if (tukdi && (r[3] || "").toString().trim() !== tukdi) continue;
      addWork(iyatta + "-" + (r[3] || "").toString().trim(), d0); // अनुपस्थित नोंद असलेला दिवसही कार्यदिवस
      var rg = (r[4] || "").toString().trim().toLowerCase();
      (absByReg[rg] = absByReg[rg] || {})[d0] = true;
    }

    var students = [], streaks = [], totalPct = 0, counted = 0;
    for (var k = 0; k < roster.length; k++) {
      var st = roster[k];
      var div2 = st.iyatta + "-" + st.tukdi;
      var days = Object.keys(workByDiv[div2] || {}).sort();
      var absDates = absByReg[(st.regNo || "").toString().trim().toLowerCase()] || {};
      var absent = 0, run = 0, maxRun = 0, runStart = "", bestFrom = "", bestTo = "";
      for (var d = 0; d < days.length; d++) {
        if (absDates[days[d]]) {
          absent++;
          if (run === 0) runStart = days[d];
          run++;
          if (run > maxRun) { maxRun = run; bestFrom = runStart; bestTo = days[d]; }
        } else { run = 0; }
      }
      var present = Math.max(0, days.length - absent);
      var pct = days.length ? Math.round(present * 1000 / days.length) / 10 : null;
      if (pct !== null) { totalPct += pct; counted++; }
      var mobile = (st.whatsappMobile || st.alternateMobile || st.contact || "").toString();
      students.push({regNo: st.regNo, fullName: st.fullName, rollNo: st.rollNo, iyatta: st.iyatta, tukdi: st.tukdi,
        workingDays: days.length, absent: absent, present: present, pct: pct, maxRun: maxRun, mobile: mobile});
      if (maxRun >= minRun) streaks.push({regNo: st.regNo, fullName: st.fullName, iyatta: st.iyatta, tukdi: st.tukdi,
        run: maxRun, from: bestFrom, to: bestTo, mobile: mobile});
    }
    students.sort(function(a, b) {
      if (a.tukdi !== b.tukdi) return a.tukdi < b.tukdi ? -1 : 1;
      return (parseInt(a.rollNo, 10) || 0) - (parseInt(b.rollNo, 10) || 0);
    });
    streaks.sort(function(a, b) { return b.run - a.run; });
    var workingDays = 0;
    Object.keys(workByDiv).forEach(function(kk) { workingDays = Math.max(workingDays, Object.keys(workByDiv[kk]).length); });
    var resA = {status:"ok", month: month, iyatta: iyatta, tukdi: tukdi, minRun: minRun, workingDays: workingDays,
      avgPct: counted ? Math.round(totalPct * 10 / counted) / 10 : null, students: students, streaks: streaks};
    if (archivedFlagMonth_(month)) resA.archived = true;   // V19.53
    cachePutBig_(ckey, resA, 600);
    return wrap(cb, resA);
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// =====================================================================
// ===== V19.36 — वर्षअखेर: 10th Pass-out Archive, Roll No. पुन्हा क्रमांक, Log Archive =====
// =====================================================================
var ARCHIVE_PREFIX = "PassedOut_";

function pruneHiddenBackups_(prefix, keep) { pruneBackupSheets_(SpreadsheetApp.getActiveSpreadsheet(), prefix, keep); }
function pruneBackupSheets_(ss, prefix, keep) {
  try {
    var list = ss.getSheets().filter(function(sh) { return sh.getName().indexOf(prefix) === 0; });
    list.sort(function(a, b) { return a.getName() < b.getName() ? 1 : (a.getName() > b.getName() ? -1 : 0); });
    for (var i = keep; i < list.length; i++) ss.deleteSheet(list[i]);
  } catch (e) {}
}

function listArchiveBatches_() {
  var out = [];
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getSheets().forEach(function(sh) {
      if (sh.getName().indexOf(ARCHIVE_PREFIX) !== 0 || sh.getLastRow() < 2) return;
      var lc = sh.getLastColumn();
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, lc).getValues();
      var map = {};
      rows.forEach(function(r) {
        var b = (r[lc - 2] || "").toString();
        if (!b) return;
        if (!map[b]) map[b] = {batchId: b, sheet: sh.getName(), time: r[lc - 3], iyatta: (r[6] || "").toString(), count: 0};
        map[b].count++;
      });
      Object.keys(map).forEach(function(k) { out.push(map[k]); });
    });
  } catch (e) {}
  out.sort(function(a, b) { return a.batchId < b.batchId ? 1 : -1; });
  return out.slice(0, 5);
}

function doArchivePassedOut(d) {
  if (d.requesterRole !== "super") return {status:"error", message:"अधिकार फक्त Super Master ला आहे."};
  if ((d.confirm || "") !== "ARCHIVE") return {status:"error", message:"खात्रीसाठी ARCHIVE टाइप करणे आवश्यक आहे."};
  var cls = (d.iyatta || "").toString().trim();
  var label = (d.label || "").toString().trim();
  if (!cls) return {status:"error", message:"वर्ग निवडा."};
  if (!/^[0-9A-Za-z_\- ]{3,20}$/.test(label)) return {status:"error", message:"Archive चे नाव (उदा. 2026-27) द्या."};
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(30000); locked = true;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("Students");
    if (!sh || sh.getLastRow() < 2) return {status:"error", message:"Students sheet रिकामी आहे."};
    var sc = sh.getLastColumn(), n = sh.getLastRow() - 1;
    var header = sh.getRange(1, 1, 1, sc).getValues()[0];
    var all = sh.getRange(2, 1, n, sc).getValues();
    var stamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyyMMdd_HHmm");

    var moved = [], kept = [];
    for (var i = 0; i < all.length; i++) {
      var r = all[i];
      if ((r[6] || "").toString().trim() === cls && (r[10] || r[2])) moved.push(r); else kept.push(r);
    }
    if (!moved.length) return {status:"error", message:"त्या वर्गात विद्यार्थी सापडले नाहीत."};

    makeStudentsBackup_(sh, "Students_BeforeArchive_" + stamp, "Students_BeforeArchive_");

    var arName = ARCHIVE_PREFIX + label.replace(/\s+/g, "_");
    var ar = ss.getSheetByName(arName);
    if (!ar) {
      ar = ss.insertSheet(arName);
      trimNewSheet_(ar, sc + 3);
      ar.getRange(1, 1, 1, sc + 3).setValues([header.concat(["ArchivedAt", "BatchId", "ArchivedBy"])]);
      ar.setFrozenRows(1);
    }
    var batchId = "A" + stamp;
    var ts = new Date().toLocaleString("en-IN");
    var outRows = moved.map(function(r) { return r.concat([ts, batchId, d.requesterUser || ""]); });
    ensureRows_(ar, ar.getLastRow() + outRows.length); ensureCols_(ar, sc + 3);
    ar.getRange(ar.getLastRow() + 1, 1, outRows.length, sc + 3).setValues(outRows);

    sh.getRange(2, 1, n, sc).clearContent();
    if (kept.length) sh.getRange(2, 1, kept.length, sc).setValues(kept);
    invalidateRosterCache_();
    return {rowIndex: 0, mode: "updated", count: moved.length, ref: "archive " + batchId,
      extra: {batchId: batchId, archiveSheet: arName}};
  } catch (err) {
    return {status:"error", message:err.toString()};
  } finally {
    if (locked) lock.releaseLock();
  }
}

function doRestorePassedOut(d) {
  if (d.requesterRole !== "super") return {status:"error", message:"अधिकार फक्त Super Master ला आहे."};
  var batchId = (d.batchId || "").toString().trim();
  if (!batchId) return {status:"error", message:"Batch ID आवश्यक आहे."};
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(30000); locked = true;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("Students");
    var sc = sh.getLastColumn();
    var existing = {};
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues().forEach(function(r) { existing[(r[0] || "").toString().trim().toLowerCase()] = true; });
    }
    var back = [], skipped = 0, found = 0;
    ss.getSheets().forEach(function(ar) {
      if (ar.getName().indexOf(ARCHIVE_PREFIX) !== 0 || ar.getLastRow() < 2) return;
      var lc = ar.getLastColumn();
      var rows = ar.getRange(2, 1, ar.getLastRow() - 1, lc).getValues();
      var keep = [], changed = false;
      rows.forEach(function(r) {
        if ((r[lc - 2] || "").toString() !== batchId) { keep.push(r); return; }
        found++;
        var key = (r[2] || "").toString().trim().toLowerCase();
        if (key && existing[key]) { skipped++; keep.push(r); return; }
        back.push(r.slice(0, sc));
        changed = true;
      });
      if (changed) {
        ar.getRange(2, 1, rows.length, lc).clearContent();
        if (keep.length) ar.getRange(2, 1, keep.length, lc).setValues(keep);
      }
    });
    if (!found) return {status:"error", message:"हा Batch सापडला नाही."};
    if (back.length) { ensureRows_(sh, sh.getLastRow() + back.length); sh.getRange(sh.getLastRow() + 1, 1, back.length, sc).setValues(back); }
    invalidateRosterCache_();
    return {rowIndex: 0, mode: "updated", count: back.length, ref: "restore " + batchId, extra: {restored: back.length, skipped: skipped}};
  } catch (err) {
    return {status:"error", message:err.toString()};
  } finally {
    if (locked) lock.releaseLock();
  }
}

function nameCmp_(a, b) {
  a = (a || "").toString(); b = (b || "").toString();
  try { return a.localeCompare(b, "mr"); } catch (e) { return a < b ? -1 : (a > b ? 1 : 0); }
}
function isGirl_(g) { return /female|मुलगी|girl|स्त्री/i.test((g || "").toString()); }

function doRenumberRolls(d) {
  if (d.requesterRole !== "super") return {status:"error", message:"अधिकार फक्त Super Master ला आहे."};
  if ((d.confirm || "") !== "RENUMBER") return {status:"error", message:"खात्रीसाठी RENUMBER टाइप करणे आवश्यक आहे."};
  var order = (d.order || "name").toString();
  if (["name", "boys_first", "girls_first"].indexOf(order) === -1) return {status:"error", message:"क्रम चुकीचा आहे."};
  var classes = [];
  try { classes = JSON.parse(d.classesJson || "[]"); } catch (e) { return {status:"error", message:"वर्ग यादी वाचता आली नाही."}; }
  if (!classes.length) return {status:"error", message:"किमान एक वर्ग निवडा."};
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(30000); locked = true;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("Students");
    if (!sh || sh.getLastRow() < 2) return {status:"error", message:"Students sheet रिकामी आहे."};
    var n = sh.getLastRow() - 1;
    var base = sh.getRange(2, 1, n, 13).getValues(); // A..M
    var stamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyyMMdd_HHmm");
    makeStudentsBackup_(sh, "Students_BeforeRoll_" + stamp, "Students_BeforeRoll_");

    var groups = {};
    for (var i = 0; i < n; i++) {
      var r = base[i];
      if (!r[10] && !r[2]) continue;
      var cls = (r[6] || "").toString().trim();
      if (classes.indexOf(cls) === -1) continue;
      var key = cls + "|" + (r[7] || "").toString().trim();
      (groups[key] = groups[key] || []).push({i: i, name: r[10], girl: isGirl_(r[12]), reg: (r[2] || "").toString()});
    }
    var rollCol = base.map(function(r) { return [r[9]]; });
    var changedCount = 0;
    Object.keys(groups).forEach(function(k) {
      var list = groups[k];
      list.sort(function(a, b) {
        if (order !== "name" && a.girl !== b.girl) return (order === "boys_first") ? (a.girl ? 1 : -1) : (a.girl ? -1 : 1);
        var c = nameCmp_(a.name, b.name);
        return c !== 0 ? c : (a.reg < b.reg ? -1 : (a.reg > b.reg ? 1 : 0));
      });
      list.forEach(function(it, idx) {
        if (String(rollCol[it.i][0]) !== String(idx + 1)) changedCount++;
        rollCol[it.i][0] = idx + 1;
      });
    });
    sh.getRange(2, 10, n, 1).setValues(rollCol);
    invalidateRosterCache_();
    return {rowIndex: 0, mode: "updated", count: changedCount, ref: "renumber " + classes.join(","),
      extra: {groups: Object.keys(groups).length, backupSheet: "Students_BeforeRoll_" + stamp}};
  } catch (err) {
    return {status:"error", message:err.toString()};
  } finally {
    if (locked) lock.releaseLock();
  }
}

// ---- Log Archive ----
function parseLogDate_(v) {
  if (v instanceof Date) return v;
  var s = (v || "").toString();
  var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s); // en-IN: dd/mm/yyyy
  if (m) return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
  m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  return null;
}
function archiveOldLog_(days) {
  if (logMigrated_()) return 0;   // V19.48: मासिक Log_ मुळे आता गरज नाही
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Log");
  if (!sh || sh.getLastRow() < 2) return 0;
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(20000); locked = true;
    var n = sh.getLastRow() - 1;
    var rows = sh.getRange(2, 1, n, 5).getValues();
    var cutoff = new Date(Date.now() - days * 86400000);
    var old = [], keep = [];
    rows.forEach(function(r) {
      var dt = parseLogDate_(r[0]);
      if (dt && dt < cutoff) old.push(r); else keep.push(r);
    });
    if (!old.length) return 0;
    var ar = ss.getSheetByName("Log_Archive");
    if (!ar) { ar = ss.insertSheet("Log_Archive"); trimNewSheet_(ar, 5); ar.appendRow(["Timestamp","User","Role","Action","Reference"]); }
    ensureRows_(ar, ar.getLastRow() + old.length);
    // आधी Archive मध्ये लिहा, मग Log मधून काढा — मध्येच अपयश आल्यास नोंदी हरवत नाहीत
    ar.getRange(ar.getLastRow() + 1, 1, old.length, 5).setValues(old.map(function(r) { r[0] = (r[0] instanceof Date) ? r[0].toLocaleString("en-IN") : r[0]; return r; }));
    sh.getRange(2, 1, n, 5).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, 5).setValues(keep);
    return old.length;
  } finally {
    if (locked) lock.releaseLock();
  }
}
function doArchiveLogNow(d) {
  if (d.requesterRole !== "super") return {status:"error", message:"अधिकार फक्त Super Master ला आहे."};
  var days = Math.max(30, parseInt(d.days, 10) || 90);
  try {
    var moved = archiveOldLog_(days);
    return {rowIndex: 0, mode: "updated", count: moved, ref: "archiveLog " + days + "d"};
  } catch (err) {
    return {status:"error", message:err.toString()};
  }
}


// =====================================================================
// ===== V19.37 — घरी जाणारे विद्यार्थी (Early Leave) =====
// =====================================================================
// नोंद: Master/Super = कोणताही वर्ग-तुकडी; वर्ग-शिक्षक = फक्त स्वतःचा वर्ग; Peon = फक्त आजची "परवानगी" यादी पाहून मार्क-आउट.
// एका विद्यार्थ्याची एका दिवसाची एकच नोंद (date + regNo). हजर विद्यार्थ्यांचीच नोंद; गैरहजर (हजेरी नोंदवहीनुसार) नाकारली जाते.
var EL_LEGACY_SHEET = "EarlyLeave";   // जुनी एकच Sheet — आता आठवडानिहाय "EL_<सोमवारची तारीख>" Sheets (V19.41)
var EL_HEADERS = ["Timestamp","Date","Class","Division","RegNo","StudentId","FullName","Type","Period","RecordedBy","RecordedRole","Status","MarkedBy","MarkedAt"];
var EL_TYPES = {permitted:1, unpermitted:1};

// ---- V19.41: आठवडानिहाय Sheet (सोमवार ते रविवार) — "EL_2026-10-05" (सोमवारची तारीख) ----
function elWeekStart_(date) {
  var m = ISO_DATE_RE.exec(date);
  var t = Date.UTC(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  var back = (new Date(t).getUTCDay() + 6) % 7;          // सोमवार=0 … रविवार=6
  var d = new Date(t - back * 86400000);
  function p2(n) { return ("0" + n).slice(-2); }
  return d.getUTCFullYear() + "-" + p2(d.getUTCMonth() + 1) + "-" + p2(d.getUTCDate());
}
function elWeekSheetName_(date) { return "EL_" + elWeekStart_(date); }
function getELWeekSheet_(date, create) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(elWeekSheetName_(date));
  if (sh) return sh;
  if (!create) return archiveSheet_(elWeekSheetName_(date));   // V19.53 (फक्त वाचन — लेखनापूर्वी archivedWriteError_ तपासली जाते)
  sh = ss.insertSheet(elWeekSheetName_(date));
  trimNewSheet_(sh, EL_HEADERS.length);
  sh.appendRow(EL_HEADERS);
  sh.getRange("B:B").setNumberFormat("@");
  sh.setFrozenRows(1);
  return sh;
}
// जुनी "EarlyLeave" Sheet एकदाच आठवडानिहाय Sheets मध्ये वाटून (जुनी Sheet "EarlyLeave_OLD_..." नावाने लपवली जाते — काहीही हरवत नाही).
// Lock स्वतः घेतो — म्हणून हे function दुसऱ्या Lock च्या आत बोलावू नये.
function elEnsureMigrated_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty("EL_WEEKLY_MIGRATED") === "1") return;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (props.getProperty("EL_WEEKLY_MIGRATED") === "1") return;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var old = ss.getSheetByName(EL_LEGACY_SHEET);
    if (old) {
      var n = old.getLastRow() - 1;
      if (n > 0) {
        var groups = {};
        old.getRange(2, 1, n, 14).getValues().forEach(function(r) {
          var d = fmt(r[1]);
          if (!elDateValid_(d)) return;
          r[1] = d;
          var wk = elWeekStart_(d);
          (groups[wk] = groups[wk] || []).push(r);
        });
        Object.keys(groups).forEach(function(wk) {
          var sh = getELWeekSheet_(wk, true), rows = groups[wk];
          ensureRows_(sh, sh.getLastRow() + rows.length);
          sh.getRange(sh.getLastRow() + 1, 1, rows.length, 14).setValues(rows);
        });
      }
      old.setName("EarlyLeave_OLD_" + Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyyMMdd_HHmm"));
      try { old.hideSheet(); } catch (e) {}
    }
    props.setProperty("EL_WEEKLY_MIGRATED", "1");
  } finally {
    lock.releaseLock();
  }
}
function elDateValid_(d) { return ISO_DATE_RE.test(d || ""); }

function elStudentMap_() {
  var map = {};
  getRosterCached_("", "", false).data.forEach(function(st) { map[(st.regNo || "").toString().trim().toLowerCase()] = st; });
  return map;
}
function elRowToObj_(r, stMap) {
  var st = stMap[(r[4] || "").toString().trim().toLowerCase()] || {};
  return {date: fmt(r[1]), iyatta: (r[2] || "").toString(), tukdi: (r[3] || "").toString(), regNo: (r[4] || "").toString(), studentId: r[5],
    fullName: r[6], type: (r[7] || "").toString(), period: parseInt(r[8], 10) || 0, recordedBy: r[9], recordedRole: r[10],
    status: (r[11] || "").toString(), markedBy: r[12] || "", markedAt: r[13] || "",
    gender: st.gender || "", rollNo: st.rollNo || "", contact: (st.contact || "").toString(), whatsapp: (st.whatsappMobile || "").toString(),
    alt: (st.alternateMobile || "").toString(), photoUrl: st.photoUrl || ""};
}
function elReadAll_(sh) {
  var n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, 14).getValues() : [];
}
function elSort_(list) {
  list.sort(function(a, b) {
    var c = classSortKey_(a.iyatta) - classSortKey_(b.iyatta);
    if (c) return c;
    if (a.tukdi !== b.tukdi) return a.tukdi < b.tukdi ? -1 : 1;
    if (a.period !== b.period) return a.period - b.period;
    return (parseInt(a.rollNo, 10) || 0) - (parseInt(b.rollNo, 10) || 0);
  });
  return list;
}
function elAbsentSet_(date, iyatta, tukdi) {
  var set = {};
  readAttendanceDateRows_(date).forEach(function(r) {
    if (fmt(r[1]) !== date) return;
    if ((r[2] || "").toString().trim() !== iyatta) return;
    if (tukdi && (r[3] || "").toString().trim() !== tukdi) return;
    set[(r[4] || "").toString().trim().toLowerCase()] = true;
  });
  return set;
}

function doGetEarlyLeave(p, cb) {
  try {
    var role = p.requesterRole;
    if (role !== "super" && role !== "master" && role !== "teacher") return wrap(cb, {status:"error", message:"हे पान पाहण्याचा अधिकार नाही."});
    var date = (p.date || todayStr()).toString().trim();
    if (!elDateValid_(date)) return wrap(cb, {status:"error", message:"तारीख YYYY-MM-DD स्वरूपात हवी."});
    var iyatta = (p.iyatta || "").toString().trim();
    var tukdi = (p.tukdi || "").toString().trim();
    if (role === "teacher" && !iyatta) return wrap(cb, {status:"error", message:"वर्ग-शिक्षकाला वर्ग नेमलेला नाही."});
    elEnsureMigrated_();
    var sh = getELWeekSheet_(date, false);      // फक्त त्या आठवड्याची Sheet (सोमवार–रविवार) वाचली जाते
    var stMap = elStudentMap_();
    var out = [];
    (sh ? elReadAll_(sh) : []).forEach(function(r) {
      if (fmt(r[1]) !== date) return;
      if (iyatta && (r[2] || "").toString().trim() !== iyatta) return;
      if (tukdi && (r[3] || "").toString().trim() !== tukdi) return;
      out.push(elRowToObj_(r, stMap));
    });
    var absent = [];
    if (iyatta) absent = Object.keys(elAbsentSet_(date, iyatta, tukdi));
    var resEl = {status:"ok", date: date, records: elSort_(out), absentRegNos: absent};
    if (archivedFlagDate_(date)) resEl.archived = true;   // V19.53
    return wrap(cb, resEl);
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function doSaveEarlyLeave(d) {
  var role = d.requesterRole;
  if (role !== "super" && role !== "master" && role !== "teacher") return {status:"error", message:"नोंद करण्याचा अधिकार नाही."};
  var date = (d.date || todayStr()).toString().trim();
  if (!elDateValid_(date)) return {status:"error", message:"तारीख YYYY-MM-DD स्वरूपात हवी."};
  if (date > todayStr()) return {status:"error", message:"भविष्यातील तारखेची नोंद करता येत नाही."};
  var archErrEl = archivedWriteError_(date); if (archErrEl) return {status:"error", message: archErrEl};   // V19.53
  var iyatta = (d.iyatta || "").toString().trim();
  var tukdi = (d.tukdi || "").toString().trim();
  if (!iyatta || !tukdi) return {status:"error", message:"वर्ग व तुकडी निवडा."};
  var entries = safeParseJson(d.entriesJson), removes = safeParseJson(d.removeJson);
  if (!entries.length && !removes.length) return {status:"error", message:"जतन करण्यासाठी काहीही बदल नाही."};

  elEnsureMigrated_();
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(25000); locked = true;
    var roster = {};
    getRosterCached_(iyatta, tukdi, true).data.forEach(function(st) { roster[(st.regNo || "").toString().trim().toLowerCase()] = st; });
    var absent = elAbsentSet_(date, iyatta, tukdi);
    var sh = getELWeekSheet_(date, true);
    var all = elReadAll_(sh);
    var idxOf = {};
    for (var i = 0; i < all.length; i++) idxOf[fmt(all[i][1]) + "|" + (all[i][4] || "").toString().trim().toLowerCase()] = i;

    var ts = new Date().toLocaleString("en-IN");
    var updates = [], appends = [], dels = [], rejected = [];
    entries.forEach(function(e) {
      var key = (e.regNo || "").toString().trim().toLowerCase();
      var st = roster[key];
      if (!st) { rejected.push({regNo: e.regNo, reason: "विद्यार्थी या वर्ग-तुकडीत नाही"}); return; }
      if (absent[key]) { rejected.push({regNo: e.regNo, reason: "गैरहजर विद्यार्थ्याची नोंद करता येत नाही"}); return; }
      var type = (e.type || "").toString();
      var period = parseInt(e.period, 10);
      if (!EL_TYPES[type]) { rejected.push({regNo: e.regNo, reason: "प्रकार चुकीचा"}); return; }
      if (!(period >= 1 && period <= 9)) { rejected.push({regNo: e.regNo, reason: "तासिका क्रमांक १ ते ९ हवा"}); return; }
      var ix = idxOf[date + "|" + key];
      if (ix !== undefined) {
        var row = all[ix];
        var typeChanged = (row[7] || "").toString() !== type;
        row[0] = ts; row[7] = type; row[8] = period; row[9] = d.requesterUser || ""; row[10] = role;
        if (typeChanged) { row[11] = ""; row[12] = ""; row[13] = ""; }
        updates.push(ix);
      } else {
        appends.push([ts, date, iyatta, tukdi, st.regNo, st.studentId || "", st.fullName || "", type, period, d.requesterUser || "", role, "", "", ""]);
      }
    });
    removes.forEach(function(rg) {
      var ix = idxOf[date + "|" + (rg || "").toString().trim().toLowerCase()];
      if (ix !== undefined && (all[ix][2] || "").toString().trim() === iyatta && (all[ix][3] || "").toString().trim() === tukdi) dels.push(ix);
    });

    var delRows = dels.map(function(ix) { return all[ix].slice(0, 14); });
    if (delRows.length) backupDeleted_("EarlyLeave", EL_HEADERS, delRows, d.requesterUser || "", "शिक्षक/Master ने नोंद काढली");
    updates.forEach(function(ix) { sh.getRange(ix + 2, 1, 1, 14).setValues([all[ix]]); });
    if (appends.length) { ensureRows_(sh, sh.getLastRow() + appends.length); sh.getRange(sh.getLastRow() + 1, 1, appends.length, 14).setValues(appends); }
    dels.sort(function(a, b) { return b - a; }).forEach(function(ix) { sh.deleteRow(ix + 2); });

    var saved = updates.length + appends.length;
    if (!saved && !dels.length && rejected.length) return {status:"error", message: rejected[0].reason + (rejected.length > 1 ? " (+" + (rejected.length - 1) + " इतर)" : "")};
    return {rowIndex: 0, mode: "updated", count: saved, ref: iyatta + "-" + tukdi + " " + date, extra: {saved: saved, removed: dels.length, rejected: rejected}};
  } catch (err) {
    return {status:"error", message:err.toString()};
  } finally {
    if (locked) lock.releaseLock();
  }
}

// Peon: आजच्या "परवानगी" दिलेल्या विद्यार्थ्यांची यादी
function doGetEarlyLeavePeon(p, cb) {
  try {
    var role = p.requesterRole;
    if (role !== "peon" && role !== "super" && role !== "master") return wrap(cb, {status:"error", message:"अधिकार नाही."});
    var date = todayStr();
    elEnsureMigrated_();
    var stMap = elStudentMap_();
    var out = [];
    var elSh = getELWeekSheet_(date, false);    // आजच्या आठवड्याची Sheet (~३०० ओळी)
    (elSh ? elReadAll_(elSh) : []).forEach(function(r) {
      if (fmt(r[1]) !== date || (r[7] || "").toString() !== "permitted") return;
      out.push(elRowToObj_(r, stMap));
    });
    return wrap(cb, {status:"ok", date: date, records: elSort_(out)});
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// Peon: विद्यार्थी गेला / गेला नाही / पूर्ववत
function doMarkEarlyLeaveOut(d) {
  var role = d.requesterRole;
  if (role !== "peon" && role !== "super" && role !== "master") return {status:"error", message:"अधिकार नाही."};
  var status = (d.status || "").toString();
  if (["gone", "notgone", ""].indexOf(status) === -1) return {status:"error", message:"स्थिती चुकीची आहे."};
  var date = (role === "peon") ? todayStr() : (d.date || todayStr()).toString().trim();
  var key = (d.regNo || "").toString().trim().toLowerCase();
  if (!key) return {status:"error", message:"RegNo आवश्यक आहे."};
  var archErrMk = archivedWriteError_(date); if (archErrMk) return {status:"error", message: archErrMk};   // V19.53
  elEnsureMigrated_();
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(20000); locked = true;
    var sh = getELWeekSheet_(date, false);
    if (!sh) return {status:"error", message:"त्या दिवसाची नोंद सापडली नाही (शिक्षकाने ती बदलली/काढली असू शकते) — यादी Refresh करा."};
    var all = elReadAll_(sh);
    for (var i = 0; i < all.length; i++) {
      var r = all[i];
      if (fmt(r[1]) !== date || (r[4] || "").toString().trim().toLowerCase() !== key) continue;
      if ((r[7] || "").toString() !== "permitted") return {status:"error", message:"फक्त परवानगी दिलेल्या विद्यार्थ्याचीच मार्क-आउट नोंद करता येते."};
      var when = status ? new Date().toLocaleString("en-IN") : "";
      sh.getRange(i + 2, 12, 1, 3).setValues([[status, status ? (d.requesterUser || "") : "", when]]);
      return {rowIndex: i + 2, mode: "updated", count: 1, ref: (r[6] || "") + " " + (status || "reset") + " " + date, extra: {status: status}};
    }
    return {status:"error", message:"आजची नोंद सापडली नाही (शिक्षकाने ती बदलली/काढली असू शकते) — यादी Refresh करा."};
  } catch (err) {
    return {status:"error", message:err.toString()};
  } finally {
    if (locked) lock.releaseLock();
  }
}


// =====================================================================
// ===== V19.38 — Profile Share Card साठी फोटो (base64) =====
// Drive thumbnail वर CORS नसल्याने browser (html2canvas) तो Canvas वर घेऊ शकत नाही → server वरून data: URL म्हणून देतो.
// =====================================================================
function doGetPhotoData(p, cb) {
  try {
    var regNo = (p.regNo || "").toString().trim();
    if (!regNo) return wrap(cb, {status:"error", message:"regNo required"});
    var file = null;
    // १) Photo फोल्डरमधील <regNo>.jpg
    try {
      file = latestFileByName_(DriveApp.getFolderById(PHOTO_FOLDER_ID), cleanRegFileName(regNo) + ".jpg");
    } catch (e1) {}
    // २) Students मध्ये साठवलेल्या Photo URL मधील File ID
    if (!file) {
      var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
      var row = sh ? (findRowByKey(sh, 3, regNo) || findRowByKey(sh, 4, regNo)) : 0;
      var url = row ? (sh.getRange(row, PHOTO_COL).getValue() || "").toString() : "";
      var m = /[?&]id=([\w-]+)|\/d\/([\w-]+)/.exec(url);
      if (m) { try { file = DriveApp.getFileById(m[1] || m[2]); } catch (e2) {} }
    }
    if (!file) return wrap(cb, {status:"notfound", message:"फोटो सापडला नाही"});
    var blob = file.getBlob();
    var bytes = blob.getBytes();
    if (bytes.length > 700000) {   // मोठा फोटो — लहान Thumbnail वापरा म्हणजे Response हलका राहतो
      try { blob = file.getThumbnail(); bytes = blob.getBytes(); } catch (e3) {}
    }
    var mime = (blob.getContentType() || "image/jpeg").toString();
    if (mime.indexOf("image/") !== 0) mime = "image/jpeg";
    return wrap(cb, {status:"ok", dataUrl: "data:" + mime + ";base64," + Utilities.base64Encode(bytes)});
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}


// =====================================================================
// ===== V19.39 — वर्ग शिक्षक यादी (Sheet: "Class_Teacher") =====
// स्तंभ: A Class (5th-A) | B Class (5 वी) | C Div (अ) | D EMPLOYEE NAME | E कर्मचारी नाव | F MOB NO | G Emp Short Name
// पाहण्याचा अधिकार: Super Master, Master, Data Operator (deo), Peon
// =====================================================================
function doGetClassTeachers(p, cb) {
  try {
    var role = p.requesterRole;
    if (["super", "master", "deo", "peon"].indexOf(role) === -1) return wrap(cb, {status:"error", message:"वर्ग शिक्षक यादी पाहण्याचा अधिकार नाही."});
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("Class_Teacher");
    if (!sh) {
      ss.getSheets().forEach(function(x) { if (!sh && x.getName().toString().toLowerCase().replace(/[\s_\-]+/g, "") === "classteacher") sh = x; });
    }
    if (!sh) return wrap(cb, {status:"error", message:"\"Class_Teacher\" नावाची Sheet विद्यार्थी-माहितीच्या Google Sheet मध्ये सापडली नाही."});
    var n = sh.getLastRow() - 1;
    var out = [];
    if (n > 0) {
      var rows = sh.getRange(2, 1, n, 7).getValues();
      rows.forEach(function(r, idx) {
        var nameEn = (r[3] || "").toString().trim(), nameMr = (r[4] || "").toString().trim();
        if (!nameEn && !nameMr) return;
        var mob = r[5];
        mob = (typeof mob === "number") ? String(mob) : (mob || "").toString().trim();
        out.push({classKey: (r[0] || "").toString().trim(), classMr: (r[1] || "").toString().trim(), div: (r[2] || "").toString().trim(),
          nameEn: nameEn, nameMr: nameMr, mobile: mob, short: (r[6] || "").toString().trim(), _i: idx});
      });
      out.sort(function(a, b) {
        var c = classSortKey_(a.classKey.split("-")[0]) - classSortKey_(b.classKey.split("-")[0]);
        return c || (a._i - b._i);
      });
      out.forEach(function(o) { delete o._i; });
    }
    return wrap(cb, {status:"ok", data: out});
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}


// =====================================================================
// ===== V19.40 — टप्पा १: पाया (लहान Sheet, वेळेचे मोजमाप, साठवण-मीटर, Deleted Backup) =====
// =====================================================================

// ---------- १. लहान आकाराच्या Sheet ----------
// Google ची १ कोटी cells मर्यादा "वाटप केलेले" सर्व cells मोजते — नवीन रिकामी Sheet सुद्धा १०००×२६ = २६,००० cells घेते.
var GRID_MIN_ROWS = 300, GRID_COL_MARGIN = 3, GRID_GROW_CHUNK = 500, GRID_MIN_COLS_SHRINK = 15, SHEET_CELL_LIMIT = 10000000;

function trimNewSheet_(sh, nCols, minRows) {
  try {
    var rows = minRows || GRID_MIN_ROWS, cols = Math.max((nCols || 1) + GRID_COL_MARGIN, 1);
    // नवीन Sheet २६ स्तंभांची असते — जास्त स्तंभ हवे असल्यास वाढवा, कमी हवे असल्यास काढा
    if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
    else if (sh.getMaxColumns() > cols) sh.deleteColumns(cols + 1, sh.getMaxColumns() - cols);
    if (sh.getMaxRows() > rows) sh.deleteRows(rows + 1, sh.getMaxRows() - rows);
  } catch (e) { /* आकार कमी करता आला नाही तरी Sheet चालतेच */ }
}
// setValues ने शेवटच्या ओळीनंतर लिहिण्याआधी जागा वाढवा (appendRow स्वतः वाढवते)
function ensureRows_(sh, needLastRow) {
  var max = sh.getMaxRows();
  if (needLastRow > max) sh.insertRowsAfter(max, Math.max(needLastRow - max, GRID_GROW_CHUNK));
}
function ensureCols_(sh, needCols) {
  var max = sh.getMaxColumns();
  if (needCols > max) sh.insertColumnsAfter(max, needCols - max);
}

// ---------- २. वेळेचे मोजमाप ----------
// प्रत्येक विनंतीचा सर्व्हर-वेळ मोजतो. सर्व विनंत्या न नोंदवता ~३४% नमुना + २ सेकंदांहून जास्त (हळू) सर्व विनंत्या नोंदवल्या जातात
// म्हणजे मोजमापाचाच अतिरिक्त भार कमी राहतो.
var PERF_SLOW_MS = 2000, PERF_SAMPLE = 0.34, PERF_LOG_MAX = 400;
function doGet(e) {
  var t0 = Date.now();
  var out = doGetCore_(e);
  perfRecord_((e && e.parameter && e.parameter.action) || "", Date.now() - t0);
  return out;
}
function doPost(e) {
  var t0 = Date.now();
  var out = doPostCore_(e);
  perfRecord_((e && e.parameter && e.parameter.action) || "", Date.now() - t0);
  return out;
}
function perfRecord_(action, ms) {
  try {
    if (!action || action === "ping") return;
    var slow = ms >= PERF_SLOW_MS;
    if (!slow && Math.random() > PERF_SAMPLE) return;
    var cache = CacheService.getScriptCache();
    var raw = cache.get("perf_agg");
    var agg = raw ? JSON.parse(raw) : {};
    var a = agg[action] || {n: 0, sum: 0, max: 0, slow: 0};
    a.n++; a.sum += ms; a.last = ms;
    if (ms > a.max) a.max = ms;
    if (slow) a.slow++;
    agg[action] = a;
    cache.put("perf_agg", JSON.stringify(agg), 21600);
    if (slow) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName("Perf_Log");
      if (!sh) { sh = ss.insertSheet("Perf_Log"); trimNewSheet_(sh, 3); sh.appendRow(["Timestamp", "Action", "ServerMs"]); }
      sh.appendRow([new Date().toLocaleString("en-IN"), action, ms]);
      if (sh.getLastRow() > PERF_LOG_MAX) sh.deleteRows(2, 100);
    }
  } catch (err) { /* मोजमाप अपयशी झाले तरी मूळ विनंती चालतेच */ }
}

// ---------- ३. Deleted Backup (स्वतंत्र फाइल) ----------
// "SGS_Deleted_Backup" फाइल Drive मधील "SGS_Data" फोल्डरमध्ये; प्रत्येक नोंद-प्रकारासाठी स्वतंत्र Sheet: Deleted_Students, Deleted_Attendance, Deleted_EarlyLeave, ...
function getDataFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("DATA_FOLDER_ID");
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder("SGS_Data");
  props.setProperty("DATA_FOLDER_ID", f.getId());
  return f;
}
function getDeletedBackupSS_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("DELETED_BACKUP_ID");
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  var ss2 = SpreadsheetApp.create("SGS_Deleted_Backup");
  try { DriveApp.getFileById(ss2.getId()).moveTo(getDataFolder_()); } catch (e2) {}
  props.setProperty("DELETED_BACKUP_ID", ss2.getId());
  return ss2;
}
// type: "Students" | "Attendance" | "EarlyLeave" | "Fees" | "Diary" | "Transport" | "Scholarship"
// Backup यशस्वी न झाल्यास मुख्य फाइलमध्येच "Deleted_<type>" मध्ये ठेवतो — नोंद कधीही हरवू नये.
function backupDeleted_(type, headers, rows, by, reason) {
  if (!rows || !rows.length) return false;
  var allHeaders = ["DeletedAt", "DeletedBy", "Reason"].concat(headers);
  var ts = new Date().toLocaleString("en-IN");
  var out = rows.map(function(r) { return [ts, by || "", reason || ""].concat(r); });
  function write(ss) {
    var name = "Deleted_" + type;
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      trimNewSheet_(sh, allHeaders.length);
      sh.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);
      sh.setFrozenRows(1);
    }
    ensureRows_(sh, sh.getLastRow() + out.length); ensureCols_(sh, allHeaders.length);
    sh.getRange(sh.getLastRow() + 1, 1, out.length, allHeaders.length).setValues(out);
  }
  try {
    var bk = getDeletedBackupSS_();
    write(bk);
    try { var def = bk.getSheetByName("Sheet1"); if (def && bk.getSheets().length > 1) bk.deleteSheet(def); } catch (e0) {}
    return true;
  } catch (e) {
    try { write(SpreadsheetApp.getActiveSpreadsheet()); } catch (e2) {}
    return false;
  }
}
// Students ची बदलापूर्वीची प्रत (बढती/Archive/Roll No.) — Backup फाइलमध्ये; ते न जमल्यास मुख्य फाइलमध्ये लपवलेली.
function makeStudentsBackup_(sh, name, prefix) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    var bk = getDeletedBackupSS_();
    var cp = sh.copyTo(bk);
    cp.setName(name);
    pruneBackupSheets_(bk, prefix, 5);
    return {ss: bk, sheet: cp, external: true};
  } catch (e) {
    var cp2 = sh.copyTo(ss);
    cp2.setName(name);
    cp2.hideSheet();
    pruneBackupSheets_(ss, prefix, 3);
    return {ss: ss, sheet: cp2, external: false};
  }
}
function dropBackup_(b) { try { b.ss.deleteSheet(b.sheet); } catch (e) {} }

// ---------- ४. साठवण-मीटर, कामगिरी, रिकामी जागा कमी करणे (फक्त Super) ----------
function shrinkTarget_(sh) {
  return {rows: Math.max(sh.getLastRow() + GRID_MIN_ROWS, GRID_MIN_ROWS), cols: Math.max(sh.getLastColumn() + GRID_COL_MARGIN, GRID_MIN_COLS_SHRINK)};
}
function doGetSheetSizes(p, cb) {
  try {
    if (p.requesterRole !== "super") return wrap(cb, {status:"error", message:"अधिकार फक्त Super Master ला आहे."});
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var total = 0, savable = 0, list = [];
    ss.getSheets().forEach(function(sh) {
      var cells = sh.getMaxRows() * sh.getMaxColumns();
      var hidden = false; try { hidden = sh.isSheetHidden(); } catch (e) {}
      var save = 0;
      if (!hidden) { var t = shrinkTarget_(sh); save = Math.max(0, cells - Math.min(sh.getMaxRows(), t.rows) * Math.min(sh.getMaxColumns(), t.cols)); }
      total += cells; savable += save;
      list.push({name: sh.getName(), rows: sh.getLastRow(), cols: sh.getLastColumn(), cells: cells, hidden: hidden, savable: save});
    });
    list.sort(function(a, b) { return b.cells - a.cells; });

    var perf = [];
    try {
      var agg = JSON.parse(CacheService.getScriptCache().get("perf_agg") || "{}");
      Object.keys(agg).forEach(function(k) { perf.push({action: k, n: agg[k].n, avg: Math.round(agg[k].sum / agg[k].n), max: agg[k].max, slow: agg[k].slow}); });
      perf.sort(function(a, b) { return b.avg - a.avg; });
    } catch (e) {}
    var slowLog = [];
    var pl = ss.getSheetByName("Perf_Log");
    if (pl && pl.getLastRow() > 1) {
      var n = Math.min(pl.getLastRow() - 1, 25);
      slowLog = pl.getRange(pl.getLastRow() - n + 1, 1, n, 3).getValues().reverse().map(function(r) { return {time: r[0], action: r[1], ms: r[2]}; });
    }
    var backup = {configured: false};
    var bid = PropertiesService.getScriptProperties().getProperty("DELETED_BACKUP_ID");
    if (bid) {
      try {
        var bk = SpreadsheetApp.openById(bid), counts = {};
        bk.getSheets().forEach(function(sh) { if (sh.getName().indexOf("Deleted_") === 0) counts[sh.getName().slice(8)] = Math.max(0, sh.getLastRow() - 1); });
        backup = {configured: true, url: bk.getUrl(), counts: counts};
      } catch (e) { backup = {configured: true, error: "Backup फाइल उघडता आली नाही"}; }
    }
    return wrap(cb, {status:"ok", total: total, limit: SHEET_CELL_LIMIT, savable: savable, sheetCount: list.length, sheets: list.slice(0, 40), perf: perf, slow: slowLog, backup: backup});
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function doShrinkSheets(d) {
  if (d.requesterRole !== "super") return {status:"error", message:"अधिकार फक्त Super Master ला आहे."};
  if ((d.confirm || "") !== "SHRINK") return {status:"error", message:"खात्रीसाठी SHRINK टाइप करणे आवश्यक आहे."};
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(30000); locked = true;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var saved = 0, done = [], failed = [];
    ss.getSheets().forEach(function(sh) {
      try {
        if (sh.isSheetHidden()) return;          // लपवलेल्या (जुन्या डेटाच्या) Sheets ला हात लावत नाही
        var t = shrinkTarget_(sh), before = sh.getMaxRows() * sh.getMaxColumns();
        if (sh.getMaxColumns() > t.cols) sh.deleteColumns(t.cols + 1, sh.getMaxColumns() - t.cols);
        if (sh.getMaxRows() > t.rows) sh.deleteRows(t.rows + 1, sh.getMaxRows() - t.rows);
        var gain = before - sh.getMaxRows() * sh.getMaxColumns();
        if (gain > 0) { saved += gain; done.push(sh.getName()); }
      } catch (e) { failed.push(sh.getName()); }
    });
    return {rowIndex: 0, mode: "updated", count: done.length, ref: "shrink " + saved, extra: {saved: saved, sheets: done, failed: failed}};
  } catch (err) {
    return {status:"error", message:err.toString()};
  } finally {
    if (locked) lock.releaseLock();
  }
}


// =====================================================================
// ===== V19.42 — टप्पा ३: वार्षिक Sheets (Fees + FeesSummary, Diary, Transport) =====
// =====================================================================
// शैक्षणिक वर्ष जून–मे: 2026-06-15 → "2026-27"; 2027-03-10 → "2026-27"
function acYearOfDate_(dateStr) {
  var m = ISO_DATE_RE.exec((dateStr || "").toString());
  if (!m) return currentAcYear_();
  var y = parseInt(m[1], 10), mo = parseInt(m[2], 10);
  var start = mo >= 6 ? y : y - 1;
  return start + "-" + String(start + 1).slice(-2);
}
function currentAcYear_() {
  try {
    var v = PropertiesService.getScriptProperties().getProperty("CURRENT_AC_YEAR");
    if (v && /^\d{4}-\d{2}$/.test(v)) return v;
  } catch (e) {}
  return acYearOfDate_(todayStr());
}
function normAcYear_(v) {
  v = (v || "").toString().trim();
  if (/^\d{4}-\d{2}$/.test(v)) return v;
  var m = /^(\d{4})-(\d{4})$/.exec(v);
  return m ? (m[1] + "-" + m[2].slice(2)) : currentAcYear_();
}

function getYearSheet_(base, year, headers, create, textCol) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = base + "_" + year;
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  if (!create) return archiveSheet_(name);   // V19.53: Archive झालेल्या वर्षाची Sheet (फक्त वाचन)
  sh = ss.insertSheet(name);
  trimNewSheet_(sh, headers.length);
  sh.appendRow(headers);
  if (textCol) sh.getRange(1, textCol, 1, 1).setNumberFormat("@");
  sh.setFrozenRows(1);
  return sh;
}
function listYearSheets_(base) {
  var re = new RegExp("^" + base + "_\\d{4}-\\d{2}$");
  var list = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(function(sh) { return re.test(sh.getName()); });
  list.sort(function(a, b) { return a.getName() < b.getName() ? 1 : (a.getName() > b.getName() ? -1 : 0); });   // नवीन वर्ष आधी
  return list;
}

// एका स्तंभातील मूल्य (उदा. RegNo) असलेल्या ओळींचे क्रमांक — TextFinder ने (मोठी Sheet वाचल्याशिवाय); न चालल्यास स्तंभ वाचतो
function findRowsByValue_(sh, col, key) {
  key = (key || "").toString().trim();
  if (!key || !sh || sh.getLastRow() < 2) return [];
  var n = sh.getLastRow() - 1;
  try {
    var found = sh.getRange(2, col, n, 1).createTextFinder(key).matchEntireCell(true).findAll();
    return found.map(function(r) { return r.getRow(); }).sort(function(a, b) { return a - b; });
  } catch (e) {
    var vals = sh.getRange(2, col, n, 1).getValues(), out = [], lk = key.toLowerCase();
    for (var i = 0; i < vals.length; i++) if ((vals[i][0] || "").toString().trim().toLowerCase() === lk) out.push(i + 2);
    return out;
  }
}
function readRowsAt_(sh, rows, ncols) {
  if (!rows.length) return [];
  if (rows.length <= 25) return rows.map(function(r) { return sh.getRange(r, 1, 1, ncols).getValues()[0]; });
  var lo = rows[0], hi = rows[rows.length - 1];
  var block = sh.getRange(lo, 1, hi - lo + 1, ncols).getValues();
  return rows.map(function(r) { return block[r - lo]; });
}

// ---------- जुन्या एकाच Sheet चे वार्षिक Sheets मध्ये एकदाचे स्थलांतर ----------
// जुनी Sheet रिकामी असेल तर काढली जाते; नोंदी असतील तर "<नाव>_OLD_<वेळ>" नावाने लपवली जाते (काहीही हरवत नाही). Lock स्वतः घेतो — दुसऱ्या Lock आत बोलावू नये.
function migrateLegacyToYearly_(spec) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(spec.flag) === "1") return;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (props.getProperty(spec.flag) === "1") return;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var old = ss.getSheetByName(spec.legacy);
    if (old) {
      var n = old.getLastRow() - 1;
      if (n > 0) {
        var groups = {};
        old.getRange(2, 1, n, spec.ncols).getValues().forEach(function(r) {
          if (!r.some(function(v) { return v !== "" && v !== null; })) return;
          var y = spec.yearOf(r);
          (groups[y] = groups[y] || []).push(r);
        });
        Object.keys(groups).forEach(function(y) {
          var sh = spec.target(y), rows = groups[y];
          ensureRows_(sh, sh.getLastRow() + rows.length);
          sh.getRange(sh.getLastRow() + 1, 1, rows.length, spec.ncols).setValues(rows);
        });
        if (spec.after) spec.after(Object.keys(groups));
        old.setName(spec.legacy + "_OLD_" + Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyyMMdd_HHmm"));
        try { old.hideSheet(); } catch (e) {}
      } else {
        try { if (ss.getSheets().length > 1) ss.deleteSheet(old); } catch (e) {}
      }
    }
    props.setProperty(spec.flag, "1");
  } finally {
    lock.releaseLock();
  }
}

// ---------- Fees ----------
var FEES_SUMMARY_HEADERS = ["RegNo","StudentId","FullName","Class","Division","AcYear","FeePaid","LastPaidAt","LastAmount","Installments","UpdatedBy","OpeningDue"];
function feesEventSheet_(year, create) { return getYearSheet_("Fees", year, FEES_HEADERS, create); }
function feesSummarySheet_(year, create) { return getYearSheet_("FeesSummary", year, FEES_SUMMARY_HEADERS, create); }
function feesEnsureMigrated_() {
  migrateLegacyToYearly_({legacy: "Fees", flag: "FEES_YEARLY_MIGRATED", ncols: 11,
    yearOf: function(r) { return normAcYear_(r[6]); },
    target: function(y) { return feesEventSheet_(y, true); },
    after: function(years) { years.forEach(rebuildFeesSummary_); }});
}
// सारांश Sheet मिळवा; ती नसेल पण जमा-नोंदी असतील तर त्यावरून आपोआप पुन्हा बांधा
function feesSummaryEnsure_(year) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FeesSummary_" + year);
  if (sh) return sh;
  var ev = feesEventSheet_(year, false);
  if (ev && ev.getLastRow() > 1) { rebuildFeesSummary_(year); return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FeesSummary_" + year); }
  return null;
}
function feesSummaryRows_(year) {
  var sh = feesSummaryEnsure_(year);
  if (!sh || sh.getLastRow() < 2) return [];
  ensureCols_(sh, 12);
  return sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues();
}
function rebuildFeesSummary_(year) {
  var ev = feesEventSheet_(year, false);
  var sum = feesSummarySheet_(year, true);
  var latest = {}, order = [];
  if (ev && ev.getLastRow() > 1) {
    ev.getRange(2, 1, ev.getLastRow() - 1, 11).getValues().forEach(function(r) {
      var rg = (r[1] || "").toString().trim().toLowerCase();
      if (!rg) return;
      var e = latest[rg];
      if (!e) { e = latest[rg] = {n: 0, prev: 0, opening: 0}; order.push(rg); }
      if ((r[10] || "").toString() === "OPENING") {        // वर्ष-बदलातील मागील वर्षाची बाकी (जमा नाही)
        e.opening += parseFloat(r[9]) || 0;
        if (!e.row) e.row = r;
        return;
      }
      var paid = parseFloat(r[8]) || 0;
      e.row = r; e.n++; e.last = paid - e.prev; e.prev = paid;
    });
  }
  var out = order.map(function(rg) {
    var e = latest[rg], r = e.row;
    return [r[1], r[2], r[3], r[4], r[5], year, e.prev, e.n ? r[0] : "", e.last || 0, e.n, e.n ? r[10] : "OPENING", e.opening || 0];
  });
  ensureCols_(sum, 12);
  try { if (!sum.getRange(1, 12).getValue()) sum.getRange(1, 12).setValue("OpeningDue"); } catch (eh) {}
  if (sum.getLastRow() > 1) sum.getRange(2, 1, sum.getLastRow() - 1, 12).clearContent();
  if (out.length) { ensureRows_(sum, 1 + out.length); sum.getRange(2, 1, out.length, 12).setValues(out); }
  return out.length;
}
// Super Admin: सारांश चुकीचा वाटल्यास जमा-नोंदींवरून पुन्हा बांधा (जमा-नोंदी हीच मूळ माहिती आहे)
function doRebuildFeesSummary(d) {
  if (d.requesterRole !== "super") return {status:"error", message:"अधिकार फक्त Super Master ला आहे."};
  var year = normAcYear_(d.year);
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    feesEnsureMigrated_();
    lock.waitLock(30000); locked = true;
    var n = rebuildFeesSummary_(year);
    return {rowIndex: 0, mode: "updated", count: n, ref: "feesSummary " + year, extra: {year: year, students: n}};
  } catch (err) {
    return {status:"error", message:err.toString()};
  } finally {
    if (locked) lock.releaseLock();
  }
}

// ---------- Diary / Transport ----------
function diarySheet_(year, create) { return getYearSheet_("Diary", year, DIARY_HEADERS, create, 2); }
function diaryEnsureMigrated_() {
  migrateLegacyToYearly_({legacy: "DiaryRemarks", flag: "DIARY_YEARLY_MIGRATED", ncols: 10,
    yearOf: function(r) { return acYearOfDate_(fmt(r[1])); },
    target: function(y) { return diarySheet_(y, true); }});
}
function transportEnsureMigrated_() {
  migrateLegacyToYearly_({legacy: "StudentTransport", flag: "TRANSPORT_YEARLY_MIGRATED", ncols: 10,
    yearOf: function() { return currentAcYear_(); },
    target: function(y) { return getYearSheet_("Transport", y, TRANSPORT_HEADERS, true); }});
}


// =====================================================================
// ===== V19.43 — टप्पा ४: हजेरी (AttMarks, Cache, वर्गनिहाय सारांश) =====
// =====================================================================
function phoneDigits_(v) {
  if (v === null || v === undefined) return "";
  return (typeof v === "number" ? String(Math.round(v)) : v.toString()).replace(/[^0-9]/g, "");
}

// हजेरीची "आवृत्ती" — प्रत्येक हजेरी Save वर बदलते; ती Cache keys मध्ये असल्याने जुना Cache आपोआप निरुपयोगी होतो
function attVer_() {
  try {
    var c = CacheService.getScriptCache(), v = c.get("att_ver");
    if (!v) { v = String(Date.now()); c.put("att_ver", v, 21600); }
    return v;
  } catch (e) { return "0"; }
}
function attBump_() { try { CacheService.getScriptCache().put("att_ver", String(Date.now()), 21600); } catch (e) {} }

// ---- AttMarks_YYYY_MM: प्रत्येक वर्ग-तुकडीची प्रत्येक दिवसाची हजेरी-Save ची एक ओळ (विद्यार्थी, गैरहजर, हजर, कोणी, केव्हा) ----
var ATT_MARKS_HEADERS = ["Date","Class","Division","Total","Absent","Present","MarkedBy","SavedAt"];
function attMarksSheet_(date, create) {
  var m = ISO_DATE_RE.exec((date || "").toString());
  if (!m) return null;
  var ss = SpreadsheetApp.getActiveSpreadsheet(), name = "AttMarks_" + m[1] + "_" + m[2];
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  if (!create) return archiveSheet_(name);   // V19.53
  sh = ss.insertSheet(name);
  trimNewSheet_(sh, ATT_MARKS_HEADERS.length);
  sh.appendRow(ATT_MARKS_HEADERS);
  sh.getRange("A:A").setNumberFormat("@");
  sh.setFrozenRows(1);
  return sh;
}
function attMarksRowsForDate_(date) {            // null = त्या महिन्याची AttMarks Sheet अस्तित्वातच नाही
  var sh = attMarksSheet_(date, false);
  if (!sh) return null;
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().filter(function(r) { return fmt(r[0]) === date; });
}
function attMarksMonthRows_(monthKey) {          // monthKey = "YYYY_MM"
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AttMarks_" + monthKey) || archiveSheet_("AttMarks_" + monthKey);   // V19.53
  if (!sh) return null;
  return sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
}
function attMarksUpsert_(date, iyatta, tukdi, absentCount, by) {
  var sh = attMarksSheet_(date, true);
  var total = getRosterCached_(iyatta, tukdi, false).data.length;
  var last = sh.getLastRow(), rowIx = 0;
  if (last > 1) {
    var vals = sh.getRange(2, 1, last - 1, 3).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (fmt(vals[i][0]) === date && (vals[i][1] || "").toString().trim() === iyatta && (vals[i][2] || "").toString().trim() === tukdi) { rowIx = i + 2; break; }
    }
  }
  var row = [date, iyatta, tukdi, total, absentCount, Math.max(0, total - absentCount), by || "", new Date().toLocaleString("en-IN")];
  if (rowIx) sh.getRange(rowIx, 1, 1, 8).setValues([row]); else sh.appendRow(row);
  attBump_();
}

// एका तारखेच्या सर्व गैरहजर नोंदी — महिन्याची मोठी Sheet पूर्ण न वाचता (TextFinder). जुनी "DailyAttendance" Sheet असल्यास ती पूर्ण वाचतो.
function readAttendanceDateRows_(date) {
  var out = [];
  attSheetsForRead_(date, date).forEach(function(sh) {
    if (sh.getName() === ATT_LEGACY_SHEET) {
      if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().forEach(function(r) { if (fmt(r[1]) === date) out.push(r); });
    } else {
      out = out.concat(readRowsAt_(sh, findRowsByValue_(sh, 2, date), 9));
    }
  });
  return out;
}

// ---- Master/Super: सर्व वर्गांचा महिन्याचा उपस्थिती सारांश (फक्त AttMarks वाचतो) ----
function doGetAttendanceOverview(p, cb) {
  try {
    if (p.requesterRole !== "super" && p.requesterRole !== "master") return wrap(cb, {status:"error", message:"हे पाहण्याचा अधिकार Super Master / Master ला आहे."});
    var month = (p.month || todayStr().slice(0, 7)).toString().trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return wrap(cb, {status:"error", message:"महिना YYYY-MM स्वरूपात हवा."});
    var today = todayStr();
    var ckey = "attov_" + month + "_" + today + "_" + attVer_() + "_" + rosterVer_();
    var hit = cacheGetBig_(ckey);
    if (hit) { hit.cached = true; return wrap(cb, hit); }

    var students = {};
    getRosterCached_("", "", false).data.forEach(function(st) {
      var k = (st.iyatta || "").toString().trim() + "|" + (st.tukdi || "").toString().trim();
      if (!st.iyatta) return;
      students[k] = (students[k] || 0) + 1;
    });
    var agg = {}, todayMark = {};
    var rows = attMarksMonthRows_(month.replace("-", "_"));
    (rows || []).forEach(function(m) {
      var d = fmt(m[0]);
      var k = (m[1] || "").toString().trim() + "|" + (m[2] || "").toString().trim();
      var a = agg[k] || (agg[k] = {days: {}, total: 0, absent: 0, last: ""});
      a.days[d] = true;
      a.total += parseInt(m[3], 10) || 0;
      a.absent += parseInt(m[4], 10) || 0;
      if (d > a.last) a.last = d;
      if (d === today) todayMark[k] = {absent: parseInt(m[4], 10) || 0};
    });
    var out = Object.keys(students).map(function(k) {
      var parts = k.split("|"), a = agg[k];
      return {iyatta: parts[0], tukdi: parts[1], students: students[k],
        days: a ? Object.keys(a.days).length : 0,
        avgPct: (a && a.total) ? Math.round((a.total - a.absent) * 1000 / a.total) / 10 : null,
        absentDays: a ? a.absent : 0, lastDate: a ? a.last : "",
        todayMarked: !!todayMark[k], todayAbsent: todayMark[k] ? todayMark[k].absent : null};
    });
    out.sort(function(a, b) { return classSortKey_(a.iyatta) - classSortKey_(b.iyatta) || (a.tukdi < b.tukdi ? -1 : (a.tukdi > b.tukdi ? 1 : 0)); });
    var res = {status:"ok", month: month, rows: out, hasMarks: rows !== null && rows.length > 0};
    if (archivedFlagMonth_(month)) res.archived = true;   // V19.53
    cachePutBig_(ckey, res, 300);
    return wrap(cb, res);
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}


// =====================================================================
// ===== V19.44 — टप्पा ५: शैक्षणिक वर्ष बदल व Archive =====
// =====================================================================
// क्रम: (१) 10th Pass-out Archive (२) वर्ग-बढती (३) "नवे वर्ष सुरू करा" (४) तपासणी (५) "जुने वर्ष Archive करा"
function acYearStart_(y) { return parseInt((y || "").toString().slice(0, 4), 10) || 0; }
function nextAcYear_(y) { var s = acYearStart_(y) + 1; return s + "-" + String(s + 1).slice(-2); }
// कॅलेंडरनुसार नवे शैक्षणिक वर्ष सुरू होऊन (१५ जूननंतर) "चालू वर्ष" अजून जुनेच असेल तर true
function needsRollover_() {
  var cur = currentAcYear_(), calc = acYearOfDate_(todayStr());
  if (acYearStart_(calc) <= acYearStart_(cur)) return false;
  return todayStr() >= (acYearStart_(calc) + "-06-15");
}
function archivedYears_() {
  var v = PropertiesService.getScriptProperties().getProperty("ARCHIVED_YEARS") || "";
  return v ? v.split(",").filter(function(x) { return /^\d{4}-\d{2}$/.test(x); }) : [];
}
function getArchiveSS_(year, create) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("ARCHIVE_ID_" + year);
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  if (!create) return null;
  var ss2 = SpreadsheetApp.create("SGS_Archive_" + year);
  try { DriveApp.getFileById(ss2.getId()).moveTo(getDataFolder_()); } catch (e2) {}
  props.setProperty("ARCHIVE_ID_" + year, ss2.getId());
  try { props.setProperty("ARCHIVE_URL_" + year, ss2.getUrl()); } catch (e3) {}
  return ss2;
}
// चालू फाइलमध्ये नसेल तर त्या वर्षाच्या Archive फाइलमधून
function findYearSheet_(base, year) {
  var sh = getYearSheet_(base, year, null, false);
  if (sh) return sh;
  var ar = getArchiveSS_(year, false);
  return ar ? ar.getSheetByName(base + "_" + year) : null;
}

// =====================================================================
// ===== V19.53 — Archive झालेल्या वर्षांची माहिती अॅपमध्ये फक्त पाहणे =====
// • वाचन-मार्गात: चालू फाइलमध्ये Sheet नसेल आणि ती Archive झालेल्या वर्षाची असेल तर त्या वर्षाच्या SGS_Archive_<वर्ष> फाइलमधून वाचतो.
// • लेखन-मार्गात: Archive झालेल्या वर्षात/तारखेत कोणतीही नोंद/बदल/मिटवणे नाकारले जाते (नाहीतर मुख्य फाइलमध्ये अर्धवट दुबार Sheet बनेल).
// • प्रतिसादात फक्त Archive असल्यास "archived: true" (क्लायंट "📦 Archive — फक्त पाहणे" दाखवतो); इतर वेळी प्रतिसाद पूर्वीसारखाच.
// =====================================================================
var _archMemo = {};
function isArchivedYear_(y) { return !!y && archivedYears_().indexOf(String(y)) !== -1; }
function archiveSS_(year) {
  if (_archMemo[year]) return _archMemo[year];
  var ar = null;
  try { ar = getArchiveSS_(year, false); } catch (e) { ar = null; }
  if (ar) _archMemo[year] = ar;
  return ar;
}
function archiveSheet_(name) {
  var y = sheetYearOf_(name);
  if (!y || !isArchivedYear_(y)) return null;
  var ar = archiveSS_(y);
  return ar ? ar.getSheetByName(name) : null;
}
function archivedFlagYear_(y) { return isArchivedYear_((y || "").toString().trim()); }
function archivedFlagDate_(d) { d = (d || "").toString().trim(); return /^\d{4}-\d{2}-\d{2}$/.test(d) && isArchivedYear_(acYearOfDate_(d)); }
function archivedFlagMonth_(m) { m = (m || "").toString().trim(); return /^\d{4}-(0[1-9]|1[0-2])$/.test(m) && isArchivedYear_(acYearOfDate_(m + "-15")); }
function archivedMsg_(y) { return "📦 " + y + " ची माहिती Archive झाली आहे — ती फक्त पाहता येते, बदलता येत नाही."; }
// लेखनापूर्वी: वर्ष (YYYY-YY) किंवा तारीख (YYYY-MM-DD) Archive झालेल्या वर्षाची असल्यास त्रुटी-संदेश, नाहीतर ""
function archivedWriteError_(v) {
  v = (v || "").toString().trim();
  var y = "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(v)) y = acYearOfDate_(v);
  else if (/^\d{4}-\d{2}$/.test(v)) y = v;
  return (y && isArchivedYear_(y)) ? archivedMsg_(y) : "";
}
// दिलेल्या तारीख-मर्यादेच्या Archive झालेल्या वर्षांतील महिन्यांच्या हजेरी Sheets (फक्त वाचनासाठी; मर्यादा न दिल्यास कधीही नाही — पूर्ण स्कॅन टाळण्यासाठी)
function attArchiveSheets_(fromDate, toDate, skipNames) {
  var fromKey = ISO_DATE_RE.test(fromDate || "") ? fromDate.slice(0, 7).replace("-", "_") : "";
  var toKey = ISO_DATE_RE.test(toDate || "") ? toDate.slice(0, 7).replace("-", "_") : "";
  var out = [];
  if (!fromKey && !toKey) return out;
  archivedYears_().forEach(function(y) {
    var ys = acYearStart_(y), yFrom = ys + "_06", yTo = (ys + 1) + "_05";
    if (fromKey && yTo < fromKey) return;
    if (toKey && yFrom > toKey) return;
    var ar = archiveSS_(y);
    if (!ar) return;
    ar.getSheets().forEach(function(sh) {
      var m = ATT_SHEET_RE.exec(sh.getName());
      if (!m) return;
      var k = m[1] + "_" + m[2];
      if (k < yFrom || k > yTo || (fromKey && k < fromKey) || (toKey && k > toKey)) return;
      if (skipNames && skipNames[sh.getName()]) return;
      out.push(sh);
    });
  });
  return out;
}
function attSheetsForRead_(fromDate, toDate) {
  var main = attMonthSheets_(fromDate, toDate), names = {};
  main.forEach(function(sh) { names[sh.getName()] = true; });
  return main.concat(attArchiveSheets_(fromDate, toDate, names));
}

// एका शैक्षणिक वर्षाच्या (जून–मे) सर्व Sheets: वार्षिक + महिनावार हजेरी + AttMarks + आठवडानिहाय घरी-जाणारे
function sheetsOfYear_(ss, year) {
  var y = acYearStart_(year), from = y + "-06", to = (y + 1) + "-05";
  var out = [];
  ss.getSheets().forEach(function(sh) {
    var n = sh.getName(), m;
    if (n === "Fees_" + year || n === "FeesSummary_" + year || n === "Diary_" + year || n === "Transport_" + year || n === "Scholarship_" + year) { out.push(sh); return; }
    if ((m = /^(Attendance|AttMarks|Log)_(\d{4})_(\d{2})$/.exec(n))) { var k = m[2] + "-" + m[3]; if (k >= from && k <= to) out.push(sh); return; }
    if ((m = /^EL_(\d{4}-\d{2}-\d{2})$/.exec(n))) { if (m[1] >= y + "-06-01" && m[1] <= (y + 1) + "-05-31") out.push(sh); }
  });
  return out;
}
function sheetYearOf_(name) {
  var m;
  if ((m = /^(?:Fees|FeesSummary|Diary|Transport|Scholarship)_(\d{4}-\d{2})$/.exec(name))) return m[1];
  if ((m = /^(?:Attendance|AttMarks|Log)_(\d{4})_(\d{2})$/.exec(name))) return acYearOfDate_(m[1] + "-" + m[2] + "-15");
  if ((m = /^EL_(\d{4}-\d{2}-\d{2})$/.exec(name))) return acYearOfDate_(m[1]);
  return "";
}

// मागील वर्षाची बाकी = (सुरुवातीची बाकी + एकूण फी − जमा), फक्त सध्या Students मध्ये असलेल्या विद्यार्थ्यांसाठी
function computeCarryDues_(prevYear) {
  var total = getTotalFeePerStudent(), map = {};
  feesSummaryRows_(prevYear).forEach(function(r) {
    var rg = (r[0] || "").toString().trim().toLowerCase();
    if (rg) map[rg] = {paid: parseFloat(r[6]) || 0, opening: parseFloat(r[11]) || 0};
  });
  var list = [], sum = 0;
  getRosterCached_("", "", true).data.forEach(function(st) {
    var rg = (st.regNo || "").toString().trim().toLowerCase();
    if (!rg) return;
    var m = map[rg] || {paid: 0, opening: 0};
    var due = Math.max(0, m.opening + total - m.paid);
    if (due > 0) { list.push({regNo: st.regNo, studentId: st.studentId || "", fullName: st.fullName || "", iyatta: st.iyatta, tukdi: st.tukdi, due: due}); sum += due; }
  });
  return {list: list, total: sum, feePerStudent: total};
}

function doGetYearStatus(p, cb) {
  try {
    if (p.requesterRole !== "super") return wrap(cb, {status:"error", message:"अधिकार फक्त Super Master ला आहे."});
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cur = currentAcYear_();
    var byYear = {};
    getRosterCached_("", "", false).data.forEach(function(st) { var y = (st.acYear || "").toString().trim() || "—"; byYear[y] = (byYear[y] || 0) + 1; });
    var years = {};
    ss.getSheets().forEach(function(sh) {
      var y = sheetYearOf_(sh.getName());
      if (!y) return;
      var o = years[y] || (years[y] = {year: y, sheets: [], cells: 0});
      o.sheets.push({name: sh.getName(), rows: sh.getLastRow(), cells: sh.getMaxRows() * sh.getMaxColumns()});
      o.cells += sh.getMaxRows() * sh.getMaxColumns();
    });
    var props = PropertiesService.getScriptProperties();
    var archives = archivedYears_().map(function(y) { return {year: y, url: props.getProperty("ARCHIVE_URL_" + y) || ""}; });
    var dues = computeCarryDues_(cur);
    return wrap(cb, {status:"ok", current: cur, suggested: nextAcYear_(cur), calendarYear: acYearOfDate_(todayStr()), needsRollover: needsRollover_(),
      studentsByYear: byYear, dues: {count: dues.list.length, total: dues.total, feePerStudent: dues.feePerStudent},
      years: Object.keys(years).sort().reverse().map(function(k) { return years[k]; }), archives: archives});
  } catch (err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function doStartNewYear(d) {
  if (d.requesterRole !== "super") return {status:"error", message:"अधिकार फक्त Super Master ला आहे."};
  if ((d.confirm || "") !== "NEWYEAR") return {status:"error", message:"खात्रीसाठी NEWYEAR टाइप करणे आवश्यक आहे."};
  var newYear = (d.newYear || "").toString().trim();
  if (!/^\d{4}-\d{2}$/.test(newYear)) return {status:"error", message:"नवे वर्ष 2027-28 या स्वरूपात टाका."};
  var prev = currentAcYear_();
  if (acYearStart_(newYear) <= acYearStart_(prev)) return {status:"error", message:"नवे वर्ष चालू वर्षापेक्षा (" + prev + ") पुढचे हवे."};
  if (acYearStart_(newYear) > acYearStart_(prev) + 1) return {status:"error", message:"एकावेळी फक्त पुढचेच वर्ष (" + nextAcYear_(prev) + ") सुरू करता येते."};
  var carry = (d.carry || "1") !== "0";
  feesEnsureMigrated_(); diaryEnsureMigrated_(); transportEnsureMigrated_();
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(30000); locked = true;
    var dues = carry ? computeCarryDues_(prev) : {list: [], total: 0, feePerStudent: getTotalFeePerStudent()};
    var ev = feesEventSheet_(newYear, true);
    feesSummarySheet_(newYear, true);
    diarySheet_(newYear, true);
    transportSheetForWrite_(newYear);                         // मागील वर्षाची प्रवास-नोंदींची प्रत
    var existing = {};
    if (ev.getLastRow() > 1) ev.getRange(2, 1, ev.getLastRow() - 1, 11).getValues().forEach(function(r) { if ((r[10] || "").toString() === "OPENING") existing[(r[1] || "").toString().trim().toLowerCase()] = true; });
    var ts = new Date().toLocaleString("en-IN");
    var rows = dues.list.filter(function(x) { return !existing[(x.regNo || "").toString().trim().toLowerCase()]; })
      .map(function(x) { return [ts, x.regNo, x.studentId, x.fullName, x.iyatta, x.tukdi, newYear, dues.feePerStudent, 0, x.due, "OPENING"]; });
    if (rows.length) { ensureRows_(ev, ev.getLastRow() + rows.length); ev.getRange(ev.getLastRow() + 1, 1, rows.length, 11).setValues(rows); }
    rebuildFeesSummary_(newYear);
    PropertiesService.getScriptProperties().setProperty("CURRENT_AC_YEAR", newYear);
    invalidateRosterCache_(); attBump_();
    return {rowIndex: 0, mode: "updated", count: rows.length, ref: "newYear " + prev + " → " + newYear,
      extra: {prev: prev, newYear: newYear, carried: rows.length, carriedTotal: dues.total}};
  } catch (err) {
    return {status:"error", message:err.toString()};
  } finally {
    if (locked) lock.releaseLock();
  }
}

function doArchiveYear(d) {
  if (d.requesterRole !== "super") return {status:"error", message:"अधिकार फक्त Super Master ला आहे."};
  if ((d.confirm || "") !== "ARCHIVEYEAR") return {status:"error", message:"खात्रीसाठी ARCHIVEYEAR टाइप करणे आवश्यक आहे."};
  var year = (d.year || "").toString().trim();
  if (!/^\d{4}-\d{2}$/.test(year)) return {status:"error", message:"वर्ष 2026-27 या स्वरूपात टाका."};
  var cur = currentAcYear_();
  if (acYearStart_(year) >= acYearStart_(cur)) return {status:"error", message:"चालू (" + cur + ") किंवा पुढचे वर्ष Archive करता येत नाही."};
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(30000); locked = true;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var list = sheetsOfYear_(ss, year);
    if (!list.length) return {status:"error", message:year + " ची Archive करण्यासारखी कोणतीही Sheet मुख्य फाइलमध्ये नाही."};
    var ar = getArchiveSS_(year, true);
    var moved = [], skipped = [], failed = [], freed = 0;
    list.forEach(function(sh) {
      var name = sh.getName();
      try {
        // प्रवासाची Sheet तेव्हाच हलवा जेव्हा चालू वर्षाची प्रत मुख्य फाइलमध्ये आहे (नाहीतर वाचन अडेल)
        if (name === "Transport_" + year && !ss.getSheetByName("Transport_" + cur)) { skipped.push(name + " (चालू वर्षाची Transport Sheet अजून नाही)"); return; }
        var target = ar.getSheetByName(name);
        if (target && target.getLastRow() !== sh.getLastRow()) { ar.deleteSheet(target); target = null; }   // अर्धवट आधीची प्रत — पुन्हा करा
        if (!target) {
          var cp = sh.copyTo(ar);
          cp.setName(name);
          if (cp.getLastRow() !== sh.getLastRow()) { ar.deleteSheet(cp); failed.push(name); return; }   // पडताळणी अयशस्वी — मूळ Sheet तशीच
        }
        freed += sh.getMaxRows() * sh.getMaxColumns();
        ss.deleteSheet(sh);
        moved.push(name);
      } catch (e1) { failed.push(name); }
    });
    try { var def = ar.getSheetByName("Sheet1"); if (def && ar.getSheets().length > 1) ar.deleteSheet(def); } catch (e2) {}
    if (moved.length) {
      var yrs = archivedYears_();
      if (yrs.indexOf(year) === -1) { yrs.push(year); yrs.sort(); PropertiesService.getScriptProperties().setProperty("ARCHIVED_YEARS", yrs.join(",")); }
      attBump_(); invalidateRosterCache_();   // V19.53: Archive नंतर जुने (cache केलेले) हजेरी-उत्तर वापरले जाऊ नये — "archived" पट्टी लगेच दिसावी
    }
    return {rowIndex: 0, mode: "updated", count: moved.length, ref: "archiveYear " + year,
      extra: {year: year, moved: moved, skipped: skipped, failed: failed, freedCells: freed, url: ar.getUrl()}};
  } catch (err) {
    return {status:"error", message:err.toString()};
  } finally {
    if (locked) lock.releaseLock();
  }
}
