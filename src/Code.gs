/**
 * Private Tutor Management System - Backend Engine (Code.gs)
 * Built for Google Apps Script bound to Google Sheets
 */

// ==========================================
// CONFIGURATION & SETUP
// ==========================================

var SHEET_NAMES = {
  GROUPS: 'Groups',
  STUDENTS: 'Students',
  ATTENDANCE: 'Attendance',
  EXAMS: 'Exams',
  EXAM_RESULTS: 'ExamResults',
  SETTINGS: 'Settings'
};

var HEADERS = {
  GROUPS: ['GroupID', 'GroupName', 'Grade', 'StudyType', 'Days', 'Time', 'CreatedAt'],
  STUDENTS: ['StudentID', 'GroupID', 'StudentName', 'ParentPhone', 'AttendanceCounter', 'PaymentStatus', 'CreatedAt'],
  ATTENDANCE: ['AttendanceID', 'Date', 'GroupID', 'StudentID', 'StudentName', 'Status'],
  EXAMS: ['ExamID', 'GroupID', 'ExamName', 'ExamDate', 'MaxScore'],
  EXAM_RESULTS: ['ExamID', 'StudentID', 'StudentName', 'Score', 'Percentage'],
  SETTINGS: ['Key', 'Value']
};

/**
 * Web App Entry Point
 */
function doGet(e) {
  try {
    initDatabase(); // Ensure all required sheets exist
  } catch (err) {
    Logger.log("Database init error: " + err.message);
  }
  
  var page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page).toLowerCase() : 'admin';
  
  if (page === 'register') {
    return HtmlService.createTemplateFromFile('register')
      .evaluate()
      .setTitle('تسجيل طالب جديد - مستر محمد عبد التواب')
      .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/3429/3429149.png')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Default to Admin Dashboard
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('مستر محمد عبد التواب - نظام إدارة الدروس الخصوصية')
    .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/3429/3429149.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper to include partial files if needed
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// DATABASE INITIALIZATION & UTILITIES
// ==========================================

function initDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  
  for (var key in SHEET_NAMES) {
    var name = SHEET_NAMES[key];
    var headers = HEADERS[key];
    getOrCreateSheet(ss, name, headers);
  }
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f5f9');
      sheet.setFrozenRows(1);
    }
  } else {
    // If sheet exists but is completely empty, add headers
    if (sheet.getLastRow() === 0 && headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f5f9');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function getSheetData(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0];
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    var isEmpty = true;
    for (var j = 0; j < headers.length; j++) {
      var val = row[j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      obj[headers[j]] = val;
      if (val !== "" && val !== null && val !== undefined) isEmpty = false;
    }
    if (!isEmpty) {
      result.push(obj);
    }
  }
  return result;
}

function generateId(prefix) {
  var timestamp = new Date().getTime().toString(36).toUpperCase();
  var random = Math.floor(Math.random() * 1000).toString();
  return prefix + '-' + timestamp + '-' + random;
}

function formatDateISO(dateObj) {
  if (!dateObj) return '';
  if (typeof dateObj === 'string') return dateObj;
  return Utilities.formatDate(new Date(dateObj), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

// ==========================================
// MAIN DATA FETCH API
// ==========================================

function getInitialData() {
  initDatabase();
  
  var groups = getSheetData(SHEET_NAMES.GROUPS);
  var students = getSheetData(SHEET_NAMES.STUDENTS);
  var attendance = getSheetData(SHEET_NAMES.ATTENDANCE);
  var exams = getSheetData(SHEET_NAMES.EXAMS);
  var examResults = getSheetData(SHEET_NAMES.EXAM_RESULTS);
  var settingsRaw = getSheetData(SHEET_NAMES.SETTINGS);
  
  var settings = {};
  settingsRaw.forEach(function(row) {
    if (row.Key) settings[row.Key] = row.Value;
  });
  
  // Default settings
  if (!settings.TeacherName) settings.TeacherName = 'مستر محمد عبد التواب';
  if (!settings.SubjectName) settings.SubjectName = 'مدرس خبير';
  if (!settings.PaymentThreshold) settings.PaymentThreshold = 8;
  if (!settings.PassScore) settings.PassScore = 50;

  var stats = calculateDashboardStats(groups, students, attendance, exams, examResults);
  
  return {
    groups: groups,
    students: students,
    attendance: attendance,
    exams: exams,
    examResults: examResults,
    settings: settings,
    stats: stats
  };
}

// ==========================================
// DASHBOARD CALCULATIONS (REAL DATA ONLY)
// ==========================================

function calculateDashboardStats(groups, students, attendance, exams, examResults) {
  var totalStudents = students.length;
  var totalGroups = groups.length;
  
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var todayAttendance = attendance.filter(function(a) {
    return formatDateISO(a.Date) === todayStr;
  });
  
  var totalAttendanceToday = todayAttendance.filter(function(a) { return a.Status === 'Present'; }).length;
  var totalAbsentToday = todayAttendance.filter(function(a) { return a.Status === 'Absent'; }).length;
  
  var studentsNeedPayment = students.filter(function(s) {
    var count = Number(s.AttendanceCounter) || 0;
    return count >= 8 || s.PaymentStatus === 'Payment Due';
  });
  
  // Attendance Rate
  var totalPresentCount = attendance.filter(function(a) { return a.Status === 'Present'; }).length;
  var overallAttendanceRate = attendance.length > 0 ? ((totalPresentCount / attendance.length) * 100).toFixed(1) : 0;
  
  // Average Score overall
  var scorePercentages = [];
  
  // Create exam map for max scores
  var examMap = {};
  exams.forEach(function(e) {
    examMap[e.ExamID] = Number(e.MaxScore) || 100;
  });
  
  examResults.forEach(function(r) {
    var pct = Number(r.Percentage);
    if (isNaN(pct) || pct === 0) {
      var max = examMap[r.ExamID] || 100;
      pct = (Number(r.Score) / max) * 100;
    }
    if (!isNaN(pct)) {
      scorePercentages.push(pct);
    }
  });
  
  var averageScore = scorePercentages.length > 0 
    ? (scorePercentages.reduce(function(a, b) { return a + b; }, 0) / scorePercentages.length).toFixed(1)
    : 0;
    
  // Group performance map
  var groupScoreMap = {};
  var groupCountMap = {};
  
  examResults.forEach(function(r) {
    var st = students.find(function(s) { return String(s.StudentID) === String(r.StudentID); });
    if (st && st.GroupID) {
      var pct = Number(r.Percentage) || 0;
      if (!groupScoreMap[st.GroupID]) {
        groupScoreMap[st.GroupID] = 0;
        groupCountMap[st.GroupID] = 0;
      }
      groupScoreMap[st.GroupID] += pct;
      groupCountMap[st.GroupID] += 1;
    }
  });
  
  var highestPerformingGroup = 'N/A';
  var lowestPerformingGroup = 'N/A';
  var maxGroupAvg = -1;
  var minGroupAvg = 999;
  
  groups.forEach(function(g) {
    if (groupCountMap[g.GroupID] && groupCountMap[g.GroupID] > 0) {
      var avg = groupScoreMap[g.GroupID] / groupCountMap[g.GroupID];
      if (avg > maxGroupAvg) {
        maxGroupAvg = avg;
        highestPerformingGroup = g.GroupName;
      }
      if (avg < minGroupAvg) {
        minGroupAvg = avg;
        lowestPerformingGroup = g.GroupName;
      }
    }
  });
  
  if (maxGroupAvg === -1) highestPerformingGroup = 'N/A';
  if (minGroupAvg === 999) lowestPerformingGroup = 'N/A';
  
  // Student Average Grades calculation
  var studentStatsMap = {};
  students.forEach(function(s) {
    studentStatsMap[s.StudentID] = {
      student: s,
      scores: [],
      presentCount: 0,
      absentCount: 0,
      excusedCount: 0
    };
  });
  
  examResults.forEach(function(r) {
    if (studentStatsMap[r.StudentID]) {
      var pct = Number(r.Percentage) || 0;
      studentStatsMap[r.StudentID].scores.push(pct);
    }
  });
  
  attendance.forEach(function(a) {
    if (studentStatsMap[a.StudentID]) {
      if (a.Status === 'Present') studentStatsMap[a.StudentID].presentCount++;
      else if (a.Status === 'Absent') studentStatsMap[a.StudentID].absentCount++;
      else if (a.Status === 'Excused') studentStatsMap[a.StudentID].excusedCount++;
    }
  });
  
  var studentListWithAvg = Object.keys(studentStatsMap).map(function(sid) {
    var item = studentStatsMap[sid];
    var avgGrade = item.scores.length > 0 
      ? (item.scores.reduce(function(a, b) { return a + b; }, 0) / item.scores.length) 
      : 0;
    return {
      StudentID: sid,
      StudentName: item.student.StudentName,
      GroupID: item.student.GroupID,
      AverageGrade: Number(avgGrade.toFixed(1)),
      PresentCount: item.presentCount,
      AbsentCount: item.absentCount,
      ExcusedCount: item.excusedCount,
      AttendanceCounter: item.student.AttendanceCounter,
      PaymentStatus: item.student.PaymentStatus,
      ParentPhone: item.student.ParentPhone
    };
  });
  
  // Top 5 Students
  var topStudents = studentListWithAvg.slice()
    .filter(function(s) { return s.AverageGrade > 0; })
    .sort(function(a, b) { return b.AverageGrade - a.AverageGrade; })
    .slice(0, 5);
    
  // Lowest 5 Performance Students
  var lowestStudents = studentListWithAvg.slice()
    .filter(function(s) { return s.AverageGrade > 0 || s.AbsentCount > 0; })
    .sort(function(a, b) { 
      if (a.AverageGrade !== b.AverageGrade) return a.AverageGrade - b.AverageGrade;
      return b.AbsentCount - a.AbsentCount;
    })
    .slice(0, 5);
    
  // Success vs Fail Rate
  var passCount = studentListWithAvg.filter(function(s) { return s.AverageGrade >= 50; }).length;
  var failCount = studentListWithAvg.filter(function(s) { return s.AverageGrade < 50 && s.AverageGrade > 0; }).length;
  
  return {
    totalStudents: totalStudents,
    totalGroups: totalGroups,
    totalAttendanceToday: totalAttendanceToday,
    totalAbsentToday: totalAbsentToday,
    studentsNeedPaymentCount: studentsNeedPayment.length,
    overallAttendanceRate: overallAttendanceRate,
    averageScore: averageScore,
    highestPerformingGroup: highestPerformingGroup,
    lowestPerformingGroup: lowestPerformingGroup,
    studentsNeedPaymentList: studentsNeedPayment.slice(0, 10),
    topStudents: topStudents,
    lowestStudents: lowestStudents,
    passCount: passCount,
    failCount: failCount
  };
}

// ==========================================
// GROUPS MANAGEMENT
// ==========================================

function saveGroup(groupData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
  var data = sheet.getDataRange().getValues();
  
  var groupID = groupData.GroupID;
  var isEdit = false;
  var rowIndex = -1;
  
  if (groupID) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(groupID)) {
        isEdit = true;
        rowIndex = i + 1;
        break;
      }
    }
  } else {
    groupID = generateId('GRP');
  }
  
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  var rowValues = [
    groupID,
    groupData.GroupName,
    groupData.Grade,
    groupData.StudyType,
    groupData.Days,
    groupData.Time,
    isEdit ? data[rowIndex - 1][6] : now
  ];
  
  if (isEdit) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  
  return getInitialData();
}

function deleteGroup(groupID) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(groupID)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getInitialData();
}

// ==========================================
// STUDENTS MANAGEMENT
// ==========================================

function saveStudent(studentData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
  var data = sheet.getDataRange().getValues();
  
  var studentID = studentData.StudentID;
  var isEdit = false;
  var rowIndex = -1;
  
  if (studentID) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(studentID)) {
        isEdit = true;
        rowIndex = i + 1;
        break;
      }
    }
  } else {
    studentID = generateId('STD');
  }
  
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  
  var rowValues = [
    studentID,
    studentData.GroupID,
    studentData.StudentName,
    studentData.ParentPhone,
    isEdit ? (data[rowIndex - 1][4] !== "" ? data[rowIndex - 1][4] : 0) : 0,
    isEdit ? (data[rowIndex - 1][5] || 'Paid') : 'Paid',
    isEdit ? data[rowIndex - 1][6] : now
  ];
  
  if (isEdit) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  
  return getInitialData();
}

function deleteStudent(studentID) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(studentID)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getInitialData();
}

function confirmPayment(studentID) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(studentID)) {
      // Reset AttendanceCounter to 0 and PaymentStatus to Paid
      sheet.getRange(i + 1, 5).setValue(0);
      sheet.getRange(i + 1, 6).setValue('Paid');
      break;
    }
  }
  return getInitialData();
}

// ==========================================
// ATTENDANCE MANAGEMENT
// ==========================================

function saveAttendance(attendancePayload) {
  // attendancePayload: { groupID: '...', date: 'yyyy-MM-dd', records: [{ studentID, studentName, status }] }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
  var stdSheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
  
  var groupID = attendancePayload.groupID;
  var dateStr = attendancePayload.date;
  var records = attendancePayload.records;
  
  // Read existing students to update counters
  var stdData = stdSheet.getDataRange().getValues();
  var stdMap = {};
  for (var i = 1; i < stdData.length; i++) {
    stdMap[String(stdData[i][0])] = {
      rowIndex: i + 1,
      counter: Number(stdData[i][4]) || 0,
      status: stdData[i][5]
    };
  }
  
  var newAttRows = [];
  
  records.forEach(function(rec) {
    var attID = generateId('ATT');
    newAttRows.push([
      attID,
      dateStr,
      groupID,
      rec.studentID,
      rec.studentName,
      rec.status
    ]);
    
    // Update student
    if (stdMap[rec.studentID]) {
      var sObj = stdMap[rec.studentID];
      var row = sObj.rowIndex;
      
      if (rec.status === 'Present') {
        var newCounter = sObj.counter + 1;
        var newPaymentStatus = newCounter >= 8 ? 'Payment Due' : 'Paid';
        stdSheet.getRange(row, 5).setValue(newCounter);
        stdSheet.getRange(row, 6).setValue(newPaymentStatus);
      }
    }
  });
  
  if (newAttRows.length > 0) {
    attSheet.getRange(attSheet.getLastRow() + 1, 1, newAttRows.length, newAttRows[0].length).setValues(newAttRows);
  }
  
  return getInitialData();
}

// ==========================================
// EXAMS MANAGEMENT
// ==========================================

function saveExam(examData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.EXAMS);
  var data = sheet.getDataRange().getValues();
  
  var examID = examData.ExamID;
  var isEdit = false;
  var rowIndex = -1;
  
  if (examID) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(examID)) {
        isEdit = true;
        rowIndex = i + 1;
        break;
      }
    }
  } else {
    examID = generateId('EXM');
  }
  
  var rowValues = [
    examID,
    examData.GroupID,
    examData.ExamName,
    examData.ExamDate,
    Number(examData.MaxScore) || 100
  ];
  
  if (isEdit) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  
  return getInitialData();
}

function deleteExam(examID) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.EXAMS);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(examID)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  
  // Also delete from ExamResults
  var resSheet = ss.getSheetByName(SHEET_NAMES.EXAM_RESULTS);
  var resData = resSheet.getDataRange().getValues();
  for (var k = resData.length - 1; k >= 1; k--) {
    if (String(resData[k][0]) === String(examID)) {
      resSheet.deleteRow(k + 1);
    }
  }
  
  return getInitialData();
}

function saveExamResults(payload) {
  // payload: { examID: '...', maxScore: 100, results: [{ studentID, studentName, score }] }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.EXAM_RESULTS);
  var data = sheet.getDataRange().getValues();
  
  var examID = payload.examID;
  var maxScore = Number(payload.maxScore) || 100;
  var results = payload.results;
  
  // Map existing results row index
  var existingMap = {};
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(examID)) {
      var sID = String(data[i][1]);
      existingMap[sID] = i + 1;
    }
  }
  
  results.forEach(function(r) {
    var score = Number(r.score) || 0;
    var pct = ((score / maxScore) * 100).toFixed(1);
    var sID = String(r.studentID);
    
    if (existingMap[sID]) {
      var rIdx = existingMap[sID];
      sheet.getRange(rIdx, 4, 1, 2).setValues([[score, pct]]);
    } else {
      sheet.appendRow([examID, r.studentID, r.studentName, score, pct]);
    }
  });
  
  return getInitialData();
}

// ==========================================
// SETTINGS MANAGEMENT
// ==========================================

function saveSettings(settingsObj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  sheet.clearContents();
  sheet.appendRow(HEADERS.SETTINGS);
  
  for (var key in settingsObj) {
    sheet.appendRow([key, settingsObj[key]]);
  }
  
  return getInitialData();
}

// ==========================================
// DEMO DATA GENERATOR
// ==========================================

function createDemoData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear all existing sheets
  for (var key in SHEET_NAMES) {
    var name = SHEET_NAMES[key];
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      sheet.clearContents();
      sheet.getRange(1, 1, 1, HEADERS[key].length).setValues([HEADERS[key]]);
      sheet.getRange(1, 1, 1, HEADERS[key].length).setFontWeight('bold').setBackground('#f1f5f9');
      sheet.setFrozenRows(1);
    }
  }
  
  var groupsSheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
  var studentsSheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
  var attendanceSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
  var examsSheet = ss.getSheetByName(SHEET_NAMES.EXAMS);
  var examResultsSheet = ss.getSheetByName(SHEET_NAMES.EXAM_RESULTS);
  var settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  
  // 1. Create 5 Groups
  var demoGroups = [
    ['GRP-101', 'Grade 1 Arabic', 'Grade 1', 'Arabic', 'Sat & Tue', '04:00 PM', '2026-07-01 10:00'],
    ['GRP-102', 'Grade 1 Languages', 'Grade 1', 'Languages', 'Sun & Wed', '05:30 PM', '2026-07-01 10:30'],
    ['GRP-103', 'Grade 2 Arabic', 'Grade 2', 'Arabic', 'Mon & Thu', '04:00 PM', '2026-07-01 11:00'],
    ['GRP-104', 'Grade 3 Languages', 'Grade 3', 'Languages', 'Sat & Tue', '06:00 PM', '2026-07-01 11:30'],
    ['GRP-105', 'Secondary 1 Arabic', 'Secondary 1', 'Arabic', 'Sun & Wed', '07:00 PM', '2026-07-01 12:00']
  ];
  groupsSheet.getRange(2, 1, demoGroups.length, HEADERS.GROUPS.length).setValues(demoGroups);
  
  // 2. Create 100 Students with Realistic Arabic Names
  var firstNames = [
    'أحمد', 'محمد', 'يوسف', 'عمر', 'محمود', 'علي', 'حسن', 'حسين', 'عمرو', 'طارق',
    'كريم', 'مصطفى', 'زياد', 'حمزة', 'ياسين', 'آدم', 'نور', 'فاطمة', 'مريم', 'سلمى',
    'نوران', 'حبيبة', 'جنى', 'ملك', 'شهد', 'منة', 'ليلى', 'مايا', 'رانيا', 'ريم',
    'خالد', 'عبد الرحمن', 'إبراهيم', 'بلال', 'سارة', 'فاروق', 'هاني', 'شريف', 'منار', 'داليا'
  ];
  
  var lastNames = [
    'المصري', 'السيد', 'إبراهيم', 'عبد الرحمن', 'حسن', 'محمود', 'النجار', 'الشامي',
    'عبد العزيز', 'غانم', 'فوزي', 'رضوان', 'حجازي', 'بدوي', 'زكي', 'فرج', 'عثمان', 'شحاتة'
  ];
  
  var phonePrefixes = ['010', '011', '012', '015'];
  
  var demoStudents = [];
  var studentObjects = [];
  
  for (var i = 1; i <= 100; i++) {
    var sid = 'STD-' + (1000 + i);
    var grp = demoGroups[(i - 1) % 5];
    var fn = firstNames[Math.floor(Math.random() * firstNames.length)];
    var ln = lastNames[Math.floor(Math.random() * lastNames.length)];
    var fullName = fn + ' ' + ln;
    var phone = phonePrefixes[Math.floor(Math.random() * phonePrefixes.length)] + Math.floor(1000007 + Math.random() * 8999990);
    
    // Vary attendance counter so some have >= 8 (payment due)
    var counter = Math.floor(Math.random() * 10);
    var payStatus = counter >= 8 ? 'Payment Due' : 'Paid';
    var createdAt = '2026-07-02 09:00';
    
    demoStudents.push([sid, grp[0], fullName, phone, counter, payStatus, createdAt]);
    studentObjects.push({
      id: sid,
      name: fullName,
      groupId: grp[0],
      groupName: grp[1]
    });
  }
  
  studentsSheet.getRange(2, 1, demoStudents.length, HEADERS.STUDENTS.length).setValues(demoStudents);
  
  // 3. Create Attendance History across past 30 days
  var demoAttendance = [];
  var dates = [
    '2026-07-05', '2026-07-08', '2026-07-12', '2026-07-15', 
    '2026-07-19', '2026-07-22', '2026-07-26', '2026-07-29'
  ];
  
  var attCounter = 1;
  dates.forEach(function(dt) {
    studentObjects.forEach(function(st) {
      // 85% Present, 10% Absent, 5% Excused
      var rand = Math.random();
      var status = rand < 0.85 ? 'Present' : (rand < 0.95 ? 'Absent' : 'Excused');
      demoAttendance.push([
        'ATT-' + (2000 + attCounter++),
        dt,
        st.groupId,
        st.id,
        st.name,
        status
      ]);
    });
  });
  
  attendanceSheet.getRange(2, 1, demoAttendance.length, HEADERS.ATTENDANCE.length).setValues(demoAttendance);
  
  // 4. Create Exams
  var demoExams = [
    ['EXM-301', 'GRP-101', 'Quiz 1 - Grammar', '2026-07-10', 20],
    ['EXM-302', 'GRP-101', 'Monthly Exam 1', '2026-07-25', 100],
    ['EXM-303', 'GRP-102', 'Vocabulary Quiz', '2026-07-12', 30],
    ['EXM-304', 'GRP-103', 'Arabic Dictation', '2026-07-15', 20],
    ['EXM-305', 'GRP-104', 'Midterm Revision Exam', '2026-07-20', 50],
    ['EXM-306', 'GRP-105', 'Comprehensive Mock Exam', '2026-07-22', 100]
  ];
  examsSheet.getRange(2, 1, demoExams.length, HEADERS.EXAMS.length).setValues(demoExams);
  
  // 5. Create Exam Results for all students in respective groups
  var demoResults = [];
  demoExams.forEach(function(ex) {
    var examID = ex[0];
    var grpID = ex[1];
    var maxScore = ex[4];
    
    var groupStds = studentObjects.filter(function(s) { return s.groupId === grpID; });
    groupStds.forEach(function(s) {
      // Score ranges between 45% and 100% of max score
      var scorePct = 45 + Math.random() * 54;
      var rawScore = Math.round((scorePct / 100) * maxScore);
      var calcPct = ((rawScore / maxScore) * 100).toFixed(1);
      
      demoResults.push([examID, s.id, s.name, rawScore, calcPct]);
    });
  });
  
  examResultsSheet.getRange(2, 1, demoResults.length, HEADERS.EXAM_RESULTS.length).setValues(demoResults);
  
  // 6. Settings
  var demoSettings = [
    ['TeacherName', 'مستر محمد عبد التواب'],
    ['SubjectName', 'مدرس خبير'],
    ['PaymentThreshold', '8'],
    ['PassScore', '50'],
    ['Currency', 'ج.م']
  ];
  settingsSheet.getRange(2, 1, demoSettings.length, HEADERS.SETTINGS.length).setValues(demoSettings);
  
  return getInitialData();
}

// ==========================================
// PUBLIC STUDENT REGISTRATION API
// ==========================================

function getPublicGroups() {
  initDatabase();
  var groups = getSheetData(SHEET_NAMES.GROUPS);
  return groups.map(function(g) {
    return {
      GroupID: g.GroupID,
      GroupName: g.GroupName,
      Grade: g.Grade,
      StudyType: g.StudyType,
      Days: g.Days,
      Time: g.Time
    };
  });
}

function registerStudentPublic(studentData) {
  initDatabase();
  var name = String(studentData.studentName || '').trim();
  var phone = String(studentData.parentPhone || '').trim();
  var groupID = String(studentData.groupID || '').trim();

  if (!name || !phone || !groupID) {
    throw new Error('جميع الحقول مطلوبة (اسم الطالب، رقم ولي الأمر، والمجموعة)');
  }

  var phoneClean = phone.replace(/[\s\-\+\(\)]/g, '');
  if (!/^\d{10,15}$/.test(phoneClean)) {
    throw new Error('رقم ولي الأمر غير صحيح. يرجى إدخال رقم هاتف محمول صحيح.');
  }

  var payload = {
    StudentID: '',
    GroupID: groupID,
    StudentName: name,
    ParentPhone: phone
  };

  // Re-use existing saveStudent function
  saveStudent(payload);

  return {
    success: true,
    message: 'تم تسجيل بياناتك بنجاح.'
  };
}
