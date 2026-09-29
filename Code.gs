// ═══════════════════════════════════════════════════════════════════════════════
//  CYS Attendance System — Apps Script
//
//  Student data lives in a "STUDENTS" sheet (editable directly in Sheets UI):
//    Col A: SEM  (e.g. II or IV)
//    Col B: SNO  (serial no within section)
//    Col C: REGNO (e.g. 25130001)
//    Col D: NAME
//
//  Edit that sheet anytime to add / remove / rename students.
//  Every month sheet is built fresh from it.
// ═══════════════════════════════════════════════════════════════════════════════

var SPREADSHEET_ID = '1SMZ48JP46DdSzpa0Krk9VRV6T1WSCGwMME41SuA6vMU';
var COLLEGE        = 'SRI RAMAKRISHNA COLLEGE OF ARTS & SCIENCE(AUTONOMOUS)';
var DEPT           = 'NAME OF THE DEPARTMENT: COMPUTER SCIENCE WITH  CYBER SECURITY';
var WEEK_NAMES     = ['I','II','III','IV','V'];

var MONTH_NUMS = {
  JAN:1,FEB:2,MAR:3,APR:4,MAY:5,JUN:6,
  JUL:7,AUG:8,SEP:9,OCT:10,NOV:11,DEC:12
};
var MONTH_FULL = {
  JAN:'JANUARY',FEB:'FEBRUARY',MAR:'MARCH',    APR:'APRIL',
  MAY:'MAY',    JUN:'JUNE',    JUL:'JULY',     AUG:'AUGUST',
  SEP:'SEPTEMBER',OCT:'OCTOBER',NOV:'NOVEMBER',DEC:'DECEMBER'
};

// Default student data — only used the very first time (to seed the STUDENTS sheet if not present)
var DEFAULT_STUDENTS = [
  ['II', 1,'25130001','AAKASH S'],
  ['II', 2,'25130002','AARIS U'],
  ['II', 3,'25130003','ABDUL ANEESH A'],
  ['II', 4,'25130004','ANUSHA R'],
  ['II', 5,'25130005','ARRISHMAA M'],
  ['II', 6,'25130006','ARUNA DEVI M'],
  ['II', 7,'25130007','ASHWIN KARTHICK K'],
  ['II', 8,'25130008','ATSHYA A'],
  ['II', 9,'25130009','BAVILAN B'],
  ['II',10,'25130010','CHEZHIYAN S'],
  ['II',11,'25130011','DANIEL S'],
  ['II',12,'25130012','DEVAASHREE R'],
  ['II',13,'25130013','DHAJIN T'],
  ['II',14,'25130014','DHAKSHAN A'],
  ['II',15,'25130015','DHARANI S'],
  ['II',16,'25130016','DHARESH T'],
  ['II',17,'25130017','DHARSHINI M'],
  ['II',18,'25130018','DHEEKSHA J'],
  ['II',19,'25130019','DINESH KARTHICK R'],
  ['II',20,'25130020','GOKULAVANI V'],
  ['IV',21,'25130021','GOPIKA S'],
  ['IV',22,'25130022','HARICHARAN S V'],
  ['IV',23,'25130023','HARIPRIYAN V'],
  ['IV',24,'25130024','HARSHINI C'],
  ['IV',25,'25130025','JAYA VIGNESH T'],
  ['IV',26,'25130026','JINCY T'],
  ['IV',27,'25130027','KAMALESH R'],
  ['IV',28,'25130028','KANMANI S'],
  ['IV',29,'25130029','KIRTHICKRAJ K'],
  ['IV',30,'25130030','MADHUMITHA J'],
  ['IV',31,'25130031','MANIMEGALAI P'],
  ['IV',32,'25130032','MATHAN KUMAR C'],
  ['IV',33,'25130033','MAYILVENI N'],
  ['IV',34,'25130034','MITHRA J'],
  ['IV',35,'25130035','MOHAMMED SUHAIL A'],
  ['IV',36,'25130036','NASRIN FATHIMA A'],
  ['IV',37,'25130037','NIRMAL AS'],
  ['IV',38,'25130038','NISHANTH R S'],
  ['IV',39,'25130039','NIVETHA V'],
  ['IV',40,'25130040','PRATHIKSHA S'],
  ['IV',41,'25130041','PRAVEEN M'],
  ['IV',42,'25130042','PRIYADHARSHINI V'],
  ['IV',43,'25130043','RENUKA S'],
  ['IV',44,'25130044','RITHIKA M'],
  ['IV',45,'25130045','MUSTAFA ELSADIG ELFAKI'],
  ['IV',46,'25130046','RUPAM DALAI S'],
  ['IV',47,'25130047','SABARIVASAN A'],
  ['IV',48,'25130048','SACHIN S'],
  ['IV',49,'25130049','SHAMNI R'],
  ['IV',50,'25130050','SHIVA GANESH P'],
  ['IV',51,'25130051','SIDDHARTH K'],
  ['IV',52,'25130052','SIVANANDAN G'],
  ['IV',53,'25130053','SOBHIKA S'],
  ['IV',54,'25130054','SREYA S'],
  ['IV',55,'25130055','SRIDHAR P'],
  ['IV',56,'25130056','SRIRAM R'],
  ['IV',57,'25130057','SRIVITHYA S'],
  ['IV',58,'25130058','SUNDARJI S'],
  ['IV',59,'25130059','SURENDAR M'],
  ['IV',60,'25130060','SWATHI S'],
  ['IV',61,'25130061','THEO ANTONY A'],
  ['IV',62,'25130062','VIBIN I'],
  ['IV',63,'25130063','VIJAY JOSHUA S'],
  ['IV',64,'25130064','YOGANANTHAM R'],
  ['IV',65,'25130065','YOGESHWARAN M'],
  ['IV',66,'25130066','YAKSHITHA P V']
];

// ── Helpers ───────────────────────────────────────────────────────────────────

function getSpreadsheet(sheetId) {
  var id = sheetId || SPREADSHEET_ID;
  // Always try opening by ID first to guarantee writing to the correct sheet
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch(err) {}
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch(e) {}
  throw new Error('Spreadsheet not accessible (' + id + '). Authorize or bind the script to your sheet.');
}

function weeksInMonth(year, month) {
  var days = new Date(year, month, 0).getDate(), count = 0;
  for (var d = 1; d <= days; d++)
    if (new Date(year, month-1, d).getDay() === 1) count++;
  return count || 4;
}

function academicYear(abbr) {
  var m = MONTH_NUMS[abbr] || 11;
  var yr = new Date().getFullYear();
  return (m >= 6) ? yr : yr + 1;
}

function colFor(weekNum, dayOrder) {
  return 4 + (weekNum - 1) * 6 + (dayOrder - 1);
}

function jsonOut(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── STUDENTS master sheet ─────────────────────────────────────────────────────

// Returns the STUDENTS sheet, creating and seeding it if missing
function getStudentsSheet(ss) {
  var sh = ss.getSheetByName('STUDENTS');
  if (sh) return sh;

  sh = ss.insertSheet('STUDENTS');
  // Header row
  sh.getRange(1, 1, 1, 4).setValues([['SEM','SNO','REGNO','NAME']]);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold');
  // Seed with default data
  sh.getRange(2, 1, DEFAULT_STUDENTS.length, 4).setValues(DEFAULT_STUDENTS);
  sh.autoResizeColumns(1, 4);
  return sh;
}

// Read all students from STUDENTS sheet, grouped by SEM
function readSections(ss) {
  var sh   = getStudentsSheet(ss);
  var rows = sh.getDataRange().getValues();
  var map  = {};  // sem → [{sno,regno,name}]

  for (var i = 1; i < rows.length; i++) {  // skip header
    var sem   = String(rows[i][0]).trim();
    var sno   = rows[i][1];
    var regno = String(rows[i][2]).trim();
    var name  = String(rows[i][3]).trim();
    if (!sem || !regno || !name) continue;
    if (!map[sem]) map[sem] = [];
    map[sem].push([sno, regno, name]);
  }

  // Return as ordered array of {sem, students}
  var sems = Object.keys(map).sort();
  return sems.map(function(s) { return { sem: s, students: map[s] }; });
}

// ── Build / refresh month sheet ───────────────────────────────────────────────

function applyColumnWidths(sheet, attCols) {
  sheet.setColumnWidth(1, 32);    // SNO
  sheet.setColumnWidth(2, 75);    // REGNO
  sheet.setColumnWidth(3, 160);   // NAME
  for (var c = 4; c <= 3 + attCols; c++) {
    sheet.setColumnWidth(c, 20);  // Days 1-6
  }
  sheet.setColumnWidth(4 + attCols, 38); // CUR MON
  sheet.setColumnWidth(5 + attCols, 38); // PRE MON
  sheet.setColumnWidth(6 + attCols, 38); // TOTAL
}

function buildSheet(ss, abbr, year) {
  var nWeeks   = weeksInMonth(year, MONTH_NUMS[abbr] || 11);
  var attCols  = 6 * nWeeks;
  var totalCols = 3 + attCols + 3;

  // If sheet already exists → re-apply compact column widths and return
  var existing = ss.getSheetByName(abbr);
  if (existing) {
    applyColumnWidths(existing, attCols);
    return { sheet: existing, nWeeks: nWeeks };
  }

  // Build from scratch using current STUDENTS list
  var sections = readSections(ss);
  var mFull    = MONTH_FULL[abbr] + ' ' + year;

  var sheet = ss.insertSheet(abbr);
  // Ensure enough columns exist before formatting
  if (sheet.getMaxColumns() < totalCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), totalCols - sheet.getMaxColumns());
  }

  var startRow = 1;

  // Split any section with >20 students into chunks of 20
  var CHUNK = 20;
  var blocks = [];
  sections.forEach(function(sec) {
    for (var i = 0; i < sec.students.length; i += CHUNK) {
      blocks.push({ sem: sec.sem, students: sec.students.slice(i, i + CHUNK) });
    }
  });

  blocks.forEach(function(sec) {
    var students = sec.students;
    var nStu     = students.length;

    // Row 1: College name (set value on cell (startRow, 1) then merge, avoiding duplication across columns)
    sheet.getRange(startRow, 1).setValue(COLLEGE);
    sheet.getRange(startRow, 1, 1, totalCols).merge()
      .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');

    // Row 2: Statement
    sheet.getRange(startRow+1, 1).setValue('STATEMENT OF THE ATTENDANCE  FOR THE MONTH : ' + mFull);
    sheet.getRange(startRow+1, 1, 1, totalCols).merge()
      .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');

    // Row 3: Dept | Semester
    sheet.getRange(startRow+2, 1).setValue(DEPT);
    sheet.getRange(startRow+2, 1, 1, totalCols-2).merge()
      .setFontWeight('bold').setFontSize(9);

    sheet.getRange(startRow+2, totalCols-1).setValue('SEMESTER : ' + sec.sem);
    sheet.getRange(startRow+2, totalCols-1, 1, 2).merge()
      .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('right');

    // Row 4: DAY ORDER label
    sheet.getRange(startRow+3, 4).setValue('DAY ORDER');
    sheet.getRange(startRow+3, 4, 1, attCols).merge()
      .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');

    // Row 5 & 6: Column headers (S.NO, REG NO, NAME vertically merged over 2 rows)
    sheet.getRange(startRow+4, 1).setValue('S.NO');
    sheet.getRange(startRow+4, 1, 2, 1).merge().setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center').setVerticalAlignment('middle');

    sheet.getRange(startRow+4, 2).setValue('REG NO');
    sheet.getRange(startRow+4, 2, 2, 1).merge().setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center').setVerticalAlignment('middle');

    sheet.getRange(startRow+4, 3).setValue('NAME');
    sheet.getRange(startRow+4, 3, 2, 1).merge().setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    
    for (var w = 0; w < nWeeks; w++) {
      var wCol = 4 + w*6;
      sheet.getRange(startRow+4, wCol).setValue(WEEK_NAMES[w] + ' WEEK');
      sheet.getRange(startRow+4, wCol, 1, 6).merge()
        .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    }
    var curCol = 4 + attCols;
    sheet.getRange(startRow+4, curCol).setValue('CUR\nMON');
    sheet.getRange(startRow+4, curCol, 2, 1).merge().setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center').setVerticalAlignment('middle');

    sheet.getRange(startRow+4, curCol+1).setValue('PRE\nMON');
    sheet.getRange(startRow+4, curCol+1, 2, 1).merge().setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center').setVerticalAlignment('middle');

    sheet.getRange(startRow+4, curCol+2).setValue('TOTAL');
    sheet.getRange(startRow+4, curCol+2, 2, 1).merge().setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center').setVerticalAlignment('middle');

    // Row 6: Day order numbers 1–6 per week
    for (var w2 = 0; w2 < nWeeks; w2++) {
      for (var d = 1; d <= 6; d++) {
        sheet.getRange(startRow+5, 4 + w2*6 + (d-1)).setValue(d).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
      }
    }

    // Rows 7+: Student rows
    students.forEach(function(stu, idx) {
      var r = startRow + 6 + idx;
      sheet.getRange(r, 1).setValue(stu[0]).setFontSize(9);  // SNO
      sheet.getRange(r, 2).setValue(stu[1]).setFontSize(9);  // REGNO
      sheet.getRange(r, 3).setValue(stu[2]).setFontSize(9);  // NAME
    });

    // Apply alignments and font size for student attendance rows
    sheet.getRange(startRow + 6, 1, nStu, 1).setHorizontalAlignment('center');
    sheet.getRange(startRow + 6, 2, nStu, 1).setHorizontalAlignment('center');
    sheet.getRange(startRow + 6, 3, nStu, 1).setHorizontalAlignment('left');
    sheet.getRange(startRow + 6, 4, nStu, attCols + 3).setHorizontalAlignment('center').setFontSize(9);

    // Set compact row heights for data rows (20px each)
    sheet.setRowHeights(startRow + 6, nStu, 20);

    // Apply full solid black borders to the entire table block (Headers + Student rows)
    var tableRange = sheet.getRange(startRow + 4, 1, 2 + nStu, totalCols);
    tableRange.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

    // Notes (Attendance Legend)
    var noteRow = startRow + 6 + nStu;
    sheet.getRange(noteRow, 1).setValue('Full Day Absent - a');
    sheet.getRange(noteRow, 1, 1, totalCols).merge().setFontSize(9);

    sheet.getRange(noteRow+1, 1).setValue('Half Day absent -   FN = a/    AN =/a');
    sheet.getRange(noteRow+1, 1, 1, Math.ceil(totalCols/2)).merge().setFontSize(9);

    sheet.getRange(noteRow+1, Math.ceil(totalCols/2)+1).setValue('CLASS INCHARGE');
    sheet.getRange(noteRow+1, Math.ceil(totalCols/2)+1, 1, Math.floor(totalCols/4)).merge()
      .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');

    var hodCol = Math.ceil(totalCols/2) + Math.floor(totalCols/4) + 1;
    var hodSpan = totalCols - hodCol + 1;
    sheet.getRange(noteRow+1, hodCol).setValue('HOD');
    sheet.getRange(noteRow+1, hodCol, 1, hodSpan).merge()
      .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');

    startRow = noteRow + 2;
  });

  // Apply compact column widths
  applyColumnWidths(sheet, attCols);

  return { sheet: sheet, nWeeks: nWeeks };
}

// ── GET ───────────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    var abbr    = ((e && e.parameter && e.parameter.sheet) || 'NOV').toUpperCase();
    var year    = parseInt(e && e.parameter && e.parameter.year) || new Date().getFullYear();
    var sheetId = (e && e.parameter && e.parameter.sheetId) || SPREADSHEET_ID;
    var ss      = getSpreadsheet(sheetId);

    var existed = ss.getSheetByName(abbr) !== null;
    var result  = buildSheet(ss, abbr, year);  // no-op if sheet already exists
    var sheet   = result.sheet;
    var nWeeks  = result.nWeeks;

    var data = sheet.getDataRange().getValues();
    var students = [];
    for (var r = 0; r < data.length; r++) {
      var sno   = String(data[r][0]).trim();
      var regno = String(data[r][1]).trim();
      var name  = String(data[r][2]).trim();
      if (/^\d{8}$/.test(regno) && name.length > 0)
        students.push({ row: r+1, sno: sno, regno: regno, name: name });
    }

    return jsonOut({
      success: true,
      students: students,
      weeks: nWeeks,
      sheet: abbr,
      year: year,
      created: !existed
    });
  } catch(err) {
    return jsonOut({
      success: false,
      error: err.toString()
    });
  }
}

// ── POST ──────────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var payload   = JSON.parse(e.postData.contents);
    var abbr      = ((payload.sheet) || 'NOV').toUpperCase();
    var year      = parseInt(payload.year) || new Date().getFullYear();
    var weekNum   = parseInt(payload.weekNum);
    var dayOrder  = parseInt(payload.dayOrder);
    var sheetId   = payload.sheetId || SPREADSHEET_ID;

    if (weekNum < 1 || weekNum > 5 || dayOrder < 1 || dayOrder > 6)
      return jsonOut({ success: false, error: 'Invalid weekNum/dayOrder' });

    // Normalize absentees into a map of regno -> mark
    // mark options: 'a' (Full Day Absent), 'a/' (FN Half Day Absent), '/a' (AN Half Day Absent)
    var absMap = {};
    if (payload.absentees) {
      if (Array.isArray(payload.absentees)) {
        payload.absentees.forEach(function(item) {
          if (typeof item === 'string') {
            absMap[item] = 'a';
          } else if (item && item.regno) {
            absMap[item.regno] = item.type || 'a';
          }
        });
      } else if (typeof payload.absentees === 'object') {
        Object.keys(payload.absentees).forEach(function(k) {
          absMap[k] = payload.absentees[k];
        });
      }
    }

    var ss    = getSpreadsheet(sheetId);
    buildSheet(ss, abbr, year);  // creates sheet only if it doesn't exist yet
    var sheet = ss.getSheetByName(abbr);
    var col   = colFor(weekNum, dayOrder);
    var data  = sheet.getDataRange().getValues();

    var marked = [];
    var presentCount = 0;
    var fullAbsentCount = 0;
    var fnAbsentCount = 0;
    var anAbsentCount = 0;

    for (var r = 0; r < data.length; r++) {
      var regno = String(data[r][1]).trim();
      var name  = String(data[r][2]).trim();
      if (!/^\d{8}$/.test(regno)) continue;

      var absentMark = absMap[regno];
      var mark = '/';
      if (typeof absentMark !== 'undefined') {
        if (absentMark === 'a/' || absentMark === 'FN' || absentMark === 'fn') {
          mark = 'a/';
          fnAbsentCount++;
        } else if (absentMark === '/a' || absentMark === 'AN' || absentMark === 'an') {
          mark = '/a';
          anAbsentCount++;
        } else {
          mark = 'a';
          fullAbsentCount++;
        }
      } else {
        mark = '/';
        presentCount++;
      }

      sheet.getRange(r+1, col).setValue(mark);
      marked.push({ regno: regno, name: name, mark: mark, row: r+1 });
    }

    return jsonOut({
      success: true,
      sheet: abbr,
      weekNum: weekNum,
      dayOrder: dayOrder,
      presentCount: presentCount,
      fullAbsentCount: fullAbsentCount,
      fnAbsentCount: fnAbsentCount,
      anAbsentCount: anAbsentCount,
      absentCount: fullAbsentCount + fnAbsentCount + anAbsentCount,
      marked: marked
    });

  } catch(err) {
    return jsonOut({ success: false, error: err.toString() });
  }
}

