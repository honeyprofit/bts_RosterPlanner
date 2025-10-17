/*** CONFIG ***/
const SOURCE_SHEET_NAME = 'matrix_employees'; // 직원 데이터 시트 (A:ID, B:이름, C:고용유형코드)
const KR_HOLIDAY_CAL_ID = 'ko.south_korea#holiday@group.v.calendar.google.com'; // 한국 공휴일 캘린더
const MON_ABBR = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RQ')
    .addItem('신규 월간 시트 생성', 'rqCreateNewSheet')
    .addToUi();
}

// RQ 버튼: 연/월 입력 → 새 시트 생성(중복 시 (1),(2)…)
function rqCreateNewSheet() {
  var ui = SpreadsheetApp.getUi();

  var y = ui.prompt('연도 입력', '예: 2025', ui.ButtonSet.OK_CANCEL);
  if (y.getSelectedButton() !== ui.Button.OK) return;
  var year = Number(String(y.getResponseText()).trim());
  if (!isFinite(year) || year < 1900 || year > 2100) {
    ui.alert('연도 입력이 올바르지 않습니다. (예: 2025)');
    return;
  }

  var m = ui.prompt('월 입력', '1 ~ 12 중 하나 (예: 10)', ui.ButtonSet.OK_CANCEL);
  if (m.getSelectedButton() !== ui.Button.OK) return;
  var month = Number(String(m.getResponseText()).trim());
  if (!isFinite(month) || month < 1 || month > 12) {
    ui.alert('월 입력이 올바르지 않습니다. (1~12)');
    return;
  }

  // Per-person limit (1~13)
  var l = ui.prompt('개인 최대 RQ 신청 가능 수 입력', '1 ~ 13 (예: 5)', ui.ButtonSet.OK_CANCEL);
  if (l.getSelectedButton() !== ui.Button.OK) return;
  var perPersonLimit = Number(String(l.getResponseText()).trim());
  if (!isFinite(perPersonLimit) || perPersonLimit < 1 || perPersonLimit > 13) {
    ui.alert('개인 최대 수는 1~13 사이의 정수여야 합니다.'); return;
  }

  var baseName = String(year).slice(-2) + ' ' + MON_ABBR[month - 1] + ' RQ';
  var sheetName = getUniqueSheetName_(baseName);

  var ss = SpreadsheetApp.getActive();
  var sheet = ss.insertSheet(sheetName); // ✅ 항상 새 시트 생성

  // ▼ 개인 최대 수를 시트의 Z1에 저장(숨겨도 됨)
  sheet.getRange('Z1').setValue(perPersonLimit);

  fillRQContent_(sheet, year, month);    // 내용 채우기

  ui.alert('완료! 새 시트 "' + sheetName + '"를 생성했습니다.\n(개인 최대 신청 수: ' + perPersonLimit + ')');

}

// 동일 이름이 있으면 (1),(2)... 번호를 붙여 유일 이름 생성
function getUniqueSheetName_(base) {
  var ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(base)) return base;
  var i = 1;
  while (ss.getSheetByName(base + ' (' + i + ')')) i++;
  return base + ' (' + i + ')';
}

// 시트 내용 작성 (직원 목록 + 날짜/요일 + 공휴일/주말 색상 + 체크박스 + 요약행)
function fillRQContent_(sheet, year, month) {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!src) throw new Error('시트 "' + SOURCE_SHEET_NAME + '" 를 찾을 수 없습니다.');

  // ----- 좌측 상단 파라미터/라벨 -----
  sheet.getRange('A1').setValue('년도');
  sheet.getRange('B1').setValue(year);
  sheet.getRange('A2').setValue('월');
  sheet.getRange('B2').setValue(month);
  sheet.getRange('A3').setValue('일자');
  sheet.getRange('A4').setValue('요일');

  // ----- 직원 목록 (PARTNER 제외) + ID 오름차순 -----
  var lastRow = src.getLastRow();
  var raw = lastRow > 1 ? src.getRange(2, 1, lastRow - 1, 3).getValues() : []; // A:ID, B:이름, C:고용유형코드
  var rows = [];
  for (var i = 0; i < raw.length; i++) {
    var id = raw[i][0], name = raw[i][1], type = raw[i][2];
    if (id && name && String(type).toUpperCase() !== 'PARTNER') rows.push([id, name]);
  }
  rows.sort(function(a,b){ return a[0] > b[0] ? 1 : (a[0] < b[0] ? -1 : 0); });

  var names = [];
  for (i = 0; i < rows.length; i++) names.push([rows[i][1]]);

  // 기준점
  var NAME_START_ROW = 5;  // B5부터 이름
  var DAY_ROW = 3;         // 3행: 일자
  var YOIL_ROW = 4;        // 4행: 요일
  var FIRST_DAY_COL = 3;   // C열부터 날짜
  var nameCount = names.length;

  if (nameCount) sheet.getRange(NAME_START_ROW, 2, nameCount, 1).setValues(names); // B열에 이름

  // ----- 달력 헤더 -----
  var lastDay = new Date(year, month, 0).getDate();
  var weekdays = ['일','월','화','수','목','금','토'];
  var holidaySet = getKoreanHolidaySet_(year, month);

  var dateRow = [], yoilRow = [], colorRow = [];
  for (var d = 1; d <= lastDay; d++) {
    var dt = new Date(year, month - 1, d);
    var dow = dt.getDay();
    var isHoliday = holidaySet.has(d);
    var color = (dow === 0 || isHoliday) ? '#d93025' : (dow === 6) ? '#1a73e8' : '#000000';
    dateRow.push(d);
    yoilRow.push(weekdays[dow]);
    colorRow.push(color);
  }

  if (lastDay > 0) {
    sheet.getRange(DAY_ROW, FIRST_DAY_COL, 1, lastDay).setValues([dateRow]);
    sheet.getRange(YOIL_ROW, FIRST_DAY_COL, 1, lastDay).setValues([yoilRow]);
    sheet.getRange(DAY_ROW, FIRST_DAY_COL, 1, lastDay).setHorizontalAlignment('center').setFontWeight('bold').setFontColors([colorRow]);
    sheet.getRange(YOIL_ROW, FIRST_DAY_COL, 1, lastDay).setHorizontalAlignment('center').setFontColors([colorRow]);
  }

  // ----- 표 테두리/고정/크기 -----
  var lastNameRow = NAME_START_ROW + Math.max(nameCount, 1) - 1;
  var lastDayCol = FIRST_DAY_COL + lastDay - 1;

  sheet.getRange(YOIL_ROW, 2, Math.max(nameCount, 1) + 1, lastDay + 1)
       .setBorder(true, true, true, true, true, true);

  sheet.setFrozenRows(YOIL_ROW);
  sheet.setFrozenColumns(2);
  if (lastDay > 0) sheet.setColumnWidths(FIRST_DAY_COL, lastDay, 38);
  sheet.setColumnWidth(2, 140);
  if (nameCount) sheet.setRowHeights(NAME_START_ROW, nameCount, 24);
  sheet.autoResizeColumn(2);

  // ----- 체크박스 + 요약행 -----
  if (lastDay > 0 && nameCount > 0) {
    var checkRange = sheet.getRange(NAME_START_ROW, FIRST_DAY_COL, nameCount, lastDay);
    var cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(true).build();
    checkRange.setDataValidation(cbRule);
    checkRange.setValue(false); // 초기화

    // 집계 행: 신청자/가능/현황
    var applicantsRow = lastNameRow + 1;
    var capacityRow   = applicantsRow + 1;
    var statusRow     = applicantsRow + 2;

    sheet.getRange(applicantsRow, 2).setValue('RQ 신청자');
    sheet.getRange(capacityRow,   2).setValue('RQ 신청 가능일');
    sheet.getRange(statusRow,     2).setValue('RQ 현황');

    for (var c = FIRST_DAY_COL; c <= lastDayCol; c++) {
      var colA1 = columnToLetter_(c);
      sheet.getRange(applicantsRow, c)
           .setFormula('=COUNTIF(' + colA1 + '$' + NAME_START_ROW + ':' + colA1 + '$' + lastNameRow + ', TRUE)');
      sheet.getRange(capacityRow, c)
           .setFormula('=IFERROR(INDEX(matrix_RQ!$H:$H, MATCH(DATE($B$1,$B$2,' + colA1 + '$' + DAY_ROW + '), matrix_RQ!$C:$C, 0)), 0)');
      sheet.getRange(statusRow, c)
           .setFormula('=' + colA1 + '$' + capacityRow + '-' + colA1 + '$' + applicantsRow);
    }
  }
}

function onEdit(e) {
  // ✅ 이벤트 객체 존재 확인
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  
  // 대상: 이름이 "... RQ" 또는 "... RQ (n)" 로 끝나는 시트
  if (!/RQ(\s\(\d+\))?$/.test(sheetName)) return;

  var NAME_START_ROW = 5;
  var DAY_ROW = 3;
  var FIRST_DAY_COL = 3;

  // 이름 마지막 행(B열 연속 값)
  var bCol = sheet.getRange(NAME_START_ROW, 2, sheet.getMaxRows() - NAME_START_ROW + 1, 1).getValues();
  var lastNameRow = NAME_START_ROW - 1;
  for (var i = 0; i < bCol.length; i++) {
    if (bCol[i][0]) lastNameRow = NAME_START_ROW + i;
    else break;
  }
  if (lastNameRow < NAME_START_ROW) return;

  // 날짜 마지막 열(3행 연속 값)
  var header = sheet.getRange(DAY_ROW, FIRST_DAY_COL, 1, sheet.getMaxColumns() - FIRST_DAY_COL + 1).getValues()[0];
  var lastDay = 0;
  for (i = 0; i < header.length; i++) {
    if (header[i] === '' || header[i] === null) break;
    lastDay++;
  }
  if (lastDay === 0) return;
  var lastDayCol = FIRST_DAY_COL + lastDay - 1;

  // 편집 범위와 체크 그리드 교집합 확인
  var er = e.range.getRow(), ec = e.range.getColumn(), erh = e.range.getNumRows(), ech = e.range.getNumColumns();
  var gridTop = NAME_START_ROW, gridBottom = lastNameRow, gridLeft = FIRST_DAY_COL, gridRight = lastDayCol;
  
  if (er > gridBottom || ec > gridRight || (er + erh - 1) < gridTop || (ec + ech - 1) < gridLeft) return;

  // ✅ FALSE → TRUE로 변경된 셀만 수집
  var changed = [];
  
  for (var r = er; r < er + erh; r++) {
    for (var c = ec; c < ec + ech; c++) {
      if (r < gridTop || r > gridBottom || c < gridLeft || c > gridRight) continue;
      
      var currentValue = sheet.getRange(r, c).getValue();
      
      // 현재 값이 true이고, 이전 값이 true가 아닌 경우만
      if (currentValue === true) {
        if (erh === 1 && ech === 1) {
          // 단일 셀: oldValue 체크
          if (e.oldValue !== true) {
            changed.push([r, c]);
          }
        } else {
          // 다중 셀: 일단 수집
          changed.push([r, c]);
        }
      }
    }
  }
  
  if (changed.length === 0) return;

  // ✅ 현재 전체 그리드 상태 읽기
  var gridValues = sheet.getRange(gridTop, gridLeft, gridBottom - gridTop + 1, gridRight - gridLeft + 1).getValues();
  
  // 행/열별 카운트
  var rowCount = [], colCount = [];
  for (r = 0; r < gridValues.length; r++) {
    rowCount[r] = 0;
    for (c = 0; c < gridValues[0].length; c++) {
      if (gridValues[r][c] === true) {
        rowCount[r]++;
        colCount[c] = (colCount[c] || 0) + 1;
      }
    }
  }
  for (c = 0; c < gridValues[0].length; c++) {
    if (!colCount[c]) colCount[c] = 0;
  }

  // 날짜별 정원 조회
  var year  = Number(sheet.getRange('B1').getValue());
  var month = Number(sheet.getRange('B2').getValue());
  var capacityByCol = {};
  
  for (i = 0; i < changed.length; i++) {
    var cc = changed[i][1];
    if (capacityByCol[cc] == null) {
      var day = Number(sheet.getRange(DAY_ROW, cc).getValue());
      capacityByCol[cc] = getCapacityFromMatrix_(year, month, day);
    }
  }

  // 규칙 위반 체크
  var rowLimitViolated = false;
  var colLimitViolated = false;
  var violationMsg = '';

  // 개인 5개 초과?
  // 규칙 위반 체크
  var rowLimitViolated = false;
  var colLimitViolated = false;
  var violationMsg = '';

  // ▼ 개인 최대 개수: Z1에 저장된 값 사용 (없으면 5로 fallback)
  var perPersonLimitCell = Number(sheet.getRange('Z1').getValue());
  var perPersonLimit = (isFinite(perPersonLimitCell) && perPersonLimitCell >= 1 && perPersonLimitCell <= 13)
    ? perPersonLimitCell : 5;

  // 개인 최대 개수 초과?
  for (i = 0; i < changed.length; i++) {
    var rr = changed[i][0];
    var count = rowCount[rr - gridTop];

    if (count > perPersonLimit) {
      rowLimitViolated = true;
      var employeeName = sheet.getRange(rr, 2).getValue();
      violationMsg = employeeName + '님은 개인당 최대 ' + perPersonLimit + '일까지만 신청할 수 있습니다. (현재: ' + count + '개)';
      break;
    }
  }
  
  // 날짜 정원 초과?
  if (!rowLimitViolated) {
    for (i = 0; i < changed.length; i++) {
      cc = changed[i][1];
      var applicants = colCount[cc - gridLeft];
      var cap = Number(capacityByCol[cc]) || 0;
      var dayNum = sheet.getRange(DAY_ROW, cc).getValue();
      
      if (applicants > cap) { 
        colLimitViolated = true;
        violationMsg = dayNum + '일은 정원(' + cap + '명)이 초과되었습니다. (신청자: ' + applicants + '명)';
        break;
      }
    }
  }

  if (rowLimitViolated || colLimitViolated) {
    // ✅ 핵심: SpreadsheetApp.flush()를 먼저 호출하여 이전 변경사항을 확정
    SpreadsheetApp.flush();
    
    // ✅ 체크 되돌리기 - 각 셀을 개별적으로 처리
    for (i = 0; i < changed.length; i++) {
      var cell = sheet.getRange(changed[i][0], changed[i][1]);
      cell.setValue(false);
    }
    
    // ✅ 변경사항 즉시 반영
    SpreadsheetApp.flush();

    var msg = '❌ 신청 제한 초과\n\n' + violationMsg + '\n\n제한사항:\n· 개인: 최대 5일\n· 날짜별: matrix_RQ의 정원 내';
    
    // Alert 표시
    SpreadsheetApp.getUi().alert('RQ 신청 제한', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * matrix_RQ 시트에서 (연,월,일) 일치 행의 정원(H열) 값을 반환
 */
function getCapacityFromMatrix_(year, month, day) {
  var ms = SpreadsheetApp.getActive().getSheetByName('matrix_RQ');
  if (!ms) return 0;

  var lastRow = ms.getLastRow();
  if (lastRow < 2) return 0;

  var dates = ms.getRange(2, 3, lastRow - 1, 1).getValues(); // C2:C
  var caps  = ms.getRange(2, 8, lastRow - 1, 1).getValues(); // H2:H

  for (var i = 0; i < dates.length; i++) {
    var d = dates[i][0];
    
    // Date 객체로 변환
    if (!(d instanceof Date)) {
      d = new Date(d);
    }
    
    if (d && d.getFullYear() === year && (d.getMonth() + 1) === month && d.getDate() === day) {
      var n = Number(caps[i][0]);
      return isFinite(n) ? n : 0;
    }
  }
  
  return 0;
}

// 숫자열 → A1 열문자
function columnToLetter_(column) {
  var temp = '';
  var col = column;
  while (col > 0) {
    var rem = (col - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    col = Math.floor((col - 1) / 26);
  }
  return temp;
}

/**
 * matrix_RQ 시트에서 (연,월,일) 일치 행의 정원(H열) 값을 반환
 * - C열: 날짜(Date)
 * - H열: 정원(숫자)
 * - 없으면 0
 */
function getCapacityFromMatrix_(year, month, day) {
  var ms = SpreadsheetApp.getActive().getSheetByName('matrix_RQ');
  if (!ms) return 0;

  var lastRow = ms.getLastRow();
  if (lastRow < 2) return 0;

  var dates = ms.getRange(2, 3, lastRow - 1, 1).getValues(); // C2:C
  var caps  = ms.getRange(2, 8, lastRow - 1, 1).getValues(); // H2:H

  for (var i = 0; i < dates.length; i++) {
    var d = dates[i][0] instanceof Date ? dates[i][0] : new Date(dates[i][0]);
    if (d && d.getFullYear() === year && (d.getMonth() + 1) === month && d.getDate() === day) {
      var n = Number(caps[i][0]);
      return isFinite(n) ? n : 0;
    }
  }
  return 0;
}

// 해당 연/월의 한국 공휴일을 day(Set)로 반환
function getKoreanHolidaySet_(year, month) {
  var cal = CalendarApp.getCalendarById(KR_HOLIDAY_CAL_ID);
  var start = new Date(year, month - 1, 1);
  var end = new Date(year, month, 1);
  var events = cal.getEvents(start, end);
  var set = new Set();

  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    var s = ev.isAllDayEvent() ? ev.getAllDayStartDate() : ev.getStartTime();
    var e = ev.isAllDayEvent() ? ev.getAllDayEndDate() : ev.getEndTime();
    var d = new Date(s.getFullYear(), s.getMonth(), s.getDate());
    while (d < e) {
      if (d.getFullYear() === year && (d.getMonth() + 1) === month) set.add(d.getDate());
      d.setDate(d.getDate() + 1);
    }
  }
  return set;
}

// 숫자열 → A1 열문자
function columnToLetter_(column) {
  var temp = '';
  var col = column;
  while (col > 0) {
    var rem = (col - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    col = Math.floor((col - 1) / 26);
  }
  return temp;
}
