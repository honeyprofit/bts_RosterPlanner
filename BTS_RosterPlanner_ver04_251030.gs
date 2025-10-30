/*** CONFIG ***/
const SOURCE_SHEET_NAME = "matrix_employees";
const KR_HOLIDAY_CAL_ID = "ko.south_korea#holiday@group.v.calendar.google.com";
const MON_ABBR = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];

/*** 메뉴 등록 ***/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("RQ")
    .addItem("신규 월간 RQ 생성", "rqCreateNewSheet")
    .addItem("RQ 데이터 확정", "rqConfirmMenu")
    .addToUi();
  
  ui.createMenu("Schedule")
    .addItem("월별 스케줄 자동 배정", "scheduleAutoAssignMenu")
    .addItem("스케줄 검증", "scheduleValidateMenu")
    .addItem("스케줄 시각화 시트 생성", "scheduleVisualizeMenu")
    .addToUi();
}

/*** 신규 월간 RQ 생성 ***/
function rqCreateNewSheet() {
  const ui = SpreadsheetApp.getUi();

  // 연도 입력 받기 (예: 2025)
  const y = ui.prompt("연도 입력", "예: 2025", ui.ButtonSet.OK_CANCEL);
  if (y.getSelectedButton() !== ui.Button.OK) return;
  const year = Number(String(y.getResponseText()).trim());
  if (!isFinite(year) || year < 1900 || year > 2100) {
    ui.alert("연도 입력이 올바르지 않습니다. (예: 2025)");
    return;
  }

  // 월 입력 받기 (1~12)
  const m = ui.prompt("월 입력", "1 ~ 12 중 하나 (예: 10)", ui.ButtonSet.OK_CANCEL);
  if (m.getSelectedButton() !== ui.Button.OK) return;
  const month = Number(String(m.getResponseText()).trim());
  if (!isFinite(month) || month < 1 || month > 12) {
    ui.alert("월 입력이 올바르지 않습니다. (1~12)");
    return;
  }

  // 개인 최대 신청 가능 수 입력 (1~13)
  const l = ui.prompt("개인 최대 RQ 신청 가능 수 입력", "1 ~ 13 (예: 5)", ui.ButtonSet.OK_CANCEL);
  if (l.getSelectedButton() !== ui.Button.OK) return;
  const perPersonLimit = Number(String(l.getResponseText()).trim());
  if (!isFinite(perPersonLimit) || perPersonLimit < 1 || perPersonLimit > 13) {
    ui.alert("개인 최대 수는 1~13 사이의 정수여야 합니다.");
    return;
  }

  // 시트명 생성: "YY MON RQ" 형식 (중복 시 (n) 붙임)
  const baseName = String(year).slice(-2) + " " + MON_ABBR[month - 1] + " RQ";
  const sheetName = getUniqueSheetName_(baseName);
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.insertSheet(sheetName);
  
  // Z1에 개인 최대 신청 수 저장 (onEdit 검증에서 사용)
  sheet.getRange("Z1").setValue(perPersonLimit);
  
  // 본문 콘텐츠 구성 (헤더/체크박스/요약행 등)
  fillRQContent_(sheet, year, month);
  
  ui.alert('완료! 새 시트 "' + sheetName + '"를 생성했습니다.\n(개인 최대 신청 수: ' + perPersonLimit + ")");
}

function getUniqueSheetName_(base) {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(base)) return base;
  let i = 1;
  // 동일 이름이 있으면 "base (n)" 순번 증가
  while (ss.getSheetByName(base + " (" + i + ")")) i++;
  return base + " (" + i + ")";
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
  // src: matrix_employees 시트에서 A:ID, B:이름, C:고용유형코드 사용
  var lastRow = src.getLastRow();
  var raw = lastRow > 1 ? src.getRange(2, 1, lastRow - 1, 3).getValues() : []; // A:ID, B:이름, C:고용유형코드
  var rows = [];
  for (var i = 0; i < raw.length; i++) {
    var id = raw[i][0], name = raw[i][1], type = raw[i][2];
    if (id && name && String(type).toUpperCase() !== 'PARTNER') rows.push([id, name]);
  }
  rows.sort(function(a,b){ return a[0] > b[0] ? 1 : (a[0] < b[0] ? -1 : 0); });

  // 이름 배열만 분리
  var names = [];
  for (i = 0; i < rows.length; i++) names.push([rows[i][1]]);

  // 기준점 (셀 위치 상수)
  var NAME_START_ROW = 5;  // B5부터 이름 목록 시작
  var DAY_ROW = 3;         // 3행: 일자 숫자
  var YOIL_ROW = 4;        // 4행: 요일 문자
  var FIRST_DAY_COL = 3;   // C열부터 날짜
  var nameCount = names.length;

  // 이름 렌더링
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
    // 일요일/공휴일: 빨강, 토요일: 파랑, 평일: 검정
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
    // 개별 신청 체크박스
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

    // 각 날짜별: 신청자 수 / 정원 / 잔여
    for (var c = FIRST_DAY_COL; c <= lastDayCol; c++) {
      var colA1 = columnToLetter_(c);
      // TRUE 체크 수
      sheet.getRange(applicantsRow, c)
           .setFormula('=COUNTIF(' + colA1 + '$' + NAME_START_ROW + ':' + colA1 + '$' + lastNameRow + ', TRUE)');
      // matrix_RQ에서 해당 날짜 정원 H열 참조
      sheet.getRange(capacityRow, c)
           .setFormula('=IFERROR(INDEX(matrix_RQ!$H:$H, MATCH(DATE($B$1,$B$2,' + colA1 + '$' + DAY_ROW + '), matrix_RQ!$C:$C, 0)), 0)');
      // 잔여 정원
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
  
  // 대상: 이름이 "... RQ" 또는 "... RQ (n)" 로 끝나는 시트만 감시
  if (!/RQ(\s\(\d+\))?$/.test(sheetName)) return;

  var NAME_START_ROW = 5, DAY_ROW = 3, FIRST_DAY_COL = 3;

  // 이름 마지막 행(B열 연속 값) 탐색 (빈칸 만나기 전까지)
  var bCol = sheet.getRange(NAME_START_ROW, 2, sheet.getMaxRows() - NAME_START_ROW + 1, 1).getValues();
  var lastNameRow = NAME_START_ROW - 1;
  for (var i = 0; i < bCol.length; i++) {
    if (bCol[i][0]) lastNameRow = NAME_START_ROW + i;
    else break;
  }
  if (lastNameRow < NAME_START_ROW) return;

  // 날짜 마지막 열(3행 연속 값) 탐색
  var header = sheet.getRange(DAY_ROW, FIRST_DAY_COL, 1, sheet.getMaxColumns() - FIRST_DAY_COL + 1).getValues()[0];
  var lastDay = 0;
  for (i = 0; i < header.length; i++) {
    if (header[i] === '' || header[i] === null) break;
    lastDay++;
  }
  if (lastDay === 0) return;
  var lastDayCol = FIRST_DAY_COL + lastDay - 1;

  // 편집 범위와 체크 그리드 교집합 확인 (그리드 밖 수정을 무시)
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
          // 단일 셀: oldValue 체크 (TRUE로 바뀐 경우만)
          if (e.oldValue !== true) 
          changed.push([r, c]);
        } else {
          // 다중 셀: 현재 TRUE인 셀 전체 수집 (이전값 추적 불가)
          changed.push([r, c]);
        }
      }
    }
  }
  
  if (changed.length === 0) return;

  // ✅ 현재 전체 그리드 상태 읽기 (행/열 카운트 계산용)
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

  // 날짜별 정원 조회 (변경된 열만 계산하여 최적화)
  var year = Number(sheet.getRange('B1').getValue());
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

  // ▼ 개인 최대 개수: Z1에 저장된 값 사용 (없으면 5로 fallback)
  var perPersonLimitCell = Number(sheet.getRange('Z1').getValue());
  var perPersonLimit = (isFinite(perPersonLimitCell) && perPersonLimitCell >= 1 && perPersonLimitCell <= 13) ? perPersonLimitCell : 5;

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
    // 핵심: SpreadsheetApp.flush()를 먼저 호출하여 이전 변경사항을 확정
    SpreadsheetApp.flush();

    // 체크 되돌리기 - 각 셀을 개별적으로 처리 (일괄 setValues 대신 안정성 확보)
    for (i = 0; i < changed.length; i++) {
      var cell = sheet.getRange(changed[i][0], changed[i][1]);
      cell.setValue(false);
    }

    // 변경사항 즉시 반영
    SpreadsheetApp.flush();

    var msg = '❌ 신청 제한 초과\n\n' + violationMsg + '\n\n제한사항:\n· 개인: 최대 5일\n· 날짜별: matrix_RQ의 정원 내';

    // Alert 표시
    SpreadsheetApp.getUi().alert('RQ 신청 제한', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * matrix_RQ 시트에서 (연,월,일) 일치 행의 정원(H열) 값을 반환
 * - 일치하는 날짜가 없으면 0 반환
 * - 데이터가 문자열이면 Date로 변환 시도
 */
function getCapacityFromMatrix_(year, month, day) {
  const ms = SpreadsheetApp.getActive().getSheetByName("matrix_RQ");
  if (!ms) return 0;
  const lastRow = ms.getLastRow();
  if (lastRow < 2) return 0;
  const dates = ms.getRange(2, 3, lastRow - 1, 1).getValues();
  const caps = ms.getRange(2, 8, lastRow - 1, 1).getValues();
  for (let i = 0; i < dates.length; i++) {
    let d = dates[i][0];
    if (!(d instanceof Date)) d = new Date(d);

    // Date 객체로 변환
    if (d && d.getFullYear() === year && d.getMonth() + 1 === month && d.getDate() === day) {
      const n = Number(caps[i][0]);
      return isFinite(n) ? n : 0;
    }
  }
  return 0;
}

/**
 * 구글 공휴일 캘린더에서 해당 연/월의 공휴일을 Set(day)로 반환
 * - 종일 이벤트의 경우 날짜 범위를 하루씩 증가시키며 포함
 */
function getKoreanHolidaySet_(year, month) {
  const cal = CalendarApp.getCalendarById(KR_HOLIDAY_CAL_ID);
  const start = new Date(year, month - 1, 1);
  const end = new Date(year, month, 1);
  const events = cal.getEvents(start, end);
  const set = new Set();
  for (let i = 0; i < events.length; i++) {
    const ev = events[i];
    const s = ev.isAllDayEvent() ? ev.getAllDayStartDate() : ev.getStartTime();
    const e = ev.isAllDayEvent() ? ev.getAllDayEndDate() : ev.getEndTime();
    const d = new Date(s.getFullYear(), s.getMonth(), s.getDate());
    while (d < e) {
      if (d.getFullYear() === year && d.getMonth() + 1 === month) set.add(d.getDate());
      d.setDate(d.getDate() + 1);
    }
  }
  return set;
}

/**
 * 숫자 열 인덱스를 A1 표기법의 열 문자로 변환 (1->A, 27->AA 등)
 */
function columnToLetter_(column) {
  let temp = "", col = column;
  while (col > 0) {
    const rem = (col - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    col = Math.floor((col - 1) / 26);
  }
  return temp;
}

/**
 * RQ 시트의 체크 결과를 DB_leave로 확정 반영하는 메뉴 엔트리
 * - 여러 RQ 시트 중 하나 선택
 */
function rqConfirmMenu() {
  const ss = SpreadsheetApp.getActive();
  const rqNames = ss.getSheets().map(s => s.getName()).filter(n => /RQ(\s\(\d+\))?$/.test(n));
  if (rqNames.length === 0) {
    SpreadsheetApp.getUi().alert("확정 가능한 RQ 시트가 없습니다.");
    return;
  }
  // 번호 목록 UI
  let msg = "확정할 RQ 시트를 선택하세요.\n\n";
  for (let i = 0; i < rqNames.length; i++) msg += i + 1 + ". " + rqNames[i] + "\n";
  msg += "\n번호를 입력하세요:";
  const res = SpreadsheetApp.getUi().prompt("RQ 데이터 확정", msg, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return;
  const idx = Number(String(res.getResponseText()).trim());
  if (!isFinite(idx) || idx < 1 || idx > rqNames.length) {
    SpreadsheetApp.getUi().alert("올바른 번호가 아닙니다.");
    return;
  }
  confirmRQDataToDB_(rqNames[idx - 1]);
}

/**
 * 선택한 RQ 시트의 체크박스 상태를 DB_leave에 행 단위로 적재
 * - DB_leave 헤더 검증
 * - (year, month, date, day, name, is_rq, is_annual_leave, is_business_trip)
 */
function confirmRQDataToDB_(sheetName) {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert("RQ 데이터 확정", sheetName + "의 RQ 데이터를 DB_leave에 업데이트 하시겠습니까?", ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActive();
  const rqSheet = ss.getSheetByName(sheetName);
  const dbSheet = ss.getSheetByName("DB_leave");
  if (!rqSheet || !dbSheet) {
    ui.alert("오류: RQ 시트 또는 DB_leave 시트를 찾을 수 없습니다.");
    return;
  }
  const year = Number(rqSheet.getRange("B1").getValue());
  const month = Number(rqSheet.getRange("B2").getValue());
  const FIRST_DAY_COL = 3, NAME_START_ROW = 5;
  const namesCol = rqSheet.getRange(NAME_START_ROW, 2, rqSheet.getLastRow() - NAME_START_ROW + 1, 1).getValues().flat();
  const names = [];
  for (let i = 0; i < namesCol.length - 3; i++) {
    if (namesCol[i]) names.push(namesCol[i]);
    else break;
  }
  if (names.length === 0) {
    ui.alert("해당 RQ 시트에서 이름을 찾지 못했습니다.");
    return;
  }

  // 날짜 헤더 추출
  const headerRow = rqSheet.getRange(3, FIRST_DAY_COL, 1, rqSheet.getLastColumn() - FIRST_DAY_COL + 1).getValues()[0];
  const days = [];
  for (let i = 0; i < headerRow.length; i++) {
    if (headerRow[i] === "" || headerRow[i] === null) break;
    days.push(Number(headerRow[i]));
  }
  if (days.length === 0) {
    ui.alert("해당 RQ 시트에서 날짜 헤더를 찾지 못했습니다.");
    return;
  }

  // 체크 그리드 읽기
  const grid = rqSheet.getRange(NAME_START_ROW, FIRST_DAY_COL, names.length, days.length).getValues();

  // DB_leave 헤더 검증
  const expectedHeader = ["year","month","date","day","name","is_rq","is_annual_leave","is_business_trip"];
  const dbHeader = dbSheet.getRange(1, 1, 1, expectedHeader.length).getValues()[0];
  for (let i = 0; i < expectedHeader.length; i++) {
    if (dbHeader[i] !== expectedHeader[i]) {
      ui.alert("DB_leave 헤더가 예상과 다릅니다. 다음과 같아야 합니다:\n" + expectedHeader.join(", "));
      return;
    }
  }

  // 적재 데이터 생성
  const weekdays = ["일","월","화","수","목","금","토"];
  const out = [];
  for (let r = 0; r < names.length; r++) {
    const emp = names[r];
    for (let d = 0; d < days.length; d++) {
      const dayNum = days[d];
      const dateObj = new Date(year, month - 1, dayNum);
      const yoil = weekdays[dateObj.getDay()];
      const isRQ = grid[r][d] === true;
      out.push([year, month, dayNum, yoil, emp, isRQ, "FALSE", "FALSE"]);
    }
  }

  // DB_leave의 첫 빈 행 탐색 후 한번에 적재
  const startRow = getFirstEmptyRow_(dbSheet, 1);
  dbSheet.getRange(startRow, 1, out.length, out[0].length).setValues(out);
  ui.alert('✅ "' + sheetName + '" 확정 완료\nDB_leave에 ' + out.length + "행을 추가했습니다.");
}

/**
 * 특정 열 기준으로 첫 빈 행의 번호 반환 (헤더 1행 가정, 데이터는 2행부터)
 * - 끝에서부터 역방향으로 검색하여 성능 개선
 */
function getFirstEmptyRow_(sheet, colIndex) {
  const max = sheet.getMaxRows();
  const vals = sheet.getRange(1, colIndex, max).getValues();
  for (let r = max - 1; r >= 1; r--) {
    if (String(vals[r][0]).trim() !== "") return r + 2;
  }
  return 2;
}

/* ================================================================
   ✅ Schedule 자동 배정
   - DB_schedule의 특정 월에 대해 제약조건을 고려하여 스케줄(ON/OFF/A계열) 자동 생성
================================================================ */

function scheduleAutoAssignMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const dbSheet = ss.getSheetByName("DB_schedule");
  
  if (!dbSheet) {
    ui.alert("DB_schedule 시트를 찾을 수 없습니다.");
    return;
  }
  
  const data = dbSheet.getDataRange().getValues();
  const headers = data[0];
  const monthCol = headers.indexOf("month");
  const scheduleCol = headers.indexOf("schedule");
  
  if (monthCol === -1 || scheduleCol === -1) {
    ui.alert("DB_schedule에 month 또는 schedule 열이 없습니다.");
    return;
  }
  
  // 스케줄이 비어있는 월 목록 수집
  const emptyMonths = new Set();
  for (let i = 1; i < data.length; i++) {
    const month = data[i][monthCol];
    const schedule = data[i][scheduleCol];
    if (month && (!schedule || String(schedule).trim() === "")) {
      emptyMonths.add(Number(month));
    }
  }
  
  if (emptyMonths.size === 0) {
    ui.alert("스케줄이 비어있는 월이 없습니다.");
    return;
  }
    
  // 사용자에게 대상 월 선택 받기
  const monthList = Array.from(emptyMonths).sort((a, b) => a - b);
  let msg = "스케줄을 배정할 월을 선택하세요.\n\n";
  for (let i = 0; i < monthList.length; i++) {
    msg += (i + 1) + ". " + monthList[i] + "월\n";
  }
  msg += "\n번호를 입력하세요:";
  
  const res = ui.prompt("월별 스케줄 자동 배정", msg, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  
  const idx = Number(String(res.getResponseText()).trim());
  if (!isFinite(idx) || idx < 1 || idx > monthList.length) {
    ui.alert("올바른 번호가 아닙니다.");
    return;
  }
  
  const targetMonth = monthList[idx - 1];

  // 경고/안내 메시지: 복잡한 제약으로 완벽하지 않을 수 있음
  const confirm = ui.alert("스케줄 자동 배정", targetMonth + "월의 스케줄을 자동으로 배정하시겠습니까?\n\n경고: 복잡한 제약조건으로 인해 완벽하지 않을 수 있습니다.\n배정 후 '스케줄 검증'을 실행하여 확인하세요.", ui.ButtonSet.YES_NO);
  
  if (confirm !== ui.Button.YES) return;
  
  try {
    assignScheduleForMonth_(targetMonth);
    ui.alert("✅ " + targetMonth + "월 스케줄 배정이 완료되었습니다.\n\n다음 단계:\n1. 'Schedule > 스케줄 검증' 메뉴를 실행하세요.\n2. 위반사항이 있다면 수동으로 조정하세요.");
  } catch (error) {
    ui.alert("❌ 오류 발생: " + error.toString());
  }
}

/**
 * 특정 월(targetMonth)의 DB_schedule을 읽어 자동 배정 로직 수행 후 결과 기록
 * - 입력: DB_schedule, matrix_workday, matrix_RQ 등
 * - 출력: DB_schedule.schedule 컬럼 채움
 */
function assignScheduleForMonth_(targetMonth) {
  const ss = SpreadsheetApp.getActive();
  const dbSheet = ss.getSheetByName("DB_schedule");
  const data = dbSheet.getDataRange().getValues();
  const headers = data[0];
  
  // 필수 컬럼 인덱스 매핑
  const colIdx = {};
  const requiredCols = ["year","month","date","day","name","is_rq","employment_type_code","driving_class","gender_code","is_disposal_day","is_hh_cleaning_day","is_available_day","is_business_trip","schedule"];
  
  for (let col of requiredCols) {
    colIdx[col] = headers.indexOf(col);
    if (colIdx[col] === -1) throw new Error("필수 열을 찾을 수 없습니다: " + col);
  }
  
  // 대상 월 데이터만 필터링
  const monthData = [];
  const rowIndices = [];
  
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][colIdx.month]) === targetMonth) {
      monthData.push(data[i]);
      rowIndices.push(i + 1); // 시트 기록용 실제 행 번호
    }
  }
  
  if (monthData.length === 0) throw new Error(targetMonth + "월 데이터가 없습니다.");
  
  const year = monthData[0][colIdx.year];

  // 직원/날짜 구조 구성
  const employees = {};
  const dateInfo = {};
  
  for (let row of monthData) {
    const name = row[colIdx.name];
    const date = Number(row[colIdx.date]);
    
    if (!employees[name]) {
      employees[name] = {
        name: name,
        employment_type: row[colIdx.employment_type_code],
        driving_class: row[colIdx.driving_class],
        gender: row[colIdx.gender_code],
        days: {}
      };
    }
    
    employees[name].days[date] = {
      is_rq: row[colIdx.is_rq] === true || row[colIdx.is_rq] === "TRUE",
      is_available: row[colIdx.is_available_day] === true || row[colIdx.is_available_day] === "TRUE",
      is_business_trip: row[colIdx.is_business_trip] === true || row[colIdx.is_business_trip] === "TRUE",
      day_of_week: row[colIdx.day]
    };
    
    if (!dateInfo[date]) {
      dateInfo[date] = {
        day_of_week: row[colIdx.day],
        is_disposal: row[colIdx.is_disposal_day] === true || row[colIdx.is_disposal_day] === "TRUE",
        is_hh_cleaning: row[colIdx.is_hh_cleaning_day] === true || row[colIdx.is_hh_cleaning_day] === "TRUE"
      };
    }
  }
  
  // 전월 연속 근무 영향, 월별 목표 근무일 수, 일자별 필요 인원 로딩
  const prevMonthData = getPreviousMonthWorkDays_(year, targetMonth, employees);
  const workdayRequirements = getWorkdayRequirements_(year, targetMonth);
  const dailyStaffRequirements = getDailyStaffRequirements_(year, targetMonth);
  
  // 핵심: 최적화 스케줄 생성
  const schedule = generateOptimizedSchedule_(employees, dateInfo, year, targetMonth, prevMonthData, workdayRequirements, dailyStaffRequirements);
  
  // 결과 기록 (값 없으면 OFF로 폴백)
  for (let i = 0; i < monthData.length; i++) {
    const name = monthData[i][colIdx.name];
    const date = Number(monthData[i][colIdx.date]);
    const scheduleValue = schedule[name] && schedule[name][date] ? schedule[name][date] : "OFF";
    dbSheet.getRange(rowIndices[i], colIdx.schedule + 1).setValue(scheduleValue);
  }
}

/**
 * 전월의 직원별 근무(ON/OFF) 이력 추출 (연속 근무 판단에 사용)
 * - 출력: { 이름: [{date, schedule}, ...] } (date 오름차순)
 */
function getPreviousMonthWorkDays_(year, month, employees) {
  const ss = SpreadsheetApp.getActive();
  const dbSheet = ss.getSheetByName("DB_schedule");
  if (!dbSheet) return {};
  
  const data = dbSheet.getDataRange().getValues();
  const headers = data[0];
  const yearCol = headers.indexOf("year");
  const monthCol = headers.indexOf("month");
  const dateCol = headers.indexOf("date");
  const nameCol = headers.indexOf("name");
  const scheduleCol = headers.indexOf("schedule");
  
  if (yearCol === -1 || monthCol === -1 || dateCol === -1 || nameCol === -1 || scheduleCol === -1) return {};
  
  // 전월 계산 (1월이면 전년도 12월)
  let prevYear = year, prevMonth = month - 1;
  if (prevMonth < 1) {
    prevMonth = 12;
    prevYear--;
  }
  
  const result = {};
  
  for (let name in employees) {
    result[name] = [];
  }
  
  // 전월 데이터 수집
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (Number(row[yearCol]) === prevYear && Number(row[monthCol]) === prevMonth) {
      const name = row[nameCol];
      const date = Number(row[dateCol]);
      const schedule = row[scheduleCol];
      
      if (result[name] !== undefined) {
        result[name].push({
          date: date,
          schedule: schedule === "ON" ? "ON" : "OFF"
        });
      }
    }
  }
  
  // 일자 오름차순 정렬
  for (let name in result) {
    result[name].sort((a, b) => a.date - b.date);
  }
  
  return result;
}

/**
 * 이름별 월 목표 근무일 수 로드 (matrix_workday)
 * - 키: name, 값: workdays
 */
function getWorkdayRequirements_(year, month) {
  const ss = SpreadsheetApp.getActive();
  const workdaySheet = ss.getSheetByName("matrix_workday");
  if (!workdaySheet) return {};
  
  const data = workdaySheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf("name");
  const yearCol = headers.indexOf("year");
  const monthCol = headers.indexOf("month");
  const workdaysCol = headers.indexOf("workdays");
  
  if (nameCol === -1 || yearCol === -1 || monthCol === -1 || workdaysCol === -1) return {};
  
  const requirements = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (Number(row[yearCol]) === year && Number(row[monthCol]) === month) {
      const name = row[nameCol];
      const workdays = Number(row[workdaysCol]);
      if (name && isFinite(workdays)) {
        requirements[name] = workdays;
      }
    }
  }
  
  return requirements;
}

/**
 * 일자별 필요 총 인원 로드 (matrix_RQ.total_staff_required)
 * - 키: date(숫자), 값: 필요 인원
 */
function getDailyStaffRequirements_(year, month) {
  const ss = SpreadsheetApp.getActive();
  const rqSheet = ss.getSheetByName("matrix_RQ");
  if (!rqSheet) return {};
  
  const data = rqSheet.getDataRange().getValues();
  const headers = data[0];
  const yearCol = headers.indexOf("year");
  const monthCol = headers.indexOf("month");
  const dateCol = headers.indexOf("date");
  const staffCol = headers.indexOf("total_staff_required");
  
  if (yearCol === -1 || monthCol === -1 || dateCol === -1 || staffCol === -1) return {};
  
  const requirements = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (Number(row[yearCol]) === year && Number(row[monthCol]) === month) {
      // const date = Number(row[dateCol]);
      const date = row[dateCol] instanceof Date ? row[dateCol].getDate() : Number(row[dateCol]);
      const staff = Number(row[staffCol]);
      if (isFinite(date) && isFinite(staff)) {
        requirements[date] = staff;
      }
    }
  }
  
  return requirements;
}

/**
//daily 필요인력 검증
// function testDailyStaffReq() {
//   const res = getDailyStaffRequirements_(2025, 11);
//   Logger.log(JSON.stringify(res, null, 2));
// }
*/

/**
 * 스케줄 최적화 생성기
 * - 규칙 요약:
 *   1) RQ 신청일은 반드시 OFF
 *   2) SP/SSV는 is_available_day가 TRUE일 때만 ON
 *   3) PARTNER는 is_rq만 OFF, 그 외 가능일은 ON 후보
 *   4) SV/FT는 is_available_day가 FALSE면 OFF
 *   5) 전월 연속 근무 고려, 연속 근무 5일 초과 금지
 *   6) 일별 최소 인원/운전자/슈퍼바이저/남성(HH/폐기물) 충족
 *   7) 정원 부족 시 PARTNER 활용
 *   8) 근무일 수 목표(workdayRequirements) 정확히 맞추기
 */
function generateOptimizedSchedule_(employees, dateInfo, year, month, prevMonthData, workdayRequirements, dailyStaffRequirements) {
  const daysInMonth = new Date(year, month, 0).getDate();
  const schedule = {};
  
  // 초기화 (기본 OFF)
  for (let name in employees) {
    schedule[name] = {};
    for (let d = 1; d <= daysInMonth; d++) {
      schedule[name][d] = "OFF";
    }
  }
  
  // ✅ 최우선 1: 모든 직원의 is_rq = TRUE 날은 반드시 OFF
  for (let name in employees) {
    for (let d = 1; d <= daysInMonth; d++) {
      if (employees[name].days[d] && employees[name].days[d].is_rq) {
        schedule[name][d] = "OFF";
      }
    }
  }
  
  // ✅ 규칙 12 개선: SP, SSV는 is_available_day가 TRUE인 날만 ON
  // is_available_day = NOT(is_dayoff OR is_business_trip OR is_rq 등)
  for (let name in employees) {
    if (employees[name].employment_type === "SP" || employees[name].employment_type === "SSV") {
      for (let d = 1; d <= daysInMonth; d++) {
        if (employees[name].days[d]) {
          const dayData = employees[name].days[d];
          
          // is_rq는 이미 위에서 OFF 처리됨
          // is_available_day가 TRUE인 날만 ON
          if (dayData.is_available && !dayData.is_rq) {
            schedule[name][d] = "ON";
          } else {
            schedule[name][d] = "OFF";
          }
        }
      }
    }
  }
  
  // 규칙 11: PARTNER는 is_rq만 OFF (나머지는 일단 ON, 나중에 조정)
  for (let name in employees) {
    if (employees[name].employment_type === "PARTNER") {
      for (let d = 1; d <= daysInMonth; d++) {
        if (employees[name].days[d] && employees[name].days[d].is_rq) {
          schedule[name][d] = "OFF";
        } else if (employees[name].days[d] && employees[name].days[d].is_available) {
          schedule[name][d] = "ON";
        }
      }
    }
  }
  
  // is_available_day가 FALSE인 경우 OFF (정직원 SV/FT)
  for (let name in employees) {
    const empType = employees[name].employment_type;
    if (empType === "SV" || empType === "FT") {
      for (let d = 1; d <= daysInMonth; d++) {
        if (employees[name].days[d] && !employees[name].days[d].is_available) {
          schedule[name][d] = "OFF";
        }
      }
    }
  }
  
  // ✅ 개선: 정직원(SV, FT) 스케줄 생성 - 근무일 수 반드시 만족
  const regularEmployees = Object.keys(employees).filter(name => 
    (employees[name].employment_type === "SV" || employees[name].employment_type === "FT")
  );
  
  // 1단계: 필수 근무일 먼저 배정 (출장일은 ON 고정)
  for (let name of regularEmployees) {
    for (let d = 1; d <= daysInMonth; d++) {
      if (employees[name].days[d] && employees[name].days[d].is_business_trip) {
        schedule[name][d] = "ON";
      }
    }
  }
  
  // 2단계: 각 직원별로 목표 근무일 수에 맞추어 배정 (연속/가용성 고려)
  for (let name of regularEmployees) {
    const targetWorkdays = workdayRequirements[name] || 22;
    let currentWorkdays = countWorkdays_(schedule, name);
    
    // 이미 배정된 근무일 (출장 등)을 기준으로 남은 필요 근무일 계산
    const fixedWorkdays = currentWorkdays;
    const remainingWorkNeeded = targetWorkdays - fixedWorkdays;
    
    if (remainingWorkNeeded <= 0) {
      // 이미 목표 달성 또는 초과 (출장이 많은 경우)
      continue;
    }
    
    // 사용 가능한 날짜 목록 생성 (현재 OFF & 가용 & RQ 아님)
    const availableDays = [];
    for (let d = 1; d <= daysInMonth; d++) {
      if (schedule[name][d] === "OFF" && 
          employees[name].days[d] && 
          employees[name].days[d].is_available &&
          !employees[name].days[d].is_rq) {
        availableDays.push(d);
      }
    }
    
    // 전월 연속 근무일 수 확인 (월초 연속 판단에 사용)
    let prevConsecutiveWork = 0;
    if (prevMonthData[name] && prevMonthData[name].length > 0) {
      for (let i = prevMonthData[name].length - 1; i >= 0; i--) {
        if (prevMonthData[name][i].schedule === "ON") {
          prevConsecutiveWork++;
        } else {
          break;
        }
      }
    }
    
    // ✅ 근무일 수를 정확히 맞추기 위한 최적화 알고리즘
    const workPattern = generateWorkPattern_(
      name,
      availableDays,
      remainingWorkNeeded,
      prevConsecutiveWork,
      daysInMonth
    );
    
    // 패턴 적용
    for (let d of workPattern) {
      schedule[name][d] = "ON";
    }
    
    // 최종 확인 및 미세 조정 (부족분 추가/초과분 제거)
    currentWorkdays = countWorkdays_(schedule, name);
    
    // 부족한 경우 추가 투입 (연속 5일 제한 및 가용성 체크)
    while (currentWorkdays < targetWorkdays) {
      let added = false;
      for (let d = 1; d <= daysInMonth; d++) {
        if (schedule[name][d] === "OFF" && 
            employees[name].days[d] && 
            employees[name].days[d].is_available &&
            !employees[name].days[d].is_rq &&
            canWorkOnDay_(schedule, name, d, prevMonthData)) {
          schedule[name][d] = "ON";
          currentWorkdays++;
          added = true;
          if (currentWorkdays >= targetWorkdays) break;
        }
      }
      if (!added) break;
    }
    
    // 초과한 경우 제거 (월말부터 역순으로 제거하여 말일 인력 보전)
    while (currentWorkdays > targetWorkdays) {
      let removed = false;
      // 뒤에서부터 제거 (월말 인원 부족 방지)
      for (let d = daysInMonth; d >= 1; d--) {
        if (schedule[name][d] === "ON" && 
            employees[name].days[d] && 
            !employees[name].days[d].is_business_trip &&
            !employees[name].days[d].is_rq) {
          
          // 단순 제거 (연속성은 다른 날로 상쇄될 여지 있음)
          schedule[name][d] = "OFF";
          currentWorkdays--;
          removed = true;
          if (currentWorkdays <= targetWorkdays) break;
        }
      }
      if (!removed) break;
    }
  }
  
  // ✅ 일별 최소 인원 10명 보장 - 불필요해서 주석처리
  // const MIN_DAILY_STAFF = 10;
  
  for (let d = 1; d <= daysInMonth; d++) {
    // let currentStaff = getDayEmployees_(schedule, employees, d).length;
    const requiredStaff = dailyStaffRequirements[d] || 10;  // matrix_RQ 기준 사용
    let currentStaff = getDayEmployees_(schedule, employees, d).length;
   
    // while (currentStaff < MIN_DAILY_STAFF) {
  
    while (currentStaff < requiredStaff) {
      let added = false;
      
      // 정직원 중 근무일이 목표 이하인 사람부터 우선 투입
      const sortedRegular = regularEmployees.slice().sort((a, b) => {
        const aWorkdays = countWorkdays_(schedule, a);
        const bWorkdays = countWorkdays_(schedule, b);
        const aTarget = workdayRequirements[a] || 22;
        const bTarget = workdayRequirements[b] || 22;
        return (aWorkdays - aTarget) - (bWorkdays - bTarget);
      });
      
      for (let name of sortedRegular) {
        if (schedule[name][d] === "OFF" && 
            employees[name].days[d] && 
            employees[name].days[d].is_available &&
            !employees[name].days[d].is_rq &&
            canWorkOnDay_(schedule, name, d, prevMonthData)) {
          
          const targetWorkdays = workdayRequirements[name] || 22;
          const currentWorkdays = countWorkdays_(schedule, name);
          
          // 최소 인원 보장을 위해 최대 2일까지 초과 허용
          if (currentWorkdays - targetWorkdays < 2) {
            schedule[name][d] = "ON";
            currentStaff++;
            added = true;
            break;
          }
        }
      }
      
      // 정직원으로 충족 불가 시 PARTNER 활용
      if (!added) {
        const partnerEmployees = Object.keys(employees).filter(name => 
          employees[name].employment_type === "PARTNER"
        );
        
        for (let name of partnerEmployees) {
          if (schedule[name][d] === "OFF" && employees[name].days[d] && 
              employees[name].days[d].is_available && !employees[name].days[d].is_rq) {
            schedule[name][d] = "ON";
            currentStaff++;
            added = true;
            break;
          }
        }
      }
      
      if (!added) break; // 더 이상 투입 불가
    }
  }
  
  // ✅ 일별 인원 균등화 (간단한 힐클라이밍 방식)
  const totalStaff = regularEmployees.reduce((sum, name) => 
    sum + countWorkdays_(schedule, name), 0
  );
  const avgStaff = totalStaff / daysInMonth;
  
  for (let iteration = 0; iteration < 3; iteration++) {
    let improved = false;
    
    for (let d = 1; d <= daysInMonth; d++) {
      const dayStaff = getDayEmployees_(schedule, employees, d).length;
      const requiredStaff = dailyStaffRequirements[d] || Math.round(avgStaff);
      
      if (dayStaff > requiredStaff + 2) {
        // 과다 일자에서 과소 일자로 일부 인원 이동 시도
        for (let name of regularEmployees) {
          if (schedule[name][d] === "ON" && 
              employees[name].days[d] && 
              !employees[name].days[d].is_business_trip) {
            
            const currentWorkdays = countWorkdays_(schedule, name);
            const targetWorkdays = workdayRequirements[name] || 22;
            
            // 근무일 수가 목표보다 많은 경우에만 이동 고려
            if (currentWorkdays > targetWorkdays) {
              for (let targetDay = 1; targetDay <= daysInMonth; targetDay++) {
                if (targetDay === d) continue;
                
                const targetStaff = getDayEmployees_(schedule, employees, targetDay).length;
                const targetRequired = dailyStaffRequirements[targetDay] || Math.round(avgStaff);
                
                if (targetStaff < targetRequired - 1 && 
                    schedule[name][targetDay] === "OFF" &&
                    employees[name].days[targetDay] &&
                    employees[name].days[targetDay].is_available &&
                    !employees[name].days[targetDay].is_rq &&
                    canWorkOnDay_(schedule, name, targetDay, prevMonthData)) {
                  
                  // 이동: d -> targetDay
                  schedule[name][d] = "OFF";
                  schedule[name][targetDay] = "ON";
                  improved = true;
                  break;
                }
              }
              if (improved) break;
            }
          }
        }
        if (improved) break;
      }
    }
    
    if (!improved) break;
  }
  
  // 일별 제약조건(운전자/슈퍼바이저/남성 등) 보정 루프
  for (let d = 1; d <= daysInMonth; d++) {
    let attempts = 0;
    while (attempts < 20) {
      const dayEmployees = getDayEmployees_(schedule, employees, d);
      const violations = checkDayConstraints_(dayEmployees, dateInfo[d], employees, dailyStaffRequirements[d] || 10);
      
      if (violations.length === 0) break;
      
      let improved = false;
      
      for (let violation of violations) {
        if (violation.type === "min_staff" && violation.current < violation.required) {
          if (addEmployeeToDayIfPossible_(schedule, employees, d, regularEmployees, prevMonthData, workdayRequirements)) {
            improved = true;
          }
        } else if (violation.type === "driver" && violation.current < violation.required) {
          if (addDriverToDayIfPossible_(schedule, employees, d, regularEmployees, prevMonthData, workdayRequirements)) {
            improved = true;
          }
        } else if (violation.type === "supervisor" && violation.current < violation.required) {
          if (addSupervisorToDayIfPossible_(schedule, employees, d, regularEmployees, prevMonthData, workdayRequirements)) {
            improved = true;
          }
        } else if (violation.type === "male_hh" && violation.current < violation.required) {
          if (addMaleEmployeeToDayIfPossible_(schedule, employees, d, regularEmployees, prevMonthData, workdayRequirements)) {
            improved = true;
          }
        } else if (violation.type === "male_disposal" && violation.current < violation.required) {
          if (addMaleEmployeeToDayIfPossible_(schedule, employees, d, regularEmployees, prevMonthData, workdayRequirements)) {
            improved = true;
          }
        }
      }
      
      if (!improved) break;
      attempts++;
    }
  }
  
  // PARTNER 추가 배정 (아직 부족한 일자에 한해)
  const partnerEmployees = Object.keys(employees).filter(name => 
    employees[name].employment_type === "PARTNER"
  );
  
  for (let d = 1; d <= daysInMonth; d++) {
    const dayEmployees = getDayEmployees_(schedule, employees, d);
    const required = dailyStaffRequirements[d] || 10;
    
    if (dayEmployees.length < required) {
      const needed = required - dayEmployees.length;
      let added = 0;
      
      for (let name of partnerEmployees) {
        if (added >= needed) break;
        if (schedule[name][d] === "OFF" && employees[name].days[d] && 
            employees[name].days[d].is_available && !employees[name].days[d].is_rq) {
          schedule[name][d] = "ON";
          added++;
        }
      }
    }
  }
  
  return schedule;
}

// ✅ 새로운 헬퍼 함수: 근무일 수를 정확히 맞추는 패턴 생성
// - 연속 5일 제한, 가벼운 랜덤성(이름 해시 기반 오프셋) 도입
function generateWorkPattern_(name, availableDays, workdaysNeeded, prevConsecutiveWork, daysInMonth) {
  if (workdaysNeeded <= 0 || availableDays.length === 0) {
    return [];
  }
  
  const pattern = [];
  let consecutiveWork = prevConsecutiveWork;
  let consecutiveOff = 0;
  let longBreakCount = 0;
  let dayIndex = 0;
  
  // 직원별 오프셋 (이름 해시를 통한 의사난수)
  let offset = 0;
  for (let i = 0; i < name.length; i++) {
    offset += name.charCodeAt(i);
  }
  offset = offset % 7;
  
  while (pattern.length < workdaysNeeded && dayIndex < availableDays.length) {
    const day = availableDays[dayIndex];
    let shouldWork = false;
    
    // 연속 근무 5일 제한: 4 이상이면 강제 OFF
    if (consecutiveWork >= 4) {
      shouldWork = false;
      consecutiveWork = 0;
      consecutiveOff++;
    }
    // 2일 연속 휴무 우선 (간헐적 장기 휴식 고려)
    else if (consecutiveOff > 0 && consecutiveOff < 2) {
      // 최근 10일 내 장기휴식 여부는 간단 카운터만 반영
      let recentLongBreak = false;
      for (let i = Math.max(0, pattern.length - 10); i < pattern.length; i++) {
        // 간단화: 여기서는 longBreakCount로 판단
      }
      
      if (longBreakCount === 0 || (day - availableDays[Math.max(0, dayIndex - 10)]) > 10) {
        shouldWork = false;
        consecutiveOff++;
        if (consecutiveOff >= 2) {
          longBreakCount++;
        }
      } else {
        shouldWork = true;
      }
    }
    // 3-4일 연속 근무 시 확률적 휴무
    else if (consecutiveWork >= 3 && consecutiveWork < 5) {
      const remainingDays = availableDays.length - dayIndex;
      const remainingWork = workdaysNeeded - pattern.length;
      const workRatio = remainingWork / Math.max(remainingDays, 1);
      
      let offProbability = 0;
      if (consecutiveWork === 3) {
        offProbability = workRatio < 0.6 ? 0.4 : (workRatio < 0.7 ? 0.2 : 0);
      } else if (consecutiveWork === 4) {
        offProbability = workRatio < 0.7 ? 0.6 : (workRatio < 0.8 ? 0.3 : 0.1);
      }
      
      const decision = ((day + offset) % 10) / 10.0;
      if (decision < offProbability && remainingWork < remainingDays) {
        shouldWork = false;
        consecutiveWork = 0;
        consecutiveOff++;
      } else {
        shouldWork = true;
      }
    }
    // 기본 케이스: 근무
    else {
      shouldWork = true;
    }
    
    if (shouldWork) {
      pattern.push(day);
      consecutiveWork++;
      consecutiveOff = 0;
    } else {
      consecutiveWork = 0;
      consecutiveOff++;
    }
    
    dayIndex++;
  }
  
  // 부족분 보완: 가능한 범위에서 연속 5일 초과를 피하며 추가
  while (pattern.length < workdaysNeeded && dayIndex < availableDays.length) {
    const day = availableDays[dayIndex];
    // 연속 근무 5일 제한만 확인
    let canAdd = true;
    
    // 바로 앞 4일 확인
    let recentWork = 0;
    for (let i = pattern.length - 1; i >= Math.max(0, pattern.length - 4); i--) {
      if (pattern[i] === day - (pattern.length - i)) {
        recentWork++;
      }
    }
    
    if (recentWork < 5) {
      pattern.push(day);
    }
    
    dayIndex++;
  }
  
  return pattern;
}

/**
 * 특정 일자의 '출근'으로 간주되는 직원 목록 반환
 * - 'ON' 또는 'A*'로 시작하는 값(A10/A8 등)을 출근으로 판단
 */
function getDayEmployees_(schedule, employees, day) {
  const result = [];
  for (let name in schedule) {
    const scheduleValue = schedule[name][day];
    // ✅ 'ON' 또는 'A'로 시작하는 값을 모두 출근으로 간주
    if (scheduleValue === "ON" || 
        (scheduleValue && String(scheduleValue).toUpperCase().startsWith("A"))) {
      result.push(name);
    }
  }
  return result;
}

/**
 * 일자 단위 제약검사
 * - 최소 인원, 운전가능자(2명), 슈퍼바이저(1명), HH/폐기물 남성 수
 * - 부족 항목을 violations 배열로 반환
 */
function checkDayConstraints_(dayEmployees, dayInfo, employees, requiredStaff) {
  const violations = [];
  
  if (dayEmployees.length < requiredStaff) {
    violations.push({type: "min_staff", current: dayEmployees.length, required: requiredStaff});
  }
  
  // 운전가능자 (ALL_VEHICLES/SMALL_ONLY)
  const drivers = dayEmployees.filter(name => {
    const dc = employees[name].driving_class;
    return dc === "ALL_VEHICLES" || dc === "SMALL_ONLY";
  });
  if (drivers.length < 2) {
    violations.push({type: "driver", current: drivers.length, required: 2});
  }
  
  // 슈퍼바이저(SV/SSV) 최소 1명
  const supervisors = dayEmployees.filter(name => {
    const et = employees[name].employment_type;
    return et === "SV" || et === "SSV";
  });
  if (supervisors.length < 1) {
    violations.push({type: "supervisor", current: supervisors.length, required: 1});
  }
  
  // HH 청소일: 남성 최소 2명
  if (dayInfo && dayInfo.is_hh_cleaning) {
    const males = dayEmployees.filter(name => employees[name].gender === "M");
    if (males.length < 2) {
      violations.push({type: "male_hh", current: males.length, required: 2});
    }
  }
  
  // 폐기물 처리일: 남성 최소 1명
  if (dayInfo && dayInfo.is_disposal) {
    const males = dayEmployees.filter(name => employees[name].gender === "M");
    if (males.length < 1) {
      violations.push({type: "male_disposal", current: males.length, required: 1});
    }
  }
  
  return violations;
}

/**
 * 일자에 추가로 투입 가능한 인원을 정직원 후보군에서 탐색하여 배치
 * - 목표 근무일 이하 우선
 */
function addEmployeeToDayIfPossible_(schedule, employees, day, candidates, prevMonthData, workdayRequirements) {
  const sortedCandidates = candidates.slice().sort((a, b) => {
    const aWorkdays = countWorkdays_(schedule, a);
    const bWorkdays = countWorkdays_(schedule, b);
    const aTarget = workdayRequirements[a] || 22;
    const bTarget = workdayRequirements[b] || 22;
    return (aWorkdays - aTarget) - (bWorkdays - bTarget);
  });
  
  for (let name of sortedCandidates) {
    if (schedule[name][day] === "OFF" && 
        employees[name].days[day] && 
        employees[name].days[day].is_available &&
        !employees[name].days[day].is_rq &&
        canWorkOnDay_(schedule, name, day, prevMonthData)) {
      const targetWorkdays = workdayRequirements[name] || 22;
      const currentWorkdays = countWorkdays_(schedule, name);
      if (currentWorkdays < targetWorkdays) {
        schedule[name][day] = "ON";
        return true;
      }
    }
  }
  return false;
}

/**
 * 운전가능자 충족을 위해 후보 중 운전자 투입
 */
function addDriverToDayIfPossible_(schedule, employees, day, candidates, prevMonthData, workdayRequirements) {
  const sortedCandidates = candidates.slice().sort((a, b) => {
    const aWorkdays = countWorkdays_(schedule, a);
    const bWorkdays = countWorkdays_(schedule, b);
    const aTarget = workdayRequirements[a] || 22;
    const bTarget = workdayRequirements[b] || 22;
    return (aWorkdays - aTarget) - (bWorkdays - bTarget);
  });
  
  for (let name of sortedCandidates) {
    const dc = employees[name].driving_class;
    if ((dc === "ALL_VEHICLES" || dc === "SMALL_ONLY") &&
        schedule[name][day] === "OFF" && 
        employees[name].days[day] &&
        employees[name].days[day].is_available &&
        !employees[name].days[day].is_rq &&
        canWorkOnDay_(schedule, name, day, prevMonthData)) {
      const targetWorkdays = workdayRequirements[name] || 22;
      const currentWorkdays = countWorkdays_(schedule, name);
      if (currentWorkdays < targetWorkdays) {
        schedule[name][day] = "ON";
        return true;
      }
    }
  }
  return false;
}

/**
 * 슈퍼바이저(SV/SSV) 충족을 위해 후보 중 SV/SSV 투입
 */
function addSupervisorToDayIfPossible_(schedule, employees, day, candidates, prevMonthData, workdayRequirements) {
  const sortedCandidates = candidates.slice().sort((a, b) => {
    const aWorkdays = countWorkdays_(schedule, a);
    const bWorkdays = countWorkdays_(schedule, b);
    const aTarget = workdayRequirements[a] || 22;
    const bTarget = workdayRequirements[b] || 22;
    return (aWorkdays - aTarget) - (bWorkdays - bTarget);
  });
  
  for (let name of sortedCandidates) {
    const et = employees[name].employment_type;
    if ((et === "SV" || et === "SSV") &&
        schedule[name][day] === "OFF" && 
        employees[name].days[day] &&
        employees[name].days[day].is_available &&
        !employees[name].days[day].is_rq &&
        canWorkOnDay_(schedule, name, day, prevMonthData)) {
      const targetWorkdays = workdayRequirements[name] || 22;
      const currentWorkdays = countWorkdays_(schedule, name);
      if (currentWorkdays < targetWorkdays) {
        schedule[name][day] = "ON";
        return true;
      }
    }
  }
  return false;
}

/**
 * 남성 인력 충족(청소/폐기물) 목적의 투입
 */
function addMaleEmployeeToDayIfPossible_(schedule, employees, day, candidates, prevMonthData, workdayRequirements) {
  const sortedCandidates = candidates.slice().sort((a, b) => {
    const aWorkdays = countWorkdays_(schedule, a);
    const bWorkdays = countWorkdays_(schedule, b);
    const aTarget = workdayRequirements[a] || 22;
    const bTarget = workdayRequirements[b] || 22;
    return (aWorkdays - aTarget) - (bWorkdays - bTarget);
  });
  
  for (let name of sortedCandidates) {
    if (employees[name].gender === "M" &&
        schedule[name][day] === "OFF" && 
        employees[name].days[day] &&
        employees[name].days[day].is_available &&
        !employees[name].days[day].is_rq &&
        canWorkOnDay_(schedule, name, day, prevMonthData)) {
      const targetWorkdays = workdayRequirements[name] || 22;
      const currentWorkdays = countWorkdays_(schedule, name);
      if (currentWorkdays < targetWorkdays) {
        schedule[name][day] = "ON";
        return true;
      }
    }
  }
  return false;
}

/**
 * 특정 직원이 해당 일자에 근무해도 되는지(연속 5일 초과 방지) 판단
 * - 월초의 경우 전월 연속 근무를 합산하여 판단
 */
function canWorkOnDay_(schedule, name, day, prevMonthData) {
  let consecutiveWork = 0;
  
  // 당월 기준 직전 연속 근무일 계산
  for (let d = day - 1; d >= 1; d--) {
    if (schedule[name][d] === "ON") {
      consecutiveWork++;
    } else {
      break;
    }
  }
  
  // 월초 케이스: 전월 말일부터 이어지는 연속 근무 합산
  if (day - 1 - consecutiveWork < 1 && prevMonthData[name]) {
    for (let i = prevMonthData[name].length - 1; i >= 0; i--) {
      if (prevMonthData[name][i].schedule === "ON") {
        consecutiveWork++;
      } else {
        break;
      }
    }
  }
  
  return consecutiveWork < 5;
}

/**
 * 스케줄에서 'ON' 카운트(근무일 수) 계산
 */
function countWorkdays_(schedule, name) {
  let count = 0;
  for (let day in schedule[name]) {
    if (schedule[name][day] === "ON") count++;
  }
  return count;
}

/* ================================================================
   ✅ 스케줄 검증 기능
   - 작성된 스케줄이 제약조건을 만족하는지 검증하고 보고서 생성
================================================================ */


function scheduleValidateMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const dbSheet = ss.getSheetByName("DB_schedule");
  
  if (!dbSheet) {
    ui.alert("DB_schedule 시트를 찾을 수 없습니다.");
    return;
  }
  
  const data = dbSheet.getDataRange().getValues();
  const headers = data[0];
  const monthCol = headers.indexOf("month");
  const scheduleCol = headers.indexOf("schedule");
  
  if (monthCol === -1 || scheduleCol === -1) {
    ui.alert("DB_schedule에 month 또는 schedule 열이 없습니다.");
    return;
  }
  
  // 스케줄이 채워진 월만 대상으로 검증
  const filledMonths = new Set();
  for (let i = 1; i < data.length; i++) {
    const month = data[i][monthCol];
    const schedule = data[i][scheduleCol];
    if (month && schedule && String(schedule).trim() !== "") {
      filledMonths.add(Number(month));
    }
  }
  
  if (filledMonths.size === 0) {
    ui.alert("스케줄이 작성된 월이 없습니다.");
    return;
  }
  
  // 검증 대상 월 선택
  const monthList = Array.from(filledMonths).sort((a, b) => a - b);
  let msg = "검증할 월을 선택하세요.\n\n";
  for (let i = 0; i < monthList.length; i++) {
    msg += (i + 1) + ". " + monthList[i] + "월\n";
  }
  msg += "\n번호를 입력하세요:";
  
  const res = ui.prompt("스케줄 검증", msg, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  
  const idx = Number(String(res.getResponseText()).trim());
  if (!isFinite(idx) || idx < 1 || idx > monthList.length) {
    ui.alert("올바른 번호가 아닙니다.");
    return;
  }
  
  const targetMonth = monthList[idx - 1];
  
  try {
    const result = validateScheduleForMonth_(targetMonth);
    
    if (result.violations.length === 0) {
      ui.alert("✅ " + targetMonth + "월 스케줄 검증 완료\n\n모든 제약조건을 만족합니다!");
    } else {
      // ✅ 시트 생성
      createViolationReportSheet_(result);
      ui.alert("⚠️ " + targetMonth + "월 스케줄 위반사항\n\n총 " + result.violations.length + "건의 위반사항이 발견되었습니다.\n\n위반 보고서 시트가 생성되었습니다.");
    }
  } catch (error) {
    ui.alert("❌ 오류 발생: " + error.toString());
  }
}

/**
 * 대상 월 스케줄을 읽어 일별/개인별 제약사항 검증
 * - roster: 'A*' 또는 'ON'을 출근으로 간주
 * - violations: [{type, date?, name?, description, severity}]
 */
function validateScheduleForMonth_(targetMonth) {
  const ss = SpreadsheetApp.getActive();
  const dbSheet = ss.getSheetByName("DB_schedule");
  const data = dbSheet.getDataRange().getValues();
  const headers = data[0];
  
  // 필수 컬럼 인덱스
  const colIdx = {};
  const requiredCols = ["year","month","date","day","name","is_rq","employment_type_code","driving_class","gender_code","is_disposal_day","is_hh_cleaning_day","is_available_day","is_business_trip","schedule","roster"];
  
  for (let col of requiredCols) {
    colIdx[col] = headers.indexOf(col);
    if (colIdx[col] === -1) throw new Error("필수 열을 찾을 수 없습니다: " + col);
  }
  
  // 대상 월 필터링
  const monthData = [];
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][colIdx.month]) === targetMonth) {
      monthData.push(data[i]);
    }
  }
  
  if (monthData.length === 0) throw new Error(targetMonth + "월 데이터가 없습니다.");
  
  const year = monthData[0][colIdx.year];
  const employees = {};
  const dateInfo = {};
  const schedule = {};
  
  for (let row of monthData) {
    const name = row[colIdx.name];
    const date = Number(row[colIdx.date]);
    
    if (!employees[name]) {
      employees[name] = {
        name: name,
        employment_type: row[colIdx.employment_type_code],
        driving_class: row[colIdx.driving_class],
        gender: row[colIdx.gender_code],
        days: {}
      };
      schedule[name] = {};
    }
    
    employees[name].days[date] = {
      is_rq: row[colIdx.is_rq] === true || row[colIdx.is_rq] === "TRUE",
      is_available: row[colIdx.is_available_day] === true || row[colIdx.is_available_day] === "TRUE",
      is_business_trip: row[colIdx.is_business_trip] === true || row[colIdx.is_business_trip] === "TRUE",
      day_of_week: row[colIdx.day]
    };
    
    // roster가 'A'로 시작(A10, A8 등)하거나 roster가 'ON'이면 출근으로 간주
    const rosterVal   = colIdx.roster   !== -1 ? String(row[colIdx.roster]).trim().toUpperCase()   : "";
    const scheduleVal = colIdx.schedule !== -1 ? String(row[colIdx.schedule]).trim().toUpperCase() : "";

    // ✅ roster가 'A'로 시작하거나 'ON'이면 출근
    const isOn = (rosterVal.startsWith("A") || rosterVal === "ON");
    schedule[name][date] = isOn ? "ON" : "OFF";
    
    if (!dateInfo[date]) {
      dateInfo[date] = {
        day_of_week: row[colIdx.day],
        is_disposal: row[colIdx.is_disposal_day] === true || row[colIdx.is_disposal_day] === "TRUE",
        is_hh_cleaning: row[colIdx.is_hh_cleaning_day] === true || row[colIdx.is_hh_cleaning_day] === "TRUE"
      };
    }
  }
  
  const workdayRequirements = getWorkdayRequirements_(year, targetMonth);
  const dailyStaffRequirements = getDailyStaffRequirements_(year, targetMonth);
  
  const violations = [];
  const daysInMonth = new Date(year, targetMonth, 0).getDate();
  
  // 일별 제약조건 검증
  for (let d = 1; d <= daysInMonth; d++) {
    const dayEmployees = getDayEmployees_(schedule, employees, d);
    const requiredStaff = dailyStaffRequirements[d] || 10;
    const dayViolations = checkDayConstraints_(dayEmployees, dateInfo[d], employees, requiredStaff);
    
    for (let v of dayViolations) {
      violations.push({
        type: "daily",
        date: d,
        description: formatViolation_(v),
        severity: v.type === "supervisor" || v.type === "driver" ? "high" : "medium"
      });
    }
  }
  
  // 개인별 제약조건 검증
  for (let name in employees) {
    const empType = employees[name].employment_type;
    
    // ✅ SP와 SSV는 제외
    if (empType === "SP" || empType === "SSV"|| empType === "PARTNER") {
      continue;
    }
    
    // RQ 신청일 출근 금지 (PARTNER 제외 이미 continue)
    if (empType !== "PARTNER") {
      for (let d = 1; d <= daysInMonth; d++) {
        if (employees[name].days[d] && employees[name].days[d].is_rq && schedule[name][d] === "ON") {
          violations.push({
            type: "rq",
            name: name,
            date: d,
            description: "RQ 신청일인데 출근으로 배정됨",
            severity: "high"
          });
        }
      }
    }
    
    // 연속 근무 5일 초과 검증
    let consecutiveWork = 0;
    for (let d = 1; d <= daysInMonth; d++) {
      if (schedule[name][d] === "ON") {
        consecutiveWork++;
        if (consecutiveWork > 5) {
          violations.push({
            type: "consecutive",
            name: name,
            date: d,
            description: "연속 근무 5일 초과 (" + consecutiveWork + "일)",
            severity: "high"
          });
        }
      } else {
        consecutiveWork = 0;
      }
    }
    
    // 근무일 수 검증 (PARTNER 제외)
    if (empType !== "PARTNER") {
      const targetWorkdays = workdayRequirements[name];
      if (targetWorkdays) {
        let actualWorkdays = 0;
        for (let d = 1; d <= daysInMonth; d++) {
          if (schedule[name][d] === "ON") actualWorkdays++;
        }
        if (actualWorkdays !== targetWorkdays) {
          violations.push({
            type: "workdays",
            name: name,
            date: null,
            description: "근무일 수 불일치 (목표: " + targetWorkdays + "일, 실제: " + actualWorkdays + "일)",
            severity: actualWorkdays < targetWorkdays ? "high" : "medium"
          });
        }
      }
    }
  }
  
  return {
    violations: violations,
    year: year,
    month: targetMonth
  };
}

// ✅ 새로운 함수: 위반사항 보고서 시트 생성
// - 심각도/유형/직원/날짜/내용 테이블, 유형별 통계 포함
function createViolationReportSheet_(result) {
  const ss = SpreadsheetApp.getActive();
  const year = result.year;
  const month = result.month;
  
  // 시트 이름 생성
  const baseName = String(year).slice(-2) + " " + MON_ABBR[month - 1] + " 위반사항";
  const sheetName = getUniqueSheetName_(baseName);
  
  // 기존 시트 삭제 (같은 월의 위반사항 시트가 있으면)
  const existingSheet = ss.getSheetByName(sheetName);
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }
  
  const sheet = ss.insertSheet(sheetName);
  
  // 헤더 설정
  sheet.getRange("A1").setValue("스케줄 위반사항 보고서");
  sheet.getRange("A1").setFontSize(14).setFontWeight("bold");
  
  sheet.getRange("A2").setValue("년도: " + year);
  sheet.getRange("B2").setValue("월: " + month);
  sheet.getRange("C2").setValue("검증일시: " + new Date().toLocaleString("ko-KR"));
  
  sheet.getRange("A3").setValue("총 위반 건수: " + result.violations.length);
  sheet.getRange("A3").setFontWeight("bold").setBackground("#fff2cc");
  
  // 테이블 헤더
  const headerRow = 5;
  const headers = ["심각도", "유형", "직원명", "날짜", "위반 내용"];
  sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(headerRow, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#4a86e8")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  
  // 위반사항 데이터 정렬 (심각도 순)
  const sortedViolations = result.violations.sort((a, b) => {
    const severityOrder = {"high": 0, "medium": 1, "low": 2};
    return severityOrder[a.severity] - severityOrder[b.severity];
  });
  
  // 데이터 입력
  const dataStartRow = headerRow + 1;
  const tableData = [];
  
  for (let v of sortedViolations) {
    const severityText = v.severity === "high" ? "🔴 높음" : v.severity === "medium" ? "🟡 중간" : "🟢 낮음";
    
    let typeText = "";
    if (v.type === "daily") typeText = "일별 제약";
    else if (v.type === "rq") typeText = "RQ 위반";
    else if (v.type === "consecutive") typeText = "연속 근무";
    else if (v.type === "workdays") typeText = "근무일 수";
    else typeText = v.type;
    
    const nameText = v.name || "-";
    const dateText = v.date ? month + "월 " + v.date + "일" : "-";
    
    tableData.push([severityText, typeText, nameText, dateText, v.description]);
  }
  
  if (tableData.length > 0) {
    sheet.getRange(dataStartRow, 1, tableData.length, headers.length).setValues(tableData);
    
    // 테두리 설정
    sheet.getRange(headerRow, 1, tableData.length + 1, headers.length)
      .setBorder(true, true, true, true, true, true);
    
    // 심각도별 색상 적용
    for (let i = 0; i < tableData.length; i++) {
      const rowNum = dataStartRow + i;
      const severity = sortedViolations[i].severity;
      
      let bgColor = "#ffffff";
      if (severity === "high") bgColor = "#f4cccc";
      else if (severity === "medium") bgColor = "#fce5cd";
      
      sheet.getRange(rowNum, 1, 1, headers.length).setBackground(bgColor);
    }
  }
  
  // 열 너비 조정
  sheet.setColumnWidth(1, 80);  // 심각도
  sheet.setColumnWidth(2, 100); // 유형
  sheet.setColumnWidth(3, 100); // 직원명
  sheet.setColumnWidth(4, 100); // 날짜
  sheet.setColumnWidth(5, 400); // 위반 내용
  
  // 행 높이 조정
  sheet.setRowHeight(1, 30);
  sheet.setRowHeight(headerRow, 30);
  
  // 고정
  sheet.setFrozenRows(headerRow);
  
  // 요약 통계 추가
  const summaryStartRow = dataStartRow + tableData.length + 2;
  
  sheet.getRange(summaryStartRow, 1).setValue("위반 유형별 통계");
  sheet.getRange(summaryStartRow, 1).setFontWeight("bold").setFontSize(12);
  
  const typeCount = {};
  for (let v of sortedViolations) {
    let typeText = "";
    if (v.type === "daily") typeText = "일별 제약";
    else if (v.type === "rq") typeText = "RQ 위반";
    else if (v.type === "consecutive") typeText = "연속 근무";
    else if (v.type === "workdays") typeText = "근무일 수";
    else typeText = v.type;
    
    typeCount[typeText] = (typeCount[typeText] || 0) + 1;
  }
  
  let statRow = summaryStartRow + 1;
  for (let type in typeCount) {
    sheet.getRange(statRow, 1).setValue(type);
    sheet.getRange(statRow, 2).setValue(typeCount[type] + "건");
    statRow++;
  }
  
  // 시트를 맨 앞으로 이동
  ss.moveActiveSheet(1);
}

/**
 * 일별 제약 위반 객체를 사람이 읽기 쉬운 설명으로 포맷
 */
function formatViolation_(v) {
  if (v.type === "min_staff") {
    return "최소 인원 부족 (현재: " + v.current + "명, 필요: " + v.required + "명)";
  } else if (v.type === "driver") {
    return "운전가능자 부족 (현재: " + v.current + "명, 필요: " + v.required + "명)";
  } else if (v.type === "supervisor") {
    return "슈퍼바이저 부족 (현재: " + v.current + "명, 필요: " + v.required + "명)";
  } else if (v.type === "male_hh") {
    return "청소일 남성 직원 부족 (현재: " + v.current + "명, 필요: " + v.required + "명)";
  } else if (v.type === "male_disposal") {
    return "처리일 남성 직원 부족 (현재: " + v.current + "명, 필요: " + v.required + "명)";
  }
  return "알 수 없는 위반";
}

/* ================================================================
   ✅ 스케줄 시각화 시트 생성
   - DB_schedule의 스케줄을 캘린더 매트릭스 형태로 시각화
   - A10/BT/H/W/"" 등 표시 및 요약 통계 포함
================================================================ */

function scheduleVisualizeMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const dbSheet = ss.getSheetByName("DB_schedule");
  
  if (!dbSheet) {
    ui.alert("DB_schedule 시트를 찾을 수 없습니다.");
    return;
  }
  
  const data = dbSheet.getDataRange().getValues();
  const headers = data[0];
  const monthCol = headers.indexOf("month");
  const scheduleCol = headers.indexOf("schedule");
  
  if (monthCol === -1 || scheduleCol === -1) {
    ui.alert("DB_schedule에 month 또는 schedule 열이 없습니다.");
    return;
  }
  
  // 스케줄이 작성된 월만 나열
  const filledMonths = new Set();
  for (let i = 1; i < data.length; i++) {
    const month = data[i][monthCol];
    const schedule = data[i][scheduleCol];
    if (month && schedule && String(schedule).trim() !== "") {
      filledMonths.add(Number(month));
    }
  }
  
  if (filledMonths.size === 0) {
    ui.alert("스케줄이 작성된 월이 없습니다.");
    return;
  }
  
  // 대상 월 선택
  const monthList = Array.from(filledMonths).sort((a, b) => a - b);
  let msg = "시각화할 월을 선택하세요.\n\n";
  for (let i = 0; i < monthList.length; i++) {
    msg += (i + 1) + ". " + monthList[i] + "월\n";
  }
  msg += "\n번호를 입력하세요:";
  
  const res = ui.prompt("스케줄 시각화 시트 생성", msg, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  
  const idx = Number(String(res.getResponseText()).trim());
  if (!isFinite(idx) || idx < 1 || idx > monthList.length) {
    ui.alert("올바른 번호가 아닙니다.");
    return;
  }
  
  const targetMonth = monthList[idx - 1];
  
  try {
    createScheduleVisualizationSheet_(targetMonth);
    ui.alert("✅ " + targetMonth + "월 스케줄 시각화 시트가 생성되었습니다.");
  } catch (error) {
    ui.alert("❌ 오류 발생: " + error.toString());
  }
}

/**
 * 시각화 시트 생성
 * - roster 우선: roster가 있으면 그대로 표시, 없으면 schedule=ON은 A10로 대체
 * - 색상: BT 회색, H/W/OFF 붉은색 계열, 그 외 흰색
 * - 하단 요약행: 출근 인원/운전가능자/슈퍼바이저/HH/폐기물
 */
function createScheduleVisualizationSheet_(targetMonth) {
  const ss = SpreadsheetApp.getActive();
  const dbSheet = ss.getSheetByName("DB_schedule");
  const empSheet = ss.getSheetByName("matrix_employees");
  
  if (!empSheet) throw new Error("matrix_employees 시트를 찾을 수 없습니다.");
  
  const data = dbSheet.getDataRange().getValues();
  const headers = data[0];
  
  // 필수 컬럼 매핑
  const colIdx = {};
  const requiredCols = ["year","month","date","day","name","employment_type_code","schedule","roster","is_hh_cleaning_day","is_disposal_day"];
  
  for (let col of requiredCols) {
    colIdx[col] = headers.indexOf(col);
    if (colIdx[col] === -1) throw new Error("필수 열을 찾을 수 없습니다: " + col);
  }
  
  // 대상 월 필터링
  const monthData = [];
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][colIdx.month]) === targetMonth) {
      monthData.push(data[i]);
    }
  }
  
  if (monthData.length === 0) throw new Error(targetMonth + "월 데이터가 없습니다.");
  
  const year = monthData[0][colIdx.year];
  const employeeSchedule = {};
  const employeeType = {};
  const dates = new Set();
  
  // 직원/일자별 셀 표시값 구성
  for (let row of monthData) {
    const name = row[colIdx.name];
    const date = Number(row[colIdx.date]);
    // const schedule = row[colIdx.schedule];
    const rosterRaw   = (colIdx.roster   !== -1) ? String(row[colIdx.roster]).trim()   : "";
    const scheduleRaw = (colIdx.schedule !== -1) ? String(row[colIdx.schedule]).trim() : "";

    // ✅ roster가 있으면 그대로, 없으면 schedule=ON을 A10으로 폴백, OFF면 공란
    const cellValue = rosterRaw || (scheduleRaw === "ON" ? "A10" : "");
    const empType = row[colIdx.employment_type_code];

  
    if (!employeeSchedule[name]) {
      employeeSchedule[name] = {};
      employeeType[name] = empType;
    }
    
    // employeeSchedule[name][date] = schedule === "ON" ? "ON" : "OFF";
    employeeSchedule[name][date] = cellValue; // ✅ 이제 셀에는 A10/BT/H/W/"" 가 들어감
    dates.add(date);
  }
  
  const sortedDates = Array.from(dates).sort((a, b) => a - b);
  const daysInMonth = sortedDates[sortedDates.length - 1];
  
   // 직원 기본 정보 (운전 가능/고용유형) 로드
  const empData = empSheet.getDataRange().getValues();
  const empHeaders = empData[0];
  const empNameCol = empHeaders.indexOf("name");
  const empTypeCol = empHeaders.indexOf("employment_type_code");
  const empDrivingCol = empHeaders.indexOf("driving_class");
  
  if (empNameCol === -1) throw new Error("matrix_employees에 name 열이 없습니다.");
  
  const employeeNames = [];
  const employeeInfo = {};
  
  for (let i = 1; i < empData.length; i++) {
    const name = empData[i][empNameCol];
    const type = empData[i][empTypeCol];
    const driving = empData[i][empDrivingCol];
    
    if (name && employeeSchedule[name] && !employeeInfo[name]) {  // ✅ 중복 방지
      employeeNames.push(name);
      employeeInfo[name] = {
        type: type,
        driving: driving
      };
    }
  }
  
  // 시트 생성 (YY MON Schedule)
  const baseName = String(year).slice(-2) + " " + MON_ABBR[targetMonth - 1] + " Schedule";
  const sheetName = getUniqueSheetName_(baseName);
  const sheet = ss.insertSheet(sheetName);
  
  // 상단 메타 (년도/월/일자/요일)
  sheet.getRange("A1").setValue("년도");
  sheet.getRange("B1").setValue(year);
  sheet.getRange("A2").setValue("월");
  sheet.getRange("B2").setValue(targetMonth);
  sheet.getRange("A3").setValue("일자");
  sheet.getRange("A4").setValue("요일");
  
  const NAME_START_ROW = 5;
  const DAY_ROW = 3;
  const YOIL_ROW = 4;
  const FIRST_DAY_COL = 3;
  
  // 이름 렌더링
  const nameValues = employeeNames.map(name => [name]);
  if (nameValues.length > 0) {
    sheet.getRange(NAME_START_ROW, 2, nameValues.length, 1).setValues(nameValues);
  }
  
  // 달력 헤더 (요일/색상)
  const weekdays = ["일","월","화","수","목","금","토"];
  const dateRow = [];
  const yoilRow = [];
  const colorRow = [];
  
  for (let d = 1; d <= daysInMonth; d++) {
    const dt = new Date(year, targetMonth - 1, d);
    const dow = dt.getDay();
    const color = dow === 0 ? "#d93025" : dow === 6 ? "#1a73e8" : "#000000";
    dateRow.push(d);
    yoilRow.push(weekdays[dow]);
    colorRow.push(color);
  }
  
  sheet.getRange(DAY_ROW, FIRST_DAY_COL, 1, daysInMonth).setValues([dateRow]);
  sheet.getRange(YOIL_ROW, FIRST_DAY_COL, 1, daysInMonth).setValues([yoilRow]);
  sheet.getRange(DAY_ROW, FIRST_DAY_COL, 1, daysInMonth)
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setFontColors([colorRow]);
  sheet.getRange(YOIL_ROW, FIRST_DAY_COL, 1, daysInMonth)
    .setHorizontalAlignment("center")
    .setFontColors([colorRow]);
  
  // 본문 스케줄 채우기
  const scheduleValues = [];
  for (let name of employeeNames) {
    const row = [];
    for (let d = 1; d <= daysInMonth; d++) {
      row.push(employeeSchedule[name][d] || "OFF");
    }
    scheduleValues.push(row);
  }
  
  if (scheduleValues.length > 0 && daysInMonth > 0) {
    const scheduleRange = sheet.getRange(NAME_START_ROW, FIRST_DAY_COL, scheduleValues.length, daysInMonth);
    scheduleRange.setValues(scheduleValues);
    scheduleRange.setHorizontalAlignment("center");
    
    // const colors = scheduleValues.map(row => 
    //   row.map(val => val === "ON" ? "#d9ead3" : "#f3f3f3")
    // );

    // 색상 지정
    const colors = scheduleValues.map(row =>
   row.map(val => {
     if (val === "BT") return "#d9d9d9";   // 회색
     if (val === "H") return "#FFAEAE";   // 붉은색
     if (val === "W"|| val === "OFF") return "#f4cccc"; // 옅은 붉은색
     return "#ffffff"; // A10 또는 공란은 기본 흰색
   })
 );
    scheduleRange.setBackgrounds(colors);
  }

  // 근무일 합계 열 추가 (A10/BT/ON 카운트)
  const workdayCol = FIRST_DAY_COL + daysInMonth;
  sheet.getRange(YOIL_ROW, workdayCol).setValue("근무일수").setHorizontalAlignment("center").setFontWeight("bold");

  for (let i = 0; i < employeeNames.length; i++) {
    const rowNum = NAME_START_ROW + i;
    const colLetter = columnToLetter_(FIRST_DAY_COL);
    const endColLetter = columnToLetter_(FIRST_DAY_COL + daysInMonth - 1);
      sheet.getRange(rowNum, workdayCol)
      .setFormula(
        '=COUNTIF(' + colLetter + rowNum + ':' + endColLetter + rowNum + ',"A*")' +
        '+COUNTIF(' + colLetter + rowNum + ':' + endColLetter + rowNum + ',"BT")' +
        '+COUNTIF(' + colLetter + rowNum + ':' + endColLetter + rowNum + ',"ON")'
      )
      .setHorizontalAlignment("center")
      .setBackground("#fff2cc");
  }
  
  // 근무일 합계 열 추가 (A10/BT/ON 카운트)
  const lastNameRow = NAME_START_ROW + Math.max(employeeNames.length, 1) - 1;
  sheet.getRange(YOIL_ROW, 2, Math.max(employeeNames.length, 1) + 1, daysInMonth + 1)
    .setBorder(true, true, true, true, true, true);
  
  sheet.setFrozenRows(YOIL_ROW);
  sheet.setFrozenColumns(2);
  if (daysInMonth > 0) sheet.setColumnWidths(FIRST_DAY_COL, daysInMonth, 45);
  sheet.setColumnWidth(2, 100);
  if (employeeNames.length > 0) sheet.setRowHeights(NAME_START_ROW, employeeNames.length, 24);
  
  // 하단 요약행: 출근/운전/슈퍼바이저 + HH/폐기물 표시
  const summaryRow = lastNameRow + 1;
  const onDutyRow = summaryRow;
  const driverRow = summaryRow + 1;
  const supervisorRow = summaryRow + 2;
  
  sheet.getRange(onDutyRow, 2).setValue("출근 인원");
  sheet.getRange(driverRow, 2).setValue("운전가능자");
  sheet.getRange(supervisorRow, 2).setValue("슈퍼바이저");
  
  // HH 청소일/ 폐기물 처리일 행
  const hhCleaningRow = supervisorRow + 1;
  const disposalRow = supervisorRow + 2;

  sheet.getRange(hhCleaningRow, 2).setValue("HH").setBackground("#e0f2f7");
  sheet.getRange(disposalRow, 2).setValue("폐기물").setBackground("#fce5cd");

  // 각 열(일자)별 요약 수식/마커
  for (let col = 0; col < daysInMonth; col++) {
    const colLetter = columnToLetter_(FIRST_DAY_COL + col);
    const startRow = NAME_START_ROW;
    const endRow = lastNameRow;

    const dateObj = new Date(year, targetMonth - 1, col + 1);
    const currentDate = col + 1;
    
    // 출근 인원: A* 또는 ON
    sheet.getRange(onDutyRow, FIRST_DAY_COL + col)
        .setFormula(
          '=COUNTIF(' + colLetter + startRow + ':' + colLetter + endRow + ',"A*")' +
          '+COUNTIF(' + colLetter + startRow + ':' + colLetter + endRow + ',"ON")'
        );
    
    // 운전가능자 수: A*로 출근하면서 운전가능한 사람 카운트 (SUMPRODUCT)
    let driverFormula = '=SUMPRODUCT((';
    for (let i = 0; i < employeeNames.length; i++) {
      const name = employeeNames[i];
      const rowNum = NAME_START_ROW + i;
      const isDriver = employeeInfo[name].driving === "ALL_VEHICLES" || employeeInfo[name].driving === "SMALL_ONLY";
      if (i > 0) driverFormula += '+';
      driverFormula += '((LEFT(' + colLetter + rowNum + ',1)="A")+(' + colLetter + rowNum + '="ON"))*' + (isDriver ? '1' : '0');
    }
    driverFormula += '))';
    sheet.getRange(driverRow, FIRST_DAY_COL + col).setFormula(driverFormula);
    
    // 슈퍼바이저 수: A10로 출근하면서 SV/SSV 인원
    let svFormula = '=SUMPRODUCT((';
    for (let i = 0; i < employeeNames.length; i++) {
      const name = employeeNames[i];
      const rowNum = NAME_START_ROW + i;
      const isSV = employeeInfo[name].type === "SV" || employeeInfo[name].type === "SSV";
      if (i > 0) svFormula += '+';
      svFormula += '((LEFT(' + colLetter + rowNum + ',1)="A")+(' + colLetter + rowNum + '="ON"))*' + (isSV ? '1' : '0');
    }
    svFormula += '))';
    sheet.getRange(supervisorRow, FIRST_DAY_COL + col).setFormula(svFormula);

  // monthData에서 해당 날짜의 is_hh_cleaning_day와 is_disposal_day 찾기
  let isHHCleaning = false;
  let isDisposal = false;
  
  for (let row of monthData) {
    if (Number(row[colIdx.date]) === currentDate) {
      isHHCleaning = row[colIdx.is_hh_cleaning_day] === true || row[colIdx.is_hh_cleaning_day] === "TRUE";
      isDisposal = row[colIdx.is_disposal_day] === true || row[colIdx.is_disposal_day] === "TRUE";
      break;
    }
  }
  
  // 마커(●) 표시
  sheet.getRange(hhCleaningRow, FIRST_DAY_COL + col)
    .setValue(isHHCleaning ? "●" : "")
    .setHorizontalAlignment("center")
    .setBackground("#e0f2f7");
    
  sheet.getRange(disposalRow, FIRST_DAY_COL + col)
    .setValue(isDisposal ? "●" : "")
    .setHorizontalAlignment("center")
    .setBackground("#fce5cd");
  }
  
  sheet.getRange(onDutyRow, 2, 3, daysInMonth + 1)
    .setFontWeight("bold")
    .setBackground("#fff2cc");
}

/**
 * (별도 유틸) 한국 공휴일 목록 반환 (테스트/참조용)
 * - 반환: [[Date], [Date], ...] 형태
 */
function getKoreanHolidays(year, month) {
  const calendarId = 'ko.south_korea#holiday@group.v.calendar.google.com';
  const calendar = CalendarApp.getCalendarById(calendarId);
  const startDate = new Date(year, month - 1, 1);
  const endDate = new Date(year, month, 0);
  const holidays = calendar.getEvents(startDate, endDate);
  return holidays.map(event => [event.getStartTime()]);
}

// 테스트용 함수
function testGetWorkdayRequirements() {
  const result = getWorkdayRequirements_(2025, 10);  // 연도, 월 입력
  Logger.log(result);
}
