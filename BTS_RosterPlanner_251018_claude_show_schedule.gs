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
  const y = ui.prompt("연도 입력", "예: 2025", ui.ButtonSet.OK_CANCEL);
  if (y.getSelectedButton() !== ui.Button.OK) return;
  const year = Number(String(y.getResponseText()).trim());
  if (!isFinite(year) || year < 1900 || year > 2100) {
    ui.alert("연도 입력이 올바르지 않습니다. (예: 2025)");
    return;
  }

  const m = ui.prompt("월 입력", "1 ~ 12 중 하나 (예: 10)", ui.ButtonSet.OK_CANCEL);
  if (m.getSelectedButton() !== ui.Button.OK) return;
  const month = Number(String(m.getResponseText()).trim());
  if (!isFinite(month) || month < 1 || month > 12) {
    ui.alert("월 입력이 올바르지 않습니다. (1~12)");
    return;
  }

  const l = ui.prompt("개인 최대 RQ 신청 가능 수 입력", "1 ~ 13 (예: 5)", ui.ButtonSet.OK_CANCEL);
  if (l.getSelectedButton() !== ui.Button.OK) return;
  const perPersonLimit = Number(String(l.getResponseText()).trim());
  if (!isFinite(perPersonLimit) || perPersonLimit < 1 || perPersonLimit > 13) {
    ui.alert("개인 최대 수는 1~13 사이의 정수여야 합니다.");
    return;
  }

  const baseName = String(year).slice(-2) + " " + MON_ABBR[month - 1] + " RQ";
  const sheetName = getUniqueSheetName_(baseName);
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.insertSheet(sheetName);
  sheet.getRange("Z1").setValue(perPersonLimit);
  fillRQContent_(sheet, year, month);
  ui.alert('완료! 새 시트 "' + sheetName + '"를 생성했습니다.\n(개인 최대 신청 수: ' + perPersonLimit + ")");
}

function getUniqueSheetName_(base) {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(base)) return base;
  let i = 1;
  while (ss.getSheetByName(base + " (" + i + ")")) i++;
  return base + " (" + i + ")";
}

function fillRQContent_(sheet, year, month) {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!src) throw new Error('시트 "' + SOURCE_SHEET_NAME + '" 를 찾을 수 없습니다.');

  sheet.getRange("A1").setValue("년도");
  sheet.getRange("B1").setValue(year);
  sheet.getRange("A2").setValue("월");
  sheet.getRange("B2").setValue(month);
  sheet.getRange("A3").setValue("일자");
  sheet.getRange("A4").setValue("요일");

  const lastRow = src.getLastRow();
  const raw = lastRow > 1 ? src.getRange(2, 1, lastRow - 1, 3).getValues() : [];
  const rows = [];
  for (let i = 0; i < raw.length; i++) {
    const id = raw[i][0], name = raw[i][1], type = raw[i][2];
    if (id && name && String(type).toUpperCase() !== "PARTNER") rows.push([id, name]);
  }
  rows.sort((a, b) => a[0] > b[0] ? 1 : a[0] < b[0] ? -1 : 0);
  const names = rows.map(r => [r[1]]);

  const NAME_START_ROW = 5, DAY_ROW = 3, YOIL_ROW = 4, FIRST_DAY_COL = 3;
  const nameCount = names.length;
  if (nameCount) sheet.getRange(NAME_START_ROW, 2, nameCount, 1).setValues(names);

  const lastDay = new Date(year, month, 0).getDate();
  const weekdays = ["일","월","화","수","목","금","토"];
  const holidaySet = getKoreanHolidaySet_(year, month);

  const dateRow = [], yoilRow = [], colorRow = [];
  for (let d = 1; d <= lastDay; d++) {
    const dt = new Date(year, month - 1, d);
    const dow = dt.getDay();
    const isHoliday = holidaySet.has(d);
    const color = dow === 0 || isHoliday ? "#d93025" : dow === 6 ? "#1a73e8" : "#000000";
    dateRow.push(d);
    yoilRow.push(weekdays[dow]);
    colorRow.push(color);
  }

  if (lastDay > 0) {
    sheet.getRange(DAY_ROW, FIRST_DAY_COL, 1, lastDay).setValues([dateRow]);
    sheet.getRange(YOIL_ROW, FIRST_DAY_COL, 1, lastDay).setValues([yoilRow]);
    sheet.getRange(DAY_ROW, FIRST_DAY_COL, 1, lastDay).setHorizontalAlignment("center").setFontWeight("bold").setFontColors([colorRow]);
    sheet.getRange(YOIL_ROW, FIRST_DAY_COL, 1, lastDay).setHorizontalAlignment("center").setFontColors([colorRow]);
  }

  const lastNameRow = NAME_START_ROW + Math.max(nameCount, 1) - 1;
  const lastDayCol = FIRST_DAY_COL + lastDay - 1;
  sheet.getRange(YOIL_ROW, 2, Math.max(nameCount, 1) + 1, lastDay + 1).setBorder(true, true, true, true, true, true);
  sheet.setFrozenRows(YOIL_ROW);
  sheet.setFrozenColumns(2);
  if (lastDay > 0) sheet.setColumnWidths(FIRST_DAY_COL, lastDay, 38);
  sheet.setColumnWidth(2, 140);
  if (nameCount) sheet.setRowHeights(NAME_START_ROW, nameCount, 24);
  sheet.autoResizeColumn(2);

  if (lastDay > 0 && nameCount > 0) {
    const checkRange = sheet.getRange(NAME_START_ROW, FIRST_DAY_COL, nameCount, lastDay);
    const cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(true).build();
    checkRange.setDataValidation(cbRule);
    checkRange.setValue(false);

    const applicantsRow = lastNameRow + 1;
    const capacityRow = applicantsRow + 1;
    const statusRow = applicantsRow + 2;

    sheet.getRange(applicantsRow, 2).setValue("RQ 신청자");
    sheet.getRange(capacityRow, 2).setValue("RQ 신청 가능일");
    sheet.getRange(statusRow, 2).setValue("RQ 현황");

    for (let c = FIRST_DAY_COL; c <= lastDayCol; c++) {
      const colA1 = columnToLetter_(c);
      sheet.getRange(applicantsRow, c).setFormula("=COUNTIF(" + colA1 + "$" + NAME_START_ROW + ":" + colA1 + "$" + lastNameRow + ", TRUE)");
      sheet.getRange(capacityRow, c).setFormula("=IFERROR(INDEX(matrix_RQ!$H:$H, MATCH(DATE($B$1,$B$2," + colA1 + "$" + DAY_ROW + "), matrix_RQ!$C:$C, 0)), 0)");
      sheet.getRange(statusRow, c).setFormula("=" + colA1 + "$" + capacityRow + "-" + colA1 + "$" + applicantsRow);
    }
  }
}

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  if (!/RQ(\s\(\d+\))?$/.test(sheetName)) return;

  var NAME_START_ROW = 5, DAY_ROW = 3, FIRST_DAY_COL = 3;
  var bCol = sheet.getRange(NAME_START_ROW, 2, sheet.getMaxRows() - NAME_START_ROW + 1, 1).getValues();
  var lastNameRow = NAME_START_ROW - 1;
  for (var i = 0; i < bCol.length; i++) {
    if (bCol[i][0]) lastNameRow = NAME_START_ROW + i;
    else break;
  }
  if (lastNameRow < NAME_START_ROW) return;

  var header = sheet.getRange(DAY_ROW, FIRST_DAY_COL, 1, sheet.getMaxColumns() - FIRST_DAY_COL + 1).getValues()[0];
  var lastDay = 0;
  for (i = 0; i < header.length; i++) {
    if (header[i] === '' || header[i] === null) break;
    lastDay++;
  }
  if (lastDay === 0) return;
  var lastDayCol = FIRST_DAY_COL + lastDay - 1;

  var er = e.range.getRow(), ec = e.range.getColumn(), erh = e.range.getNumRows(), ech = e.range.getNumColumns();
  var gridTop = NAME_START_ROW, gridBottom = lastNameRow, gridLeft = FIRST_DAY_COL, gridRight = lastDayCol;
  
  if (er > gridBottom || ec > gridRight || (er + erh - 1) < gridTop || (ec + ech - 1) < gridLeft) return;

  var changed = [];
  for (var r = er; r < er + erh; r++) {
    for (var c = ec; c < ec + ech; c++) {
      if (r < gridTop || r > gridBottom || c < gridLeft || c > gridRight) continue;
      var currentValue = sheet.getRange(r, c).getValue();
      if (currentValue === true) {
        if (erh === 1 && ech === 1) {
          if (e.oldValue !== true) changed.push([r, c]);
        } else {
          changed.push([r, c]);
        }
      }
    }
  }
  
  if (changed.length === 0) return;

  var gridValues = sheet.getRange(gridTop, gridLeft, gridBottom - gridTop + 1, gridRight - gridLeft + 1).getValues();
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

  var rowLimitViolated = false, colLimitViolated = false, violationMsg = '';
  var perPersonLimitCell = Number(sheet.getRange('Z1').getValue());
  var perPersonLimit = (isFinite(perPersonLimitCell) && perPersonLimitCell >= 1 && perPersonLimitCell <= 13) ? perPersonLimitCell : 5;

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
    SpreadsheetApp.flush();
    for (i = 0; i < changed.length; i++) {
      var cell = sheet.getRange(changed[i][0], changed[i][1]);
      cell.setValue(false);
    }
    SpreadsheetApp.flush();
    var msg = '❌ 신청 제한 초과\n\n' + violationMsg + '\n\n제한사항:\n· 개인: 최대 5일\n· 날짜별: matrix_RQ의 정원 내';
    SpreadsheetApp.getUi().alert('RQ 신청 제한', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

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
    if (d && d.getFullYear() === year && d.getMonth() + 1 === month && d.getDate() === day) {
      const n = Number(caps[i][0]);
      return isFinite(n) ? n : 0;
    }
  }
  return 0;
}

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

function columnToLetter_(column) {
  let temp = "", col = column;
  while (col > 0) {
    const rem = (col - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    col = Math.floor((col - 1) / 26);
  }
  return temp;
}

function rqConfirmMenu() {
  const ss = SpreadsheetApp.getActive();
  const rqNames = ss.getSheets().map(s => s.getName()).filter(n => /RQ(\s\(\d+\))?$/.test(n));
  if (rqNames.length === 0) {
    SpreadsheetApp.getUi().alert("확정 가능한 RQ 시트가 없습니다.");
    return;
  }
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
  const grid = rqSheet.getRange(NAME_START_ROW, FIRST_DAY_COL, names.length, days.length).getValues();
  const expectedHeader = ["year","month","date","day","name","is_rq","is_annual_leave","is_business_trip"];
  const dbHeader = dbSheet.getRange(1, 1, 1, expectedHeader.length).getValues()[0];
  for (let i = 0; i < expectedHeader.length; i++) {
    if (dbHeader[i] !== expectedHeader[i]) {
      ui.alert("DB_leave 헤더가 예상과 다릅니다. 다음과 같아야 합니다:\n" + expectedHeader.join(", "));
      return;
    }
  }
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
  const startRow = getFirstEmptyRow_(dbSheet, 1);
  dbSheet.getRange(startRow, 1, out.length, out[0].length).setValues(out);
  ui.alert('✅ "' + sheetName + '" 확정 완료\nDB_leave에 ' + out.length + "행을 추가했습니다.");
}

function getFirstEmptyRow_(sheet, colIndex) {
  const max = sheet.getMaxRows();
  const vals = sheet.getRange(1, colIndex, max).getValues();
  for (let r = max - 1; r >= 1; r--) {
    if (String(vals[r][0]).trim() !== "") return r + 2;
  }
  return 2;
}

/* ================================================================
   ✅ Schedule 자동 배정 (단계적 구현)
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
  const confirm = ui.alert("스케줄 자동 배정", targetMonth + "월의 스케줄을 자동으로 배정하시겠습니까?\n\n경고: 복잡한 제약조건으로 인해 완벽하지 않을 수 있습니다.\n배정 후 '스케줄 검증'을 실행하여 확인하세요.", ui.ButtonSet.YES_NO);
  
  if (confirm !== ui.Button.YES) return;
  
  try {
    assignScheduleForMonth_(targetMonth);
    ui.alert("✅ " + targetMonth + "월 스케줄 배정이 완료되었습니다.\n\n다음 단계:\n1. 'Schedule > 스케줄 검증' 메뉴를 실행하세요.\n2. 위반사항이 있다면 수동으로 조정하세요.");
  } catch (error) {
    ui.alert("❌ 오류 발생: " + error.toString());
  }
}

function assignScheduleForMonth_(targetMonth) {
  const ss = SpreadsheetApp.getActive();
  const dbSheet = ss.getSheetByName("DB_schedule");
  const data = dbSheet.getDataRange().getValues();
  const headers = data[0];
  
  const colIdx = {};
  const requiredCols = ["year","month","date","day","name","is_rq","employment_type_code","driving_class","gender_code","is_disposal_day","is_hh_cleaning_day","schedule"];
  
  for (let col of requiredCols) {
    colIdx[col] = headers.indexOf(col);
    if (colIdx[col] === -1) throw new Error("필수 열을 찾을 수 없습니다: " + col);
  }
  
  const monthData = [];
  const rowIndices = [];
  
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][colIdx.month]) === targetMonth) {
      monthData.push(data[i]);
      rowIndices.push(i + 1);
    }
  }
  
  if (monthData.length === 0) throw new Error(targetMonth + "월 데이터가 없습니다.");
  
  const year = monthData[0][colIdx.year];
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
  
  const prevMonthData = getPreviousMonthWorkDays_(year, targetMonth, employees);
  const schedule = generateOptimizedSchedule_(employees, dateInfo, year, targetMonth, prevMonthData);
  
  for (let i = 0; i < monthData.length; i++) {
    const name = monthData[i][colIdx.name];
    const date = Number(monthData[i][colIdx.date]);
    const scheduleValue = schedule[name] && schedule[name][date] ? schedule[name][date] : "OFF";
    dbSheet.getRange(rowIndices[i], colIdx.schedule + 1).setValue(scheduleValue);
  }
}

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
  
  let prevYear = year, prevMonth = month - 1;
  if (prevMonth < 1) {
    prevMonth = 12;
    prevYear--;
  }
  
  const prevMonthLastDate = new Date(year, month - 1, 0).getDate();
  const result = {};
  
  for (let name in employees) {
    result[name] = [];
  }
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (Number(row[yearCol]) === prevYear && Number(row[monthCol]) === prevMonth) {
      const name = row[nameCol];
      const date = Number(row[dateCol]);
      const schedule = row[scheduleCol];
      
      if (result[name] !== undefined && date >= prevMonthLastDate - 4) {
        result[name].push({
          date: date,
          schedule: schedule === "ON" ? "ON" : "OFF"
        });
      }
    }
  }
  
  for (let name in result) {
    result[name].sort((a, b) => a.date - b.date);
  }
  
  return result;
}

function generateOptimizedSchedule_(employees, dateInfo, year, month, prevMonthData) {
  const daysInMonth = new Date(year, month, 0).getDate();
  const schedule = {};
  
  // 초기화
  for (let name in employees) {
    schedule[name] = {};
    for (let d = 1; d <= daysInMonth; d++) {
      schedule[name][d] = "OFF";
    }
  }
  
  // === 단계 1: 확정 규칙 적용 ===
  
  // 규칙 11: PARTNER는 is_rq만 ON
  for (let name in employees) {
    if (employees[name].employment_type === "PARTNER") {
      for (let d = 1; d <= daysInMonth; d++) {
        if (employees[name].days[d] && employees[name].days[d].is_rq) {
          schedule[name][d] = "ON";
        }
      }
    }
  }
  
  // 규칙 15: 박정희와 최원진은 is_rq 제외하고 모두 ON
  for (let specialName of ["박정희", "최원진"]) {
    if (employees[specialName]) {
      for (let d = 1; d <= daysInMonth; d++) {
        if (employees[specialName].days[d]) {
          schedule[specialName][d] = employees[specialName].days[d].is_rq ? "OFF" : "ON";
        }
      }
    }
  }
  
  // 규칙 14: is_rq는 반드시 OFF
  for (let name in employees) {
    if (employees[name].employment_type !== "PARTNER" && name !== "최원진" && name !== "박정희") {
      for (let d = 1; d <= daysInMonth; d++) {
        if (employees[name].days[d] && employees[name].days[d].is_rq) {
          schedule[name][d] = "OFF";
        }
      }
    }
  }
  
  // === 단계 2: 정직원 스케줄 생성 ===
  
  const regularEmployees = Object.keys(employees).filter(name => 
    (employees[name].employment_type === "SV" || employees[name].employment_type === "FT") &&
    name !== "최원진" && name !== "박정희" && name !== "곽은태" && name !== "김성화"
  );
  
  // 규칙 10: 곽은태, 김성화 타겟 근무일
  const daysIn31 = [1, 3, 5, 7, 8, 10, 12];
  const specialTargetDays = daysIn31.includes(month) ? 22 : 21;
  
  // 간단한 그리디 알고리즘으로 스케줄 생성
  for (let name of regularEmployees) {
    let workDays = 0;
    let consecutiveWork = 0;
    let consecutiveOff = 0;
    
    // 전월 마지막 근무 확인 (규칙 3)
    if (prevMonthData[name] && prevMonthData[name].length > 0) {
      for (let i = prevMonthData[name].length - 1; i >= 0; i--) {
        if (prevMonthData[name][i].schedule === "ON") {
          consecutiveWork++;
        } else {
          break;
        }
      }
    }
    
    for (let d = 1; d <= daysInMonth; d++) {
      // 이미 OFF로 확정된 경우 스킵
      if (schedule[name][d] === "OFF" && employees[name].days[d] && employees[name].days[d].is_rq) {
        consecutiveWork = 0;
        consecutiveOff++;
        continue;
      }
      
      // 규칙 1: 최대 연속 근무 5일
      if (consecutiveWork >= 5) {
        schedule[name][d] = "OFF";
        consecutiveWork = 0;
        consecutiveOff++;
        continue;
      }
      
      // 규칙 2: 2일 이상 연속 휴무 우선
      if (consecutiveOff > 0 && consecutiveOff < 2 && workDays < 22) {
        schedule[name][d] = "OFF";
        consecutiveWork = 0;
        consecutiveOff++;
        continue;
      }
      
      // 기본적으로 ON
      schedule[name][d] = "ON";
      workDays++;
      consecutiveWork++;
      consecutiveOff = 0;
    }
  }
  
  // 곽은태, 김성화 특수 처리
  for (let specialName of ["곽은태", "김성화"]) {
    if (!employees[specialName]) continue;
    
    let workDays = 0;
    let consecutiveWork = 0;
    
    if (prevMonthData[specialName] && prevMonthData[specialName].length > 0) {
      for (let i = prevMonthData[specialName].length - 1; i >= 0; i--) {
        if (prevMonthData[specialName][i].schedule === "ON") {
          consecutiveWork++;
        } else {
          break;
        }
      }
    }
    
    for (let d = 1; d <= daysInMonth; d++) {
      if (schedule[specialName][d] === "OFF" && employees[specialName].days[d] && employees[specialName].days[d].is_rq) {
        consecutiveWork = 0;
        continue;
      }
      
      if (consecutiveWork >= 5 || workDays >= specialTargetDays) {
        schedule[specialName][d] = "OFF";
        consecutiveWork = 0;
      } else {
        schedule[specialName][d] = "ON";
        workDays++;
        consecutiveWork++;
      }
    }
  }
  
  // === 단계 3: 일별 제약조건 보정 ===
  
  for (let d = 1; d <= daysInMonth; d++) {
    let attempts = 0;
    while (attempts < 10) {
      const dayEmployees = getDayEmployees_(schedule, employees, d);
      const violations = checkDayConstraints_(dayEmployees, dateInfo[d], employees);
      
      if (violations.length === 0) break;
      
      // 위반사항 해결 시도
      for (let violation of violations) {
        if (violation.type === "min_staff" && violation.current < violation.required) {
          // 직원 추가 필요
          addEmployeeToDayIfPossible_(schedule, employees, d, regularEmployees);
        } else if (violation.type === "driver" && violation.current < violation.required) {
          // 운전가능자 추가
          addDriverToDayIfPossible_(schedule, employees, d, regularEmployees);
        } else if (violation.type === "supervisor" && violation.current < violation.required) {
          // 슈퍼바이저 추가
          addSupervisorToDayIfPossible_(schedule, employees, d, regularEmployees);
        } else if (violation.type === "male_hh" && violation.current < violation.required) {
          // 남성 직원 추가
          addMaleEmployeeToDayIfPossible_(schedule, employees, d, regularEmployees);
        } else if (violation.type === "male_disposal" && violation.current < violation.required) {
          // 남성 직원 추가
          addMaleEmployeeToDayIfPossible_(schedule, employees, d, regularEmployees);
        } else if (violation.type === "ft_sv_min" && violation.current < violation.required) {
          // 정직원 추가
          addFTSVToDayIfPossible_(schedule, employees, d, regularEmployees);
        }
      }
      
      attempts++;
    }
  }
  
  return schedule;
}

function getDayEmployees_(schedule, employees, day) {
  const result = [];
  for (let name in schedule) {
    if (schedule[name][day] === "ON") {
      result.push(name);
    }
  }
  return result;
}

function checkDayConstraints_(dayEmployees, dayInfo, employees) {
  const violations = [];
  
  // 규칙 9: 최소 10명
  if (dayEmployees.length < 10) {
    violations.push({type: "min_staff", current: dayEmployees.length, required: 10});
  }
  
  // 규칙 4: 운전가능자 최소 2명
  const drivers = dayEmployees.filter(name => {
    const dc = employees[name].driving_class;
    return dc === "ALL_VEHICLES" || dc === "SMALL_ONLY";
  });
  if (drivers.length < 2) {
    violations.push({type: "driver", current: drivers.length, required: 2});
  }
  
  // 규칙 5: 슈퍼바이저 최소 1명
  const supervisors = dayEmployees.filter(name => {
    const et = employees[name].employment_type;
    return et === "SV" || et === "SSV";
  });
  if (supervisors.length < 1) {
    violations.push({type: "supervisor", current: supervisors.length, required: 1});
  }
  
  // 규칙 6: 월목토(is_hh_cleaning) 남성 2명 이상
  if (dayInfo && dayInfo.is_hh_cleaning) {
    const males = dayEmployees.filter(name => employees[name].gender === "M");
    if (males.length < 2) {
      violations.push({type: "male_hh", current: males.length, required: 2});
    }
  }
  
  // 규칙 7: 수토(is_disposal) 남성 1명 이상
  if (dayInfo && dayInfo.is_disposal) {
    const males = dayEmployees.filter(name => employees[name].gender === "M");
    if (males.length < 1) {
      violations.push({type: "male_disposal", current: males.length, required: 1});
    }
  }
  
  // 규칙 8: 정직원 최소 수 (요일별)
  const ftSv = dayEmployees.filter(name => {
    const et = employees[name].employment_type;
    return et === "SV" || et === "FT";
  });
  
  if (dayInfo) {
    const dow = dayInfo.day_of_week;
    let minFtSv = 7;
    if (dow === "월" || dow === "목" || dow === "토") minFtSv = 9;
    else if (dow === "수") minFtSv = 6;
    
    if (ftSv.length < minFtSv) {
      violations.push({type: "ft_sv_min", current: ftSv.length, required: minFtSv, dow: dow});
    }
  }
  
  return violations;
}

function addEmployeeToDayIfPossible_(schedule, employees, day, candidates) {
  for (let name of candidates) {
    if (schedule[name][day] === "OFF" && 
        (!employees[name].days[day] || !employees[name].days[day].is_rq) &&
        canWorkOnDay_(schedule, name, day)) {
      schedule[name][day] = "ON";
      return true;
    }
  }
  return false;
}

function addDriverToDayIfPossible_(schedule, employees, day, candidates) {
  for (let name of candidates) {
    const dc = employees[name].driving_class;
    if ((dc === "ALL_VEHICLES" || dc === "SMALL_ONLY") &&
        schedule[name][day] === "OFF" && 
        (!employees[name].days[day] || !employees[name].days[day].is_rq) &&
        canWorkOnDay_(schedule, name, day)) {
      schedule[name][day] = "ON";
      return true;
    }
  }
  return false;
}

function addSupervisorToDayIfPossible_(schedule, employees, day, candidates) {
  for (let name of candidates) {
    const et = employees[name].employment_type;
    if ((et === "SV" || et === "SSV") &&
        schedule[name][day] === "OFF" && 
        (!employees[name].days[day] || !employees[name].days[day].is_rq) &&
        canWorkOnDay_(schedule, name, day)) {
      schedule[name][day] = "ON";
      return true;
    }
  }
  return false;
}

function addMaleEmployeeToDayIfPossible_(schedule, employees, day, candidates) {
  for (let name of candidates) {
    if (employees[name].gender === "M" &&
        schedule[name][day] === "OFF" && 
        (!employees[name].days[day] || !employees[name].days[day].is_rq) &&
        canWorkOnDay_(schedule, name, day)) {
      schedule[name][day] = "ON";
      return true;
    }
  }
  return false;
}

function addFTSVToDayIfPossible_(schedule, employees, day, candidates) {
  for (let name of candidates) {
    const et = employees[name].employment_type;
    if ((et === "SV" || et === "FT") &&
        schedule[name][day] === "OFF" && 
        (!employees[name].days[day] || !employees[name].days[day].is_rq) &&
        canWorkOnDay_(schedule, name, day)) {
      schedule[name][day] = "ON";
      return true;
    }
  }
  return false;
}

function canWorkOnDay_(schedule, name, day) {
  // 규칙 1: 최대 연속 근무 5일 체크
  let consecutiveWork = 0;
  for (let d = day - 1; d >= 1 && d >= day - 5; d--) {
    if (schedule[name][d] === "ON") {
      consecutiveWork++;
    } else {
      break;
    }
  }
  
  return consecutiveWork < 5;
}

/* ================================================================
   ✅ 스케줄 검증 기능
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
    const violations = validateScheduleForMonth_(targetMonth);
    
    if (violations.length === 0) {
      ui.alert("✅ " + targetMonth + "월 스케줄 검증 완료\n\n모든 제약조건을 만족합니다!");
    } else {
      let report = "⚠️ " + targetMonth + "월 스케줄 위반사항\n\n총 " + violations.length + "건의 위반사항이 발견되었습니다.\n\n";
      
      // 처음 10개만 표시
      for (let i = 0; i < Math.min(violations.length, 10); i++) {
        report += (i + 1) + ". " + violations[i] + "\n";
      }
      
      if (violations.length > 10) {
        report += "\n... 외 " + (violations.length - 10) + "건";
      }
      
      ui.alert(report);
    }
  } catch (error) {
    ui.alert("❌ 오류 발생: " + error.toString());
  }
}

function validateScheduleForMonth_(targetMonth) {
  const ss = SpreadsheetApp.getActive();
  const dbSheet = ss.getSheetByName("DB_schedule");
  const data = dbSheet.getDataRange().getValues();
  const headers = data[0];
  
  const colIdx = {};
  const requiredCols = ["year","month","date","day","name","is_rq","employment_type_code","driving_class","gender_code","is_disposal_day","is_hh_cleaning_day","schedule"];
  
  for (let col of requiredCols) {
    colIdx[col] = headers.indexOf(col);
    if (colIdx[col] === -1) throw new Error("필수 열을 찾을 수 없습니다: " + col);
  }
  
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
      day_of_week: row[colIdx.day]
    };
    
    schedule[name][date] = row[colIdx.schedule] === "ON" ? "ON" : "OFF";
    
    if (!dateInfo[date]) {
      dateInfo[date] = {
        day_of_week: row[colIdx.day],
        is_disposal: row[colIdx.is_disposal_day] === true || row[colIdx.is_disposal_day] === "TRUE",
        is_hh_cleaning: row[colIdx.is_hh_cleaning_day] === true || row[colIdx.is_hh_cleaning_day] === "TRUE"
      };
    }
  }
  
  const violations = [];
  const daysInMonth = new Date(year, targetMonth, 0).getDate();
  
  // 일별 검증
  for (let d = 1; d <= daysInMonth; d++) {
    const dayEmployees = getDayEmployees_(schedule, employees, d);
    const dayViolations = checkDayConstraints_(dayEmployees, dateInfo[d], employees);
    
    for (let v of dayViolations) {
      violations.push(targetMonth + "월 " + d + "일: " + formatViolation_(v));
    }
  }
  
  // 직원별 검증
  for (let name in employees) {
    if (name === "최원진") continue; // 예외
    
    // 규칙 14: is_rq 체크
    for (let d = 1; d <= daysInMonth; d++) {
      if (employees[name].days[d] && employees[name].days[d].is_rq && schedule[name][d] === "ON") {
        violations.push(name + " " + targetMonth + "월 " + d + "일: RQ 신청일인데 출근으로 배정됨");
      }
    }
    
    // 규칙 1: 최대 연속 근무 5일
    let consecutiveWork = 0;
    for (let d = 1; d <= daysInMonth; d++) {
      if (schedule[name][d] === "ON") {
        consecutiveWork++;
        if (consecutiveWork > 5) {
          violations.push(name + " " + targetMonth + "월 " + d + "일: 연속 근무 5일 초과 (" + consecutiveWork + "일)");
        }
      } else {
        consecutiveWork = 0;
      }
    }
  }
  
  return violations;
}

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
  } else if (v.type === "ft_sv_min") {
    return v.dow + "요일 정직원 부족 (현재: " + v.current + "명, 필요: " + v.required + "명)";
  }
  return "알 수 없는 위반";
}

/* ================================================================
   ✅ 스케줄 시각화 시트 생성
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

function createScheduleVisualizationSheet_(targetMonth) {
  const ss = SpreadsheetApp.getActive();
  const dbSheet = ss.getSheetByName("DB_schedule");
  const empSheet = ss.getSheetByName("matrix_employees");
  
  if (!empSheet) throw new Error("matrix_employees 시트를 찾을 수 없습니다.");
  
  const data = dbSheet.getDataRange().getValues();
  const headers = data[0];
  
  const colIdx = {};
  const requiredCols = ["year","month","date","day","name","employment_type_code","schedule"];
  
  for (let col of requiredCols) {
    colIdx[col] = headers.indexOf(col);
    if (colIdx[col] === -1) throw new Error("필수 열을 찾을 수 없습니다: " + col);
  }
  
  // 해당 월 데이터 수집
  const monthData = [];
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][colIdx.month]) === targetMonth) {
      monthData.push(data[i]);
    }
  }
  
  if (monthData.length === 0) throw new Error(targetMonth + "월 데이터가 없습니다.");
  
  const year = monthData[0][colIdx.year];
  const employeeSchedule = {}; // {name: {date: "ON/OFF"}}
  const employeeType = {}; // {name: employment_type}
  const dates = new Set();
  
  for (let row of monthData) {
    const name = row[colIdx.name];
    const date = Number(row[colIdx.date]);
    const schedule = row[colIdx.schedule];
    const empType = row[colIdx.employment_type_code];
    
    if (!employeeSchedule[name]) {
      employeeSchedule[name] = {};
      employeeType[name] = empType;
    }
    
    employeeSchedule[name][date] = schedule === "ON" ? "ON" : "OFF";
    dates.add(date);
  }
  
  // 날짜 정렬
  const sortedDates = Array.from(dates).sort((a, b) => a - b);
  const daysInMonth = sortedDates[sortedDates.length - 1];
  
  // matrix_employees에서 직원 순서 가져오기
  const empData = empSheet.getDataRange().getValues();
  const empHeaders = empData[0];
  const empNameCol = empHeaders.indexOf("name");
  const empTypeCol = empHeaders.indexOf("employment_type_code");
  const empDrivingCol = empHeaders.indexOf("driving_class");
  
  if (empNameCol === -1) throw new Error("matrix_employees에 name 열이 없습니다.");
  
  const employeeNames = [];
  const employeeInfo = {}; // {name: {type, driving}}
  
  for (let i = 1; i < empData.length; i++) {
    const name = empData[i][empNameCol];
    const type = empData[i][empTypeCol];
    const driving = empData[i][empDrivingCol];
    
    if (name && employeeSchedule[name]) {
      employeeNames.push(name);
      employeeInfo[name] = {
        type: type,
        driving: driving
      };
    }
  }
  
  // 새 시트 생성
  const baseName = String(year).slice(-2) + " " + MON_ABBR[targetMonth - 1] + " Schedule";
  const sheetName = getUniqueSheetName_(baseName);
  const sheet = ss.insertSheet(sheetName);
  
  // 헤더 영역
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
  
  // 직원 이름 쓰기 (B열)
  const nameValues = employeeNames.map(name => [name]);
  if (nameValues.length > 0) {
    sheet.getRange(NAME_START_ROW, 2, nameValues.length, 1).setValues(nameValues);
  }
  
  // 날짜 헤더 쓰기
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
  
  // 스케줄 데이터 쓰기
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
    
    // ON은 초록, OFF는 회색 배경
    const colors = scheduleValues.map(row => 
      row.map(val => val === "ON" ? "#d9ead3" : "#f3f3f3")
    );
    scheduleRange.setBackgrounds(colors);
  }
  
  // 표 스타일
  const lastNameRow = NAME_START_ROW + Math.max(employeeNames.length, 1) - 1;
  sheet.getRange(YOIL_ROW, 2, Math.max(employeeNames.length, 1) + 1, daysInMonth + 1)
    .setBorder(true, true, true, true, true, true);
  
  sheet.setFrozenRows(YOIL_ROW);
  sheet.setFrozenColumns(2);
  if (daysInMonth > 0) sheet.setColumnWidths(FIRST_DAY_COL, daysInMonth, 45);
  sheet.setColumnWidth(2, 100);
  if (employeeNames.length > 0) sheet.setRowHeights(NAME_START_ROW, employeeNames.length, 24);
  
  // 요약 행 추가
  const summaryRow = lastNameRow + 1;
  const onDutyRow = summaryRow;
  const driverRow = summaryRow + 1;
  const supervisorRow = summaryRow + 2;
  const ftSvRow = summaryRow + 3;
  
  sheet.getRange(onDutyRow, 2).setValue("출근 인원");
  sheet.getRange(driverRow, 2).setValue("운전가능자");
  sheet.getRange(supervisorRow, 2).setValue("슈퍼바이저");
  sheet.getRange(ftSvRow, 2).setValue("정직원(FT/SV)");
  
  // 요약 공식 - 각 날짜별 집계
  for (let col = 0; col < daysInMonth; col++) {
    const colLetter = columnToLetter_(FIRST_DAY_COL + col);
    const startRow = NAME_START_ROW;
    const endRow = lastNameRow;
    
    // 출근 인원 (ON 카운트)
    sheet.getRange(onDutyRow, FIRST_DAY_COL + col)
      .setFormula('=COUNTIF(' + colLetter + startRow + ':' + colLetter + endRow + ',"ON")');
    
    // 운전가능자 카운트
    let driverFormula = '=SUMPRODUCT((';
    for (let i = 0; i < employeeNames.length; i++) {
      const name = employeeNames[i];
      const rowNum = NAME_START_ROW + i;
      const isDrvier = employeeInfo[name].driving === "ALL_VEHICLES" || employeeInfo[name].driving === "SMALL_ONLY";
      if (i > 0) driverFormula += '+';
      driverFormula += '(' + colLetter + rowNum + '="ON")*' + (isDrvier ? '1' : '0');
    }
    driverFormula += '))';
    sheet.getRange(driverRow, FIRST_DAY_COL + col).setFormula(driverFormula);
    
    // 슈퍼바이저 카운트
    let svFormula = '=SUMPRODUCT((';
    for (let i = 0; i < employeeNames.length; i++) {
      const name = employeeNames[i];
      const rowNum = NAME_START_ROW + i;
      const isSV = employeeInfo[name].type === "SV" || employeeInfo[name].type === "SSV";
      if (i > 0) svFormula += '+';
      svFormula += '(' + colLetter + rowNum + '="ON")*' + (isSV ? '1' : '0');
    }
    svFormula += '))';
    sheet.getRange(supervisorRow, FIRST_DAY_COL + col).setFormula(svFormula);
    
    // 정직원(FT/SV) 카운트
    let ftSvFormula = '=SUMPRODUCT((';
    for (let i = 0; i < employeeNames.length; i++) {
      const name = employeeNames[i];
      const rowNum = NAME_START_ROW + i;
      const isFtSv = employeeInfo[name].type === "SV" || employeeInfo[name].type === "FT";
      if (i > 0) ftSvFormula += '+';
      ftSvFormula += '(' + colLetter + rowNum + '="ON")*' + (isFtSv ? '1' : '0');
    }
    ftSvFormula += '))';
    sheet.getRange(ftSvRow, FIRST_DAY_COL + col).setFormula(ftSvFormula);
  }
  
  // 요약 행 스타일
  sheet.getRange(onDutyRow, 2, 4, daysInMonth + 1)
    .setFontWeight("bold")
    .setBackground("#fff2cc");
}
