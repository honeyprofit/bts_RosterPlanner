/*** CONFIG ***/
const SOURCE_SHEET_NAME = 'employees'; // 직원 데이터 시트 (A:ID, B:이름)
const KR_HOLIDAY_CAL_ID = 'ko.south_korea#holiday@group.v.calendar.google.com'; // 한국 공휴일 캘린더
const MON_ABBR = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RQ')
    .addItem('신규 월간 시트 생성', 'rqCreateNewSheet')
    .addToUi();
}

// RQ 버튼 동작: 연/월 입력 → 새 시트 생성(이름 자동, 중복 시 (1),(2)...)
function rqCreateNewSheet() {
  const ui = SpreadsheetApp.getUi();

  const y = ui.prompt('연도 입력', '예: 2025', ui.ButtonSet.OK_CANCEL);
  if (y.getSelectedButton() !== ui.Button.OK) return;
  const year = Number(String(y.getResponseText()).trim());
  if (!Number.isInteger(year) || year < 1900 || year > 2100) {
    ui.alert('연도 입력이 올바르지 않습니다. (예: 2025)'); return;
  }

  const m = ui.prompt('월 입력', '1 ~ 12 중 하나 (예: 10)', ui.ButtonSet.OK_CANCEL);
  if (m.getSelectedButton() !== ui.Button.OK) return;
  const month = Number(String(m.getResponseText()).trim());
  if (!Number.isInteger(month) || month < 1 || month > 12) {
    ui.alert('월 입력이 올바르지 않습니다. (1~12)'); return;
  }

  const baseName = `${String(year).slice(-2)} ${MON_ABBR[month - 1]} RQ`;
  const sheetName = getUniqueSheetName_(baseName);

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.insertSheet(sheetName); // ✅ 항상 새 시트 생성
  fillRQContent_(sheet, year, month);      // 내용 채우기

  SpreadsheetApp.getUi().alert(`완료! 새 시트 "${sheetName}"를 생성했습니다.`);
}

// 동일 이름이 있으면 (1),(2)... 번호를 붙여 유일 이름 생성
function getUniqueSheetName_(base) {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(base)) return base;
  let i = 1;
  while (ss.getSheetByName(`${base} (${i})`)) i++;
  return `${base} (${i})`;
}

// 시트 내용 작성 (직원 목록 + 날짜/요일 + 공휴일/주말 색상)
function fillRQContent_(sheet, year, month) {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!src) throw new Error(`시트 "${SOURCE_SHEET_NAME}" 를 찾을 수 없습니다.`);

  // 좌측 상단 파라미터 표시
  sheet.getRange('A1').setValue('년도');
  sheet.getRange('B1').setValue(year);
  sheet.getRange('A2').setValue('월');
  sheet.getRange('B2').setValue(month);

  // 직원 목록 불러와 ID 오름차순 정렬
  const lastRow = src.getLastRow();
  const raw = lastRow > 1 ? src.getRange(2, 1, lastRow - 1, 2).getValues() : []; // A:ID, B:이름
  const rows = raw.filter(r => r[0] && r[1]).sort((a,b) => (a[0] > b[0]) - (a[0] < b[0]));

  const names = rows.map(r => [r[1]]);
  if (names.length) sheet.getRange(4, 2, names.length, 1).setValues(names); // B4↓

  // 헤더 계산
  const lastDay = new Date(year, month, 0).getDate();
  const weekdays = ['일','월','화','수','목','금','토'];
  const holidaySet = getKoreanHolidaySet_(year, month);

  const dateRow = [];
  const yoilRow = [];
  const colorRow = []; // 1×N

  for (let d = 1; d <= lastDay; d++) {
    const dt = new Date(year, month - 1, d);
    const dow = dt.getDay();
    const isHoliday = holidaySet.has(d);
    const color = (dow === 0 || isHoliday) ? '#d93025' : (dow === 6) ? '#1a73e8' : '#000000';

    dateRow.push(d);
    yoilRow.push(weekdays[dow]);
    colorRow.push(color);
  }

  // 헤더 쓰기 (C열부터)
  if (lastDay > 0) {
    sheet.getRange(3, 3, 1, lastDay).setValues([dateRow]); // 일자
    sheet.getRange(4, 3, 1, lastDay).setValues([yoilRow]); // 요일
    const dateRange = sheet.getRange(3, 3, 1, lastDay);
    const yoilRange = sheet.getRange(4, 3, 1, lastDay);
    dateRange.setHorizontalAlignment('center').setFontWeight('bold').setFontColors([colorRow]);
    yoilRange.setHorizontalAlignment('center').setFontColors([colorRow]);
  }

  // 라벨 / 테두리 / 고정 / 크기
  sheet.getRange('A3').setValue('일자');
  sheet.getRange('A4').setValue('요일');
  sheet.getRange(3, 2, Math.max(names.length, 1) + 1, lastDay + 1)
       .setBorder(true, true, true, true, true, true);

  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(2);
  sheet.setColumnWidths(3, lastDay, 38);
  sheet.setColumnWidth(2, 140);
  if (names.length) sheet.setRowHeights(4, names.length, 24);
  sheet.autoResizeColumn(2);
}

// 해당 연/월의 한국 공휴일을 day(Set)로 반환
function getKoreanHolidaySet_(year, month) {
  const cal = CalendarApp.getCalendarById(KR_HOLIDAY_CAL_ID);
  const start = new Date(year, month - 1, 1);
  const end = new Date(year, month, 1);
  const events = cal.getEvents(start, end);
  const set = new Set();

  events.forEach(ev => {
    const s = ev.isAllDayEvent() ? ev.getAllDayStartDate() : ev.getStartTime();
    const e = ev.isAllDayEvent() ? ev.getAllDayEndDate() : ev.getEndTime();
    for (let d = new Date(s.getFullYear(), s.getMonth(), s.getDate()); d < e; d.setDate(d.getDate() + 1)) {
      if (d.getFullYear() === year && d.getMonth() + 1 === month) set.add(d.getDate());
    }
  });
  return set;
}
