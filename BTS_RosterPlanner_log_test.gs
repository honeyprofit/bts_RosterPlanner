function onEdit(e) {
  // ✅ 이벤트 객체 존재 확인
  if (!e || !e.range) {
    Logger.log('onEdit: 이벤트 객체 없음');
    return;
  }

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  
  Logger.log('onEdit 실행: 시트명=' + sheetName);
  
  // 대상: 이름이 "... RQ" 또는 "... RQ (n)" 로 끝나는 시트
  if (!/RQ(\s\(\d+\))?$/.test(sheetName)) {
    Logger.log('RQ 시트 아님, 종료');
    return;
  }

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
  if (lastNameRow < NAME_START_ROW) {
    Logger.log('직원 목록 없음, 종료');
    return;
  }

  // 날짜 마지막 열(3행 연속 값)
  var header = sheet.getRange(DAY_ROW, FIRST_DAY_COL, 1, sheet.getMaxColumns() - FIRST_DAY_COL + 1).getValues()[0];
  var lastDay = 0;
  for (i = 0; i < header.length; i++) {
    if (header[i] === '' || header[i] === null) break;
    lastDay++;
  }
  if (lastDay === 0) {
    Logger.log('날짜 없음, 종료');
    return;
  }
  var lastDayCol = FIRST_DAY_COL + lastDay - 1;

  // 편집 범위와 체크 그리드 교집합 확인
  var er = e.range.getRow(), ec = e.range.getColumn(), erh = e.range.getNumRows(), ech = e.range.getNumColumns();
  var gridTop = NAME_START_ROW, gridBottom = lastNameRow, gridLeft = FIRST_DAY_COL, gridRight = lastDayCol;
  
  Logger.log('편집범위: 행' + er + '~' + (er+erh-1) + ', 열' + ec + '~' + (ec+ech-1));
  Logger.log('체크그리드: 행' + gridTop + '~' + gridBottom + ', 열' + gridLeft + '~' + gridRight);
  
  if (er > gridBottom || ec > gridRight || (er + erh - 1) < gridTop || (ec + ech - 1) < gridLeft) {
    Logger.log('체크 그리드 밖 편집, 종료');
    return;
  }

  // ✅ FALSE → TRUE로 변경된 셀만 수집
  var changed = [];
  
  for (var r = er; r < er + erh; r++) {
    for (var c = ec; c < ec + ech; c++) {
      if (r < gridTop || r > gridBottom || c < gridLeft || c > gridRight) continue;
      
      var currentValue = sheet.getRange(r, c).getValue();
      
      // ✅ 현재 값이 true인 경우에만 체크 (체크 해제는 무시)
      if (currentValue === true) {
        // oldValue는 단일 셀 편집에만 존재
        if (erh === 1 && ech === 1) {
          // oldValue가 false이거나 없는 경우만 (새로 체크한 경우)
          if (e.oldValue !== true) {
            changed.push([r, c]);
            Logger.log('체크 감지: 행' + r + ', 열' + c + ' (이전값=' + e.oldValue + ')');
          }
        } else {
          // 다중 편집: 일단 true인 것 수집 (이전 값 비교 불가)
          changed.push([r, c]);
          Logger.log('다중편집 체크: 행' + r + ', 열' + c);
        }
      }
    }
  }
  
  if (changed.length === 0) {
    Logger.log('체크 변경 없음 (해제만 함), 종료');
    return;
  }

  Logger.log('총 ' + changed.length + '개 셀 체크됨');

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
      Logger.log('날짜 ' + day + '일 정원: ' + capacityByCol[cc]);
    }
  }

  // 규칙 위반 체크
  var rowLimitViolated = false;
  var colLimitViolated = false;
  var violationMsg = '';

  // 개인 5개 초과?
  for (i = 0; i < changed.length; i++) {
    var rr = changed[i][0];
    var count = rowCount[rr - gridTop];
    Logger.log('행' + rr + ' 개인 체크 수: ' + count);
    
    if (count > 5) { 
      rowLimitViolated = true;
      violationMsg = '개인당 최대 5일까지만 신청할 수 있습니다. (현재: ' + count + '개)';
      Logger.log('개인 제한 위반!');
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
      
      Logger.log('날짜 ' + dayNum + '일: 신청자=' + applicants + ', 정원=' + cap);
      
      if (applicants > cap) { 
        colLimitViolated = true;
        violationMsg = dayNum + '일은 정원(' + cap + '명)이 초과되었습니다. (신청자: ' + applicants + '명)';
        Logger.log('날짜 정원 위반!');
        break;
      }
    }
  }

  if (rowLimitViolated || colLimitViolated) {
    Logger.log('위반 감지! 체크 취소 실행');
    
    // 이번 편집에서 TRUE로 바뀐 모든 칸을 되돌림
    for (i = 0; i < changed.length; i++) {
      sheet.getRange(changed[i][0], changed[i][1]).setValue(false);
      Logger.log('되돌림: 행' + changed[i][0] + ', 열' + changed[i][1]);
    }

    var msg = '신청 제한 때문에 체크가 취소되었습니다.\n\n' + violationMsg + '\n\n제한사항:\n· 개인: 최대 5일\n· 날짜별: matrix_RQ의 정원 내';
    
    // ✅ UI 알림 (단순 트리거에서는 toast가 안 보일 수 있음)
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'RQ 제한', 5);
    } catch(e) {
      Logger.log('Toast 실패: ' + e);
    }
    
    // Alert 표시
    try {
      SpreadsheetApp.getUi().alert('RQ 신청 제한', msg, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch(e) {
      Logger.log('Alert 실패: ' + e);
    }
  } else {
    Logger.log('위반 없음, 체크 허용');
  }
}

/**
 * matrix_RQ 시트에서 (연,월,일) 일치 행의 정원(H열) 값을 반환
 */
function getCapacityFromMatrix_(year, month, day) {
  var ms = SpreadsheetApp.getActive().getSheetByName('matrix_RQ');
  if (!ms) {
    Logger.log('matrix_RQ 시트 없음');
    return 0;
  }

  var lastRow = ms.getLastRow();
  if (lastRow < 2) {
    Logger.log('matrix_RQ 데이터 없음');
    return 0;
  }

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
      Logger.log('정원 찾음: ' + year + '-' + month + '-' + day + ' = ' + n);
      return isFinite(n) ? n : 0;
    }
  }
  
  Logger.log('정원 못 찾음: ' + year + '-' + month + '-' + day);
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
