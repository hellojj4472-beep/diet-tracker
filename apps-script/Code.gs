// 다이어트표 - Google Sheets 연동용 Apps Script
// 이 파일은 참고/백업용입니다. 실제로는 구글시트의 확장프로그램 > Apps Script 편집기에 붙여넣어 사용하세요.

var SHEET_NAME = '기록';
var TOKEN = '여기에-원하는-비밀번호처럼-긴-문자열을-넣으세요'; // 예: diet-8x3kQ9zL2m

var HEADERS = ['날짜','체중','허리둘레','허벅지둘레','골반둘레','측정건너뜀','섭취칼로리','탄수화물','단백질','지방','단순당','운동소모칼로리','순섭취칼로리','수면시간','배변','특이사항태그','메모','식사상세JSON','운동상세JSON','일정','근무'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else {
    // 예전에 만든 시트는 '근무' 칸이 없을 수 있으니 헤더 비교해서 있으면 추가
    var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (existingHeaders.indexOf('근무') === -1) {
      sheet.getRange(1, existingHeaders.length + 1).setValue('근무');
    }
  }
  return sheet;
}

function checkToken_(token) {
  if (token !== TOKEN) {
    throw new Error('unauthorized');
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function formatDate_(v) {
  try {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return String(v);
  }
}

// 일정 칸은 JSON 배열로 저장됨. 예전 버전(문자열 하나)이 남아있으면 배열 1개짜리로 변환
function parseScheduleItems_(v) {
  if (!v) return [];
  try {
    var parsed = JSON.parse(v);
    if (Array.isArray(parsed)) return parsed;
  } catch (e) { /* 예전 형식(순수 문자열) */ }
  return [String(v)];
}

function parseJsonOr_(v, fallback) {
  if (!v) return fallback;
  try { return JSON.parse(v); } catch (e) { return fallback; }
}

function rowToObject_(r) {
  return {
    date: formatDate_(r[0]), weight: r[1], waist: r[2], thigh: r[3], hip: r[4], skipMeasure: r[5],
    kcal: r[6], carb: r[7], prot: r[8], fat: r[9], sugar: r[10], burn: r[11], net: r[12],
    sleep: r[13], bowel: r[14], tags: r[15], note: r[16],
    meals: parseJsonOr_(r[17], {}), exercises: parseJsonOr_(r[18], []),
    scheduleItems: parseScheduleItems_(r[19]), work: r[20] || ''
  };
}

// 날짜 칸은 항상 텍스트로 고정해서 저장 (구글시트가 자동으로 '진짜 날짜'로 바꿔버리는 것 방지)
function writeRow_(sheet, rowIndex, row) {
  sheet.getRange(rowIndex, 1).setNumberFormat('@').setValue(row[0]);
  if (row.length > 1) {
    sheet.getRange(rowIndex, 2, 1, row.length - 1).setValues([row.slice(1)]);
  }
}

// 히스토리 조회: GET {웹앱주소}?token=...&days=30
function doGet(e) {
  try {
    checkToken_(e.parameter.token);
    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    var rows = data.slice(1).map(rowToObject_);
    var days = e.parameter.days ? parseInt(e.parameter.days, 10) : 30;
    var recent = rows.slice(-days);
    return jsonOutput_({ ok: true, rows: recent });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err) });
  }
}

function findRowIndex_(data, dateStr) {
  for (var i = 1; i < data.length; i++) {
    if (formatDate_(data[i][0]) === dateStr) return i + 1;
  }
  return -1;
}

// 일정 칸 하나만 업데이트 (그 날의 다른 기록은 건드리지 않음)
function setScheduleOnly_(sheet, payload) {
  var data = sheet.getDataRange().getValues();
  var rowIndex = findRowIndex_(data, payload.date);
  if (rowIndex === -1) {
    rowIndex = sheet.getLastRow() + 1;
    sheet.getRange(rowIndex, 1).setNumberFormat('@').setValue(payload.date);
  }
  var scheduleCol = HEADERS.indexOf('일정') + 1;
  sheet.getRange(rowIndex, scheduleCol).setValue(JSON.stringify(payload.scheduleItems || []));
  return { ok: true };
}

// 근무 칸 하나만 업데이트 (그 날의 다른 기록은 건드리지 않음)
function setWorkOnly_(sheet, payload) {
  var data = sheet.getDataRange().getValues();
  var rowIndex = findRowIndex_(data, payload.date);
  if (rowIndex === -1) {
    rowIndex = sheet.getLastRow() + 1;
    sheet.getRange(rowIndex, 1).setNumberFormat('@').setValue(payload.date);
  }
  var workCol = HEADERS.indexOf('근무') + 1;
  sheet.getRange(rowIndex, workCol).setValue(payload.work || '');
  return { ok: true };
}

// 오늘 기록 저장/갱신: POST JSON body { token, date, weight, ... }
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    checkToken_(payload.token);
    var sheet = getSheet_();
    if (payload.action === 'schedule') {
      return jsonOutput_(setScheduleOnly_(sheet, payload));
    }
    if (payload.action === 'work') {
      return jsonOutput_(setWorkOnly_(sheet, payload));
    }
    var data = sheet.getDataRange().getValues();
    var rowIndex = findRowIndex_(data, payload.date);
    var existingRow = rowIndex !== -1 ? data[rowIndex - 1] : null;
    // 근무 칸은 이 저장 경로가 모르는 값이니, 이미 있던 값을 그대로 유지
    var workValue = payload.work != null ? payload.work : (existingRow ? existingRow[HEADERS.indexOf('근무')] : '');
    var row = [
      payload.date,
      payload.weight != null ? payload.weight : '',
      payload.waist != null ? payload.waist : '',
      payload.thigh != null ? payload.thigh : '',
      payload.hip != null ? payload.hip : '',
      !!payload.skipMeasure,
      payload.kcal || 0,
      payload.carb || 0,
      payload.prot || 0,
      payload.fat || 0,
      payload.sugar || 0,
      payload.burn || 0,
      payload.net || 0,
      payload.sleep != null ? payload.sleep : '',
      payload.bowel === true ? '배변함' : (payload.bowel === false ? '배변안함' : ''),
      (payload.tags || []).join(','),
      payload.note || '',
      JSON.stringify(payload.meals || {}),
      JSON.stringify(payload.exercises || []),
      JSON.stringify(payload.scheduleItems || []),
      workValue || ''
    ];
    if (rowIndex === -1) {
      rowIndex = sheet.getLastRow() + 1;
    }
    writeRow_(sheet, rowIndex, row);
    return jsonOutput_({ ok: true });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err) });
  }
}
