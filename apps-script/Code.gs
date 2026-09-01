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

// 체중 칸 하나만 업데이트 (그 날의 다른 기록은 건드리지 않음)
function setWeightOnly_(sheet, payload) {
  var data = sheet.getDataRange().getValues();
  var rowIndex = findRowIndex_(data, payload.date);
  if (rowIndex === -1) {
    rowIndex = sheet.getLastRow() + 1;
    sheet.getRange(rowIndex, 1).setNumberFormat('@').setValue(payload.date);
  }
  var weightCol = HEADERS.indexOf('체중') + 1;
  sheet.getRange(rowIndex, weightCol).setValue(payload.weight != null ? payload.weight : '');
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

// 지난 날짜 기록을 통째로 삭제 (칼로리를 덜 입력했거나 잘못 저장된 날 등을 지울 때 사용)
function deleteDay_(sheet, payload) {
  var data = sheet.getDataRange().getValues();
  var rowIndex = findRowIndex_(data, payload.date);
  if (rowIndex === -1) return { ok: false, error: 'row not found' };
  sheet.deleteRow(rowIndex);
  return { ok: true };
}

// 지난 날짜 팝업에서 개별 항목(허리/허벅지/골반/수면/배변/특이사항/메모/식사/운동) 하나만 업데이트
var FIELD_TO_HEADER = {
  waist: '허리둘레', thigh: '허벅지둘레', hip: '골반둘레',
  sleep: '수면시간', bowel: '배변', tags: '특이사항태그', note: '메모',
  meals: '식사상세JSON', exercises: '운동상세JSON'
};
function setFieldOnly_(sheet, payload) {
  var header = FIELD_TO_HEADER[payload.field];
  if (!header) return { ok: false, error: 'unknown field' };
  var data = sheet.getDataRange().getValues();
  var rowIndex = findRowIndex_(data, payload.date);
  if (rowIndex === -1) {
    rowIndex = sheet.getLastRow() + 1;
    sheet.getRange(rowIndex, 1).setNumberFormat('@').setValue(payload.date);
  }
  var col = HEADERS.indexOf(header) + 1;
  var value = payload.value;
  if (payload.field === 'meals') value = JSON.stringify(value || {});
  else if (payload.field === 'exercises') value = JSON.stringify(value || []);
  else if (payload.field === 'bowel') value = value === true ? '배변함' : (value === false ? '배변안함' : '');
  else if (payload.field === 'tags') value = (value || []).join(',');
  else value = value != null ? value : '';
  sheet.getRange(rowIndex, col).setValue(value);
  return { ok: true };
}

// 지난 날짜 팝업에서 식사/운동을 고치면 칼로리 합계 칸도 같이 갱신
function setTotalsOnly_(sheet, payload) {
  var data = sheet.getDataRange().getValues();
  var rowIndex = findRowIndex_(data, payload.date);
  if (rowIndex === -1) {
    rowIndex = sheet.getLastRow() + 1;
    sheet.getRange(rowIndex, 1).setNumberFormat('@').setValue(payload.date);
  }
  var t = payload.totals || {};
  var fields = [
    ['섭취칼로리', t.kcal], ['탄수화물', t.carb], ['단백질', t.prot], ['지방', t.fat],
    ['단순당', t.sugar], ['운동소모칼로리', t.burn], ['순섭취칼로리', t.net]
  ];
  fields.forEach(function (pair) {
    var col = HEADERS.indexOf(pair[0]) + 1;
    sheet.getRange(rowIndex, col).setValue(pair[1] || 0);
  });
  return { ok: true };
}

// 오늘 기록 저장/갱신: POST JSON body { token, date, weight, ... }
function doPost(e) {
  // 여러 요청이 동시에 들어오면 같은 날짜에 중복 행이 생길 수 있어서, 한 번에 하나씩만 처리되도록 잠금
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return jsonOutput_({ ok: false, error: 'busy, try again' });
  }
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
    if (payload.action === 'weight') {
      return jsonOutput_(setWeightOnly_(sheet, payload));
    }
    if (payload.action === 'field') {
      return jsonOutput_(setFieldOnly_(sheet, payload));
    }
    if (payload.action === 'totals') {
      return jsonOutput_(setTotalsOnly_(sheet, payload));
    }
    if (payload.action === 'deleteDay') {
      return jsonOutput_(deleteDay_(sheet, payload));
    }
    var data = sheet.getDataRange().getValues();
    var rowIndex = findRowIndex_(data, payload.date);
    var existingRow = rowIndex !== -1 ? data[rowIndex - 1] : null;
    // 근무/일정 칸은 각각 실시간으로 따로 저장되니, 이 저장 경로에서 값을 안 보내주면 이미 있던 값을 그대로 유지
    var workValue = payload.work != null ? payload.work : (existingRow ? existingRow[HEADERS.indexOf('근무')] : '');
    var scheduleValue = payload.scheduleItems !== undefined
      ? JSON.stringify(payload.scheduleItems)
      : (existingRow ? existingRow[HEADERS.indexOf('일정')] : JSON.stringify([]));
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
      scheduleValue,
      workValue || ''
    ];
    if (rowIndex === -1) {
      rowIndex = sheet.getLastRow() + 1;
    }
    writeRow_(sheet, rowIndex, row);
    return jsonOutput_({ ok: true });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
