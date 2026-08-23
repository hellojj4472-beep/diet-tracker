// 다이어트표 - Google Sheets 연동용 Apps Script
// 이 파일은 참고/백업용입니다. 실제로는 구글시트의 확장프로그램 > Apps Script 편집기에 붙여넣어 사용하세요.

var SHEET_NAME = '기록';
var TOKEN = '여기에-원하는-비밀번호처럼-긴-문자열을-넣으세요'; // 예: diet-8x3kQ9zL2m

var HEADERS = ['날짜','체중','허리둘레','허벅지둘레','골반둘레','측정건너뜀','섭취칼로리','탄수화물','단백질','지방','단순당','운동소모칼로리','순섭취칼로리','수면시간','배변','특이사항태그','메모','식사상세JSON','운동상세JSON','일정'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
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

function rowToObject_(r) {
  return {
    date: formatDate_(r[0]), weight: r[1], waist: r[2], thigh: r[3], hip: r[4], skipMeasure: r[5],
    kcal: r[6], carb: r[7], prot: r[8], fat: r[9], sugar: r[10], burn: r[11], net: r[12],
    sleep: r[13], bowel: r[14], tags: r[15], note: r[16], schedule: r[19]
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

// 오늘 기록 저장/갱신: POST JSON body { token, date, weight, ... }
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    checkToken_(payload.token);
    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (formatDate_(data[i][0]) === payload.date) { rowIndex = i + 1; break; }
    }
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
      payload.schedule || ''
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
