// ===== 천체관측동아리 방과후 불가능시간 조사 =====

var SHEET_NAME = '응답';
var ADMIN_PW = 'cos2023';

function getSS_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SS_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) {}
  }
  var ss = SpreadsheetApp.create('천체관측동아리 방과후 시간조사 응답');
  props.setProperty('SS_ID', ss.getId());
  return ss;
}

function getSheet_() {
  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['학번', '이름', '제출시각', '불가능시간', '설문답변']);
    var sheets = ss.getSheets();
    if (sheets.length > 1) {
      sheets.forEach(function (s) {
        if (s.getName() !== SHEET_NAME && ss.getSheets().length > 1) {
          try { ss.deleteSheet(s); } catch (e) {}
        }
      });
    }
  }
  return sheet;
}

function norm_(s) {
  return String(s == null ? '' : s).trim().normalize('NFC');
}

function findRow_(id, name) {
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var nid = norm_(id), nname = norm_(name);
  for (var i = 1; i < data.length; i++) {
    if (norm_(data[i][0]) === nid && norm_(data[i][1]) === nname) return i + 1;
  }
  return -1;
}

function submit_(id, name, slots, survey) {
  id = norm_(id); name = norm_(name);
  if (!id || !name) return { error: '학번/이름을 입력해주세요.' };
  if (survey !== '1' && survey !== '2' && survey !== '3') return { error: '설문에 답변해주세요.' };
  var sheet = getSheet_();
  var row = findRow_(id, name);
  var rowData = [id, name, new Date(), slots || '', survey];
  if (row === -1) {
    sheet.appendRow(rowData);
  } else {
    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
  }
  return { ok: true };
}

function getMy_(id, name) {
  var row = findRow_(id, name);
  if (row === -1) return { found: false };
  var sheet = getSheet_();
  var data = sheet.getRange(row, 1, 1, 5).getValues()[0];
  return { found: true, slots: String(data[3] || ''), survey: String(data[4] || '') };
}

function getStats_(pw) {
  if (String(pw) !== ADMIN_PW) return { error: '비밀번호가 일치하지 않습니다.' };
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[1]) continue;
    list.push({
      id: String(row[0]),
      name: String(row[1]),
      time: row[2] ? new Date(row[2]).toISOString() : '',
      slots: String(row[3] || '').split(',').filter(function (x) { return x; }),
      survey: String(row[4] || '')
    });
  }
  return { list: list };
}

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action;
  var callback = params.callback;

  if (!action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('천체관측동아리 방과후 시간조사')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  var data;
  try {
    if (action === 'submit') data = submit_(params.id, params.name, params.slots, params.survey);
    else if (action === 'getMy') data = getMy_(params.id, params.name);
    else if (action === 'getStats') data = getStats_(params.pw);
    else data = { error: 'unknown action' };
  } catch (err) {
    data = { error: err.toString() };
  }

  var json = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
