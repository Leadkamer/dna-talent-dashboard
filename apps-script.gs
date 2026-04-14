// DNA Talent - Opvolging sync endpoint
// Deployen als Web App (Execute as: Me, Access: Anyone) op dezelfde Sheet.

var SHEET_ID = '1D2OOrxtTIt0tRAIJRR5Tkjh_FBnbSN8lu7BM-xyfz5U';
var SHEET_NAME = 'Opvolging';

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'list';
    var sheet = getOrCreateSheet();

    if (action === 'list') {
      var data = sheet.getDataRange().getValues();
      var result = {};
      for (var i = 1; i < data.length; i++) {
        if (data[i][0]) result[String(data[i][0])] = String(data[i][1] || '');
      }
      return jsonResponse({ok: true, actions: result});
    }

    if (action === 'set') {
      var key = e.parameter.key || '';
      var value = e.parameter.value || '';
      if (!key) return jsonResponse({ok: false, error: 'missing key'});
      upsert(sheet, key, value);
      return jsonResponse({ok: true});
    }

    return jsonResponse({ok: false, error: 'unknown action'});
  } catch (err) {
    return jsonResponse({ok: false, error: String(err)});
  }
}

function upsert(sheet, key, value) {
  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(key)) { rowIdx = i + 1; break; }
  }
  var now = new Date();
  if (!value) {
    if (rowIdx > 0) sheet.deleteRow(rowIdx);
    return;
  }
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 2).setValue(value);
    sheet.getRange(rowIdx, 3).setValue(now);
  } else {
    sheet.appendRow([key, value, now]);
  }
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Key', 'Actie', 'Laatst gewijzigd']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:C1').setFontWeight('bold');
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
