// ============================================
// MAGAZYN / WYPOŻYCZALNIA - BACKEND v3.0
// Nowy model: Katalog + Przesunięcia
// ============================================

var CACHE_TTL = 60 * 5;
var SPREADSHEET_ID = '1bGjJ4NYfrdtKcqX2GIWIheo_KQQHdhH6keDurrZEHxM';

var SHEET_KATALOG = "Katalog";
var SHEET_OSOBY = "Osoby";
var SHEET_PRZESUNIECIA = "Przesunięcia";
var SHEET_INWENTARYZACJA = "Inwentaryzacja";
var SHEET_INV_DOSTAWA = "Inv_Dostawa";
var SHEET_INV_WYNIKI = "Inv_Wyniki";
var SHEET_INV_BRAKI = "Inv_Braki";
var DRIVE_FOLDER_NAME = "Magazyn_Zdjecia";

var COLS_KATALOG = {
  ID: 0,
  NAZWA_SYSTEMOWA: 1,
  NAZWA_WYSWIETLANA: 2,
  KATEGORIA: 3,
  SN: 4,
  STAN_POCZATKOWY: 5,
  AKTUALNIE_NA_STANIE: 6,
  FLAGA: 7,
  TAGI: 8,
  OSTATNIO_WIDZIANE: 9,
  OPIS: 10,
  PRZESUN: 11,
  DATA_PRZESUN: 12
};

var COLS_PRZES = {
  ID_OPERACJI: 0,
  DATA_WYDANIA: 1,
  OSOBA: 2,
  NAZWA_SYSTEMOWA: 3,
  SN: 4,
  ILOSC: 5,
  KATEGORIA: 6,
  STATUS: 7,
  ZDJECIE_WYDANIE_URL: 8,
  ZDJECIE_ZWROT_URL: 9,
  DATA_ZWROTU: 10,
  OPIS_USZKODZENIA: 11,
  OPERATOR: 12
};

function clearAllCache() {
  CacheService.getScriptCache().removeAll(["osoby","katalog","katalogGrouped"]);
  Logger.log("Cache wyczyszczony");
}

function doGet(e) {
  // Auto-version: po deploy curl ?action=setVersion&v=X&url=URL
  if (e && e.parameter && e.parameter.action === 'restoreSN') {
    // Przywraca brakujące/błędne SN wg mapy {KAT_ID: correctSN}
    var fixes = JSON.parse(e.parameter.data || '{}');
    var ks = getSheet(SHEET_KATALOG);
    var lr = ks.getLastRow();
    if (lr < 2) return ContentService.createTextOutput('{}').setMimeType(ContentService.MimeType.JSON);
    var ids = ks.getRange(2, COLS_KATALOG.ID + 1, lr - 1, 1).getValues();
    var snCol = COLS_KATALOG.SN + 1;
    var fixed = 0;
    var details = [];
    for (var ri = 0; ri < ids.length; ri++) {
      var katId = String(ids[ri][0]);
      if (fixes[katId] !== undefined) {
        var oldSn = String(ks.getRange(ri + 2, snCol).getValue() || '');
        ks.getRange(ri + 2, snCol).setValue(fixes[katId]);
        details.push({id: katId, row: ri + 2, old: oldSn, new_val: fixes[katId]});
        fixed++;
      }
    }
    CacheService.getScriptCache().remove('katalog');
    return ContentService.createTextOutput(JSON.stringify({fixed: fixed, details: details})).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'fixSN') {
    // Czyści śmieciowe SN: fragmenty tekstu z Excela, za krótkie, interpunkcję
    var ks = getSheet(SHEET_KATALOG);
    var lr = ks.getLastRow();
    if (lr < 2) return ContentService.createTextOutput('{}').setMimeType(ContentService.MimeType.JSON);
    var snCol = COLS_KATALOG.SN + 1;
    var vals = ks.getRange(2, snCol, lr - 1, 1).getValues();
    var fixed = 0;
    var details = [];
    // Znane śmieci z Excela (fragmenty opisów z sąsiednich kolumn)
    var garbageExact = ['PACK)', 'I', 'ÓWKA', 'inox', '2', 'zt', 'ZT', 'BG'];
    for (var ci = 0; ci < vals.length; ci++) {
      var v = String(vals[ci][0] || '').trim();
      if (!v) continue;
      var clean = v.replace(/^[;:.,]+/, '').replace(/[;:.,]+$/, '').trim();
      // Usuń jeśli: znany śmieć, za krótki (1 znak), sam interpunkcja, kończy się na ")"
      var isGarbage = false;
      if (garbageExact.indexOf(clean) >= 0) isGarbage = true;
      else if (clean.length <= 1) isGarbage = true;
      else if (/^[;:.,\-\/\s\(\)]+$/.test(clean)) isGarbage = true;
      else if (/\)$/.test(clean) && clean.length < 10) isGarbage = true;
      if (isGarbage) {
        ks.getRange(ci + 2, snCol).setValue('');
        details.push({row: ci + 2, old: v, new_val: ''});
        fixed++;
      } else if (clean !== v) {
        ks.getRange(ci + 2, snCol).setValue(clean);
        details.push({row: ci + 2, old: v, new_val: clean});
        fixed++;
      }
    }
    CacheService.getScriptCache().remove('katalog');
    return ContentService.createTextOutput(JSON.stringify({fixed: fixed, details: details})).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'writeCol') {
    // Zapisuje dane do wybranej kolumny w Katalogu
    // col = nazwa kolumny (OPIS, PRZESUN), data = JSON {KAT_ID: "wartość", ...}
    var colName = e.parameter.col || 'OPIS';
    var colIdx = COLS_KATALOG[colName];
    if (colIdx === undefined) return ContentService.createTextOutput('{"error":"bad col"}').setMimeType(ContentService.MimeType.JSON);
    var dataMap = JSON.parse(e.parameter.data || '{}');
    var ks = getSheet(SHEET_KATALOG);
    var lr = ks.getLastRow();
    if (lr < 2) return ContentService.createTextOutput('{}').setMimeType(ContentService.MimeType.JSON);
    var ids = ks.getRange(2, COLS_KATALOG.ID + 1, lr - 1, 1).getValues();
    var targetCol = colIdx + 1;
    var written = 0;
    for (var ri = 0; ri < ids.length; ri++) {
      var katId = String(ids[ri][0]);
      if (dataMap[katId] !== undefined) {
        ks.getRange(ri + 2, targetCol).setValue(dataMap[katId]);
        written++;
      }
    }
    CacheService.getScriptCache().remove('katalog');
    return ContentService.createTextOutput(JSON.stringify({written: written, col: colName})).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'deleteRows') {
    // Usuwa wiersze z Katalogu po KAT_ID — JSON array ["KAT02479","KAT02480",...]
    var ids = JSON.parse(e.parameter.ids || '[]');
    var ks = getSheet(SHEET_KATALOG);
    var lr = ks.getLastRow();
    if (lr < 2 || !ids.length) return ContentService.createTextOutput('{"deleted":0}').setMimeType(ContentService.MimeType.JSON);
    var allIds = ks.getRange(2, COLS_KATALOG.ID + 1, lr - 1, 1).getValues();
    var idSet = {};
    for (var di = 0; di < ids.length; di++) idSet[ids[di]] = true;
    var deleted = 0;
    for (var ri = allIds.length - 1; ri >= 0; ri--) {
      if (idSet[String(allIds[ri][0])]) {
        ks.deleteRow(ri + 2);
        deleted++;
      }
    }
    CacheService.getScriptCache().remove('katalog');
    return ContentService.createTextOutput(JSON.stringify({deleted: deleted})).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'dumpKatalog') {
    var ks = getSheet(SHEET_KATALOG);
    var lr = ks.getLastRow();
    if (lr < 2) return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
    var d = ks.getRange(2, 1, lr - 1, 13).getValues();
    var rows = [];
    for (var i = 0; i < d.length; i++) {
      rows.push({r: i+2, id: String(d[i][0]), kod: String(d[i][1]), nazwa: String(d[i][2]), kat: String(d[i][3]), sn: String(d[i][4]), sp: Number(d[i][5])||0, as: Number(d[i][6])||0, flaga: Number(d[i][7])||0, tagi: String(d[i][8]), ost: String(d[i][9]), opis: String(d[i][10]), przesun: String(d[i][11]), dp: String(d[i][12])});
    }
    return ContentService.createTextOutput(JSON.stringify(rows)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'dumpPrzesuniecia') {
    var ps = getSheet(SHEET_PRZESUNIECIA);
    var lr = ps.getLastRow();
    if (lr < 2) return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
    var d = ps.getRange(2, 1, lr - 1, 13).getValues();
    var rows = [];
    for (var i = 0; i < d.length; i++) {
      rows.push({r: i+2, idOp: String(d[i][0]), data: String(d[i][1]), osoba: String(d[i][2]), kod: String(d[i][3]), sn: String(d[i][4]), ilosc: Number(d[i][5])||0, kat: String(d[i][6]), status: String(d[i][7])});
    }
    return ContentService.createTextOutput(JSON.stringify(rows)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'backupSpreadsheet') {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var copy = ss.copy('BACKUP_' + new Date().toISOString().slice(0,10) + '_' + ss.getName());
    return ContentService.createTextOutput(JSON.stringify({ok: true, name: copy.getName(), url: copy.getUrl()})).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'setLokalizacje') {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_OSOBY);
    var last = sheet.getLastRow();
    if (last < 2) return ContentService.createTextOutput('{}').setMimeType(ContentService.MimeType.JSON);
    var data = sheet.getRange(2, 1, last - 1, 5).getValues();
    var map = JSON.parse(e.parameter.map || '{}');
    var count = 0;
    for (var i = 0; i < data.length; i++) {
      var imie = String(data[i][1]).trim();
      if (map[imie]) {
        sheet.getRange(i + 2, 4).setValue(map[imie]);
        count++;
      }
    }
    CacheService.getScriptCache().remove('osoby');
    return ContentService.createTextOutput(JSON.stringify({updated: count})).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'dumpOsoby') {
    var osoby = getOsoby();
    return ContentService.createTextOutput(JSON.stringify(osoby)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'fixName') {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(SHEET_PRZESUNIECIA);
    var last = sh.getLastRow();
    if (last < 2) return ContentService.createTextOutput('{}').setMimeType(ContentService.MimeType.JSON);
    var data = sh.getRange(2, COLS_PRZES.OSOBA + 1, last - 1, 1).getValues();
    var count = 0;
    var oldName = e.parameter.oldName || '';
    var newName = e.parameter.newName || '';
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === oldName) {
        sh.getRange(i + 2, COLS_PRZES.OSOBA + 1).setValue(newName);
        count++;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({fixed: count, oldName: oldName, newName: newName})).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'searchKatalog') {
    var q = (e.parameter.q || '').toLowerCase();
    var ks = getSheet(SHEET_KATALOG);
    var lr = ks.getLastRow();
    var results = [];
    if (lr >= 2) {
      var d = ks.getRange(2, 1, lr - 1, 13).getValues();
      for (var i = 0; i < d.length; i++) {
        var row = d[i].map(function(c){return String(c)}).join('|').toLowerCase();
        if (row.indexOf(q) >= 0) results.push({row: i+2, id: d[i][0], kod: d[i][1], nazwa: d[i][2], kat: d[i][3], sn: d[i][4], stanPocz: d[i][5], aktStan: d[i][6], opis: d[i][10], przesun: d[i][11]});
      }
    }
    return ContentService.createTextOutput(JSON.stringify(results)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'getRaportState') {
    var items = _getRaportItems();
    var matrix = _getRaportMatrix();
    var loks = _getUniqueLokalizacje();
    var email = getRaportEmail();
    var trigOn = false;
    var triggers = ScriptApp.getProjectTriggers();
    for (var t = 0; t < triggers.length; t++) {
      if (triggers[t].getHandlerFunction() === 'wyslijSzablonyEmail') { trigOn = true; break; }
    }
    return ContentService.createTextOutput(JSON.stringify({email: email, items: items, matrix: matrix, lokalizacje: loks, triggerActive: trigOn})).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'testSendRaport') {
    ScriptApp.newTrigger('wyslijTestRaport').timeBased().after(5000).create();
    return ContentService.createTextOutput(JSON.stringify({success: true, msg: 'trigger za 5s'})).setMimeType(ContentService.MimeType.JSON);
  }
    if (e && e.parameter && e.parameter.action === 'installTrigger') {
    var r = installSzablonyTrigger();
    return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
  }
    if (e && e.parameter && e.parameter.action === 'setVersion') {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('APP_VERSION', e.parameter.v || '0');
    props.setProperty('APP_URL', e.parameter.url || '');
    return ContentService.createTextOutput(JSON.stringify({ok: true, version: e.parameter.v}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'renameLoks') {
    var map = JSON.parse(e.parameter.data || '{}'); // {oldName: newName, ...}
    var keys = Object.keys(map);
    var results = [];
    // 1. Update Osoby sheet
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_OSOBY);
    var lr = sheet.getLastRow();
    if (lr >= 2) {
      var hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var lokCol = -1;
      for (var c = 0; c < hdr.length; c++) { if (String(hdr[c]).toLowerCase() === 'lokalizacja') { lokCol = c; break; } }
      if (lokCol >= 0) {
        var vals = sheet.getRange(2, lokCol + 1, lr - 1, 1).getValues();
        var updated = 0;
        for (var i = 0; i < vals.length; i++) {
          var v = String(vals[i][0] || '').trim();
          if (map[v]) { sheet.getRange(i + 2, lokCol + 1).setValue(map[v]); updated++; }
        }
        results.push({step: 'osoby', updated: updated});
      }
    }
    // 2. Rename matrix keys
    var matrix = _getRaportMatrix();
    for (var k = 0; k < keys.length; k++) {
      var oldK = keys[k];
      var newK = map[oldK];
      if (matrix[oldK]) {
        if (!matrix[newK]) matrix[newK] = {};
        var items = Object.keys(matrix[oldK]);
        for (var m = 0; m < items.length; m++) {
          matrix[newK][items[m]] = matrix[oldK][items[m]];
        }
        delete matrix[oldK];
      }
    }
    // 3. Remove empty entries
    var allKeys = Object.keys(matrix);
    for (var e2 = 0; e2 < allKeys.length; e2++) {
      if (!matrix[allKeys[e2]] || Object.keys(matrix[allKeys[e2]]).length === 0) {
        delete matrix[allKeys[e2]];
      }
    }
    _setRaportMatrix(matrix);
    results.push({step: 'matrix', keys: Object.keys(matrix)});
    // 4. Update custom loks
    var custom = _getCustomLokalizacje();
    var newCustom = custom.map(function(l) { return map[l] || l; });
    _getProps().setProperty('customLoks', JSON.stringify(newCustom));
    CacheService.getScriptCache().remove('osoby');
    return ContentService.createTextOutput(JSON.stringify({success: true, results: results})).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'matrixBatch') {
    var data = JSON.parse(e.parameter.data || '{}');
    var results = [];
    // Add items
    if (data.addItems) {
      for (var ai = 0; ai < data.addItems.length; ai++) {
        var it = data.addItems[ai];
        results.push({op: 'addItem', nazwa: it.nazwa, r: addMatrixItem(it.nazwa, it.ilosc || 1, it.jm || 'szt')});
      }
    }
    // Set cells
    if (data.setCells) {
      for (var ci = 0; ci < data.setCells.length; ci++) {
        var c = data.setCells[ci];
        results.push({op: 'setCell', nazwa: c.nazwa, lok: c.lokalizacja, r: setMatrixCell(c.nazwa, c.lokalizacja, c.checked !== false, c.ilosc || 1)});
      }
    }
    // Remove items
    if (data.removeItems) {
      for (var ri = 0; ri < data.removeItems.length; ri++) {
        results.push({op: 'removeItem', nazwa: data.removeItems[ri], r: removeMatrixItem(data.removeItems[ri])});
      }
    }
    // Set quantities
    if (data.setQty) {
      for (var qi = 0; qi < data.setQty.length; qi++) {
        var q = data.setQty[qi];
        results.push({op: 'setQty', nazwa: q.nazwa, lok: q.lokalizacja, r: setMatrixQty(q.nazwa, q.lokalizacja, q.ilosc)});
      }
    }
    return ContentService.createTextOutput(JSON.stringify({success: true, results: results})).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'previewZamowienie') {
    var data = sendZamowienieAkceptacja();
    return HtmlService.createHtmlOutput(data.html);
  }
  if (e && e.parameter && e.parameter.action === 'sendZamowienie') {
    var r = sendZamowienieEmail();
    return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'testZamowienie') {
    var data = sendZamowienieAkceptacja();
    var r = sendEmail('stechnij.kamil@gmail.com', data.subject, data.plain, data.html, CC_MAGAZYN);
    return ContentService.createTextOutput(JSON.stringify({ success: true, sent: r })).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'formatPro') {
    var r = formatSpreadsheetPro();
    return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'dumpExtSheet') {
    var ssId = e.parameter.ssId || '';
    var gid = Number(e.parameter.gid || 0);
    var ss2 = SpreadsheetApp.openById(ssId);
    var sheets = ss2.getSheets();
    var sh2 = null;
    for (var si = 0; si < sheets.length; si++) {
      if (sheets[si].getSheetId() === gid) { sh2 = sheets[si]; break; }
    }
    if (!sh2) sh2 = sheets[0];
    var lr2 = sh2.getLastRow();
    var lc2 = sh2.getLastColumn();
    var d2 = lr2 >= 1 ? sh2.getRange(1, 1, lr2, lc2).getValues() : [];
    var rows2 = [];
    for (var ri = 0; ri < d2.length; ri++) {
      rows2.push(d2[ri].map(function(c) { return String(c); }));
    }
    return ContentService.createTextOutput(JSON.stringify({name: sh2.getName(), rows: rows2})).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'listRevisions') {
    var ssId = e.parameter.ssId || SPREADSHEET_ID;
    var token = ScriptApp.getOAuthToken();
    var url = 'https://www.googleapis.com/drive/v3/files/' + ssId + '/revisions?fields=revisions(id,modifiedTime,lastModifyingUser/displayName)&pageSize=200';
    var resp = UrlFetchApp.fetch(url, {headers: {'Authorization': 'Bearer ' + token}});
    return ContentService.createTextOutput(resp.getContentText()).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'exportRevision') {
    var ssId = e.parameter.ssId || SPREADSHEET_ID;
    var revId = e.parameter.revId || '';
    var gid = e.parameter.gid || '0';
    var token = ScriptApp.getOAuthToken();
    var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=csv&gid=' + gid + '&revision=' + revId;
    var resp = UrlFetchApp.fetch(url, {headers: {'Authorization': 'Bearer ' + token}});
    return ContentService.createTextOutput(resp.getContentText()).setMimeType(ContentService.MimeType.TEXT);
  }
  if (e && e.parameter && e.parameter.action === 'searchExtSheet') {
    var ssId = e.parameter.ssId || '';
    var gid = Number(e.parameter.gid || 0);
    var q = (e.parameter.q || '').toLowerCase();
    var ss2 = SpreadsheetApp.openById(ssId);
    var sheets = ss2.getSheets();
    var sh2 = null;
    for (var si = 0; si < sheets.length; si++) {
      if (sheets[si].getSheetId() === gid) { sh2 = sheets[si]; break; }
    }
    if (!sh2) sh2 = sheets[0];
    var lr2 = sh2.getLastRow();
    var lc2 = sh2.getLastColumn();
    var d2 = lr2 >= 1 ? sh2.getRange(1, 1, lr2, lc2).getValues() : [];
    var rows2 = [];
    for (var ri = 0; ri < d2.length; ri++) {
      var line = d2[ri].map(function(c) { return String(c); }).join('|').toLowerCase();
      if (ri === 0 || line.indexOf(q) >= 0) {
        rows2.push({r: ri, d: d2[ri].map(function(c) { return String(c); })});
      }
    }
    return ContentService.createTextOutput(JSON.stringify({name: sh2.getName(), total: lr2, matches: rows2})).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'listExtSheets') {
    var ssId = e.parameter.ssId || '';
    var ss2 = SpreadsheetApp.openById(ssId);
    var sheets = ss2.getSheets();
    var result = [];
    for (var si = 0; si < sheets.length; si++) {
      result.push({name: sheets[si].getName(), gid: sheets[si].getSheetId(), rows: sheets[si].getLastRow()});
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'importExcelHistory') {
    var r = importExcelHistory();
    return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'importMissing') {
    var r = importMissingRecords();
    return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
  }
    var tpl = HtmlService.createTemplateFromFile('index');
  return tpl.evaluate()
    .setTitle('Magazyn Wypożyczalnia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var batchId = body.batchId || '';
    // Dedup check jest WEWNĄTRZ każdej funkcji (po lock.waitLock),
    // NIE tutaj — bo doPost + funkcja = 2x _isDuplicateBatch = zawsze "duplikat"
    var result;
    if (action === 'wydajBatch') {
      result = wydajBatch(body.idOsoby, body.items, body.operator, batchId);
    } else if (action === 'zwrocBatch') {
      result = zwrocBatch(body.ids || [], body.photoDataMap || {}, body.qtyMap || {}, body.operator, batchId);
    } else if (action === 'przeniesNaOsobe') {
      result = przeniesNaOsobe(body.ids || [], body.nowaOsoba, body.operator, batchId);
    } else if (action === 'przeniesUszkodzone') {
      result = przeniesUszkodzone(body.idOp, body.nowaOsoba, body.operator, batchId);
    } else {
      result = { error: 'unknown action' };
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function _isDuplicateBatch(batchId) {
  var cache = CacheService.getScriptCache();
  var key = 'batch_' + batchId;
  if (cache.get(key)) return true;
  cache.put(key, '1', 3600);
  return false;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAppVersion() {
  var props = PropertiesService.getScriptProperties();
  return {
    version: props.getProperty('APP_VERSION') || '0',
    url: props.getProperty('APP_URL') || ''
  };
}

function getSheet(name) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
  if (!sheet) throw new Error('Brak arkusza: ' + name);
  return sheet;
}

function generateId(prefix) {
  return prefix + Date.now() + Math.random().toString(36).substr(2, 5);
}

function extractSN(raw) {
  if (!raw) return '';
  var str = String(raw).trim();
  // Odrzuć same znaki interpunkcyjne (";", ":", "s/n;", itp.)
  if (/^[;:.,\-\/\s]+$/.test(str)) return '';
  if (/^(?:s\/n|sn)[.:;\s]*$/i.test(str)) return '';
  // Obsługa S/N; SPACE SN (np. "S/N; J233111 2022R, WALIZKA")
  var match = str.match(/(?:s\/n|sn)[.:;\s]+([A-Za-z0-9][\w\-]*)/i);
  if (match && match[1].length > 1) str = match[1];
  // Usuń wiodące i końcowe interpunkcje
  str = str.replace(/^[;:.,]+/, '').replace(/[;:.,]+$/, '');
  return str;
}

function simplifyToolName(fullName) {
  if (!fullName) return '';
  var name = String(fullName);

  // Usuń kod na początku (np. SYN/6/52, ABC/12/345)
  name = name.replace(/^[A-Z]{2,}[\/-]\d+[\/-]\d+\s*/i, '');
  // Usuń angielskie nazwy i inne treści w nawiasach
  name = name.replace(/\s*\([^)]*\)\s*/g, ' ');

  var brands = [
    'STANLEY', 'MAKITA', 'BOSCH', 'YATO', 'DEWALT', 'MILWAUKEE', 'HILTI',
    'METABO', 'FESTOOL', 'STIHL', 'HUSQVARNA', 'KARCHER', 'RYOBI', 'EINHELL',
    'PARKSIDE', 'GRAPHITE', 'DEDRA', 'TOPEX', 'VOREL', 'STHOR', 'NEO',
    'PROLINE', 'HOGERT', 'BAHCO', 'KNIPEX', 'WIHA', 'WERA', 'IRWIN',
    'TAJIMA', 'STABILA', 'WURTH', 'FISCHER', 'RAWLPLUG', 'WOLFCRAFT',
    'HIKOKI', 'HITACHI', 'KRESS', 'FLEX', 'FEIN', 'AEG', 'RIDGID', 'RUBI',
    'BERNER', 'ERBAUER', 'TITAN', 'TOTAL', 'INGCO', 'TOYA', 'EXTOL', 'KLINGSPOR'
  ];

  var colors = [
    'ZOLTA', 'ŻÓŁTA', 'ZÓŁTA', 'CZARNA', 'CZERWONA', 'ZIELONA', 'NIEBIESKA',
    'BIALA', 'BIAŁA', 'SZARA', 'POMARANCZOWA', 'FIOLETOWA', 'BRAZOWA', 'BRĄZOWA',
    'SREBRNA', 'ZLOTA', 'ZŁOTA'
  ];

  var words = name.trim().split(/\s+/);
  var result = [];

  for (var i = 0; i < words.length; i++) {
    var w = words[i];
    var wUp = w.toUpperCase();

    // Pomiń marki
    if (brands.indexOf(wUp) !== -1) continue;

    // Pomiń numery modeli (30-457, 12345+)
    if (/^\d{2,}-\d+/.test(w)) continue;
    if (/^\d{5,}$/.test(w)) continue;

    // Pomiń kolory
    if (colors.indexOf(wUp) !== -1) continue;

    // Pomiń wymiary NxN (300X175, 150x50)
    if (/^\d+[Xx]\d+/.test(w)) continue;

    // Zachowaj proste miary (8M, 5M, 10MM, 2.5M) ale pomiń długie kody (300MM+)
    if (/^\d{3,}[Mm]{1,2}$/.test(w)) continue;

    // Modele alfanumeryczne (DS18DE, DHP453, GBH2-26, WR18DBDL2)
    if (/^[A-Z]{1,4}\d{2,}[A-Z]*\d*$/i.test(w)) continue;
    if (/^[A-Z]{2,}\d+-\w+$/i.test(w)) continue;

    // Samotne kreski/myślniki
    if (/^[-–—]+$/.test(w)) continue;

    result.push(w);
  }

  return result.join(' ').trim();
}

function formatDate(d) {
  if (!d) return '';
  try {
    var date = new Date(d);
    if (isNaN(date.getTime())) return '';
    var dd = ('0' + date.getDate()).slice(-2);
    var mm = ('0' + (date.getMonth() + 1)).slice(-2);
    var yy = date.getFullYear();
    var hh = ('0' + date.getHours()).slice(-2);
    var mi = ('0' + date.getMinutes()).slice(-2);
    return dd + '.' + mm + '.' + yy + ' ' + hh + ':' + mi;
  } catch (e) {
    return String(d);
  }
}

// ============================================
// INITIAL DATA
// ============================================

function getInitialData() {
  return {
    osoby: getOsoby(),
    katalog: getKatalog(),
    katalogGrouped: getKatalogGrouped()
  };
}

// ============================================
// SCHEMA — auto-setup arkuszy
// ============================================

var SCHEMA = {
  Katalog: ['ID', 'Nazwa_Systemowa', 'Nazwa_Wyswietlana', 'Kategoria', 'SN', 'Stan_Poczatkowy', 'Aktualnie_Na_Stanie', 'Flaga', 'Tagi', 'Ostatnio_Widziane'],
  Osoby: ['ID', 'Imie', 'Telefon'],
  'Przesunięcia': ['ID_Operacji', 'Data_Wydania', 'Osoba', 'Nazwa_Systemowa', 'SN', 'Ilosc', 'Kategoria', 'Status', 'Zdjecie_Wydanie_URL', 'Zdjecie_Zwrot_URL', 'Data_Zwrotu', 'Opis_Uszkodzenia', 'Operator'],
  Inwentaryzacja: ['Timestamp', 'Sesja', 'DriveURL', 'DriveFileID', 'Items', 'Texts', 'Opis'],
  Inv_Dostawa: ['Timestamp', 'Sesja', 'Osoba', 'Typ', 'Opis', 'DriveFileID', 'Items_JSON'],
  Inv_Wyniki: ['Timestamp', 'Sesja', 'Osoba', 'Nazwa_Systemowa', 'Nazwa_Wyswietlana', 'SN', 'Ilosc', 'AI_Nazwa'],
  Inv_Braki: ['Timestamp', 'Sesja', 'Osoba', 'Nazwa_AI', 'Ilosc']
};

