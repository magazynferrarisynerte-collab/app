// ============================================
// MAGAZYN / WYPOŻYCZALNIA - BACKEND v2.3 FIXED
// ============================================

const CACHE_TTL = 60 * 5;
const SPREADSHEET_ID = '1OCm_8VZ1Q-z1sXGsthJKUgHS_H13rdXa9uzGKdNrc6A';

const SHEET_NARZEDZIA = "Narzędzia";
const SHEET_OSOBY = "Osoby";
const SHEET_WYPOZYCZENIA = "Wypożyczenia";
const SHEET_USZKODZONE = "Uszkodzone";

const COLS_NARZ = {
  KOD: 0,
  NAZWA: 1,
  OPIS: 2,
  ILOSC: 3,
  JEDNOSTKA: 4,
  KATEGORIA: 5
};

// ============================================
// SYSTEM
// ============================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Magazyn Wypożyczalnia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet(name) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
  if (!sheet) throw new Error('Brak arkusza: ' + name);
  return sheet;
}

function generateId(prefix) {
  return prefix + Date.now() + Math.random().toString(36).substr(2, 5);
}

// Bezpieczne formatowanie daty do stringa
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
// LOG
// ============================================

function getLog() {
  var sheet = getSheet(SHEET_WYPOZYCZENIA);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  return data.map(function (r) {
    return {
      idWyp: String(r[0]),
      kod: String(r[1] || ''),
      nazwa: String(r[2] || ''),
      opis: String(r[3] || ''),
      idOsoby: String(r[4] || ''),
      imie: String(r[5] || ''),
      telefon: String(r[6] || ''),
      dataWyp: formatDate(r[7]),
      dataZwrotu: formatDate(r[8]),
      ilosc: Number(r[9]) || 1
    };
  }).sort(function (a, b) {
    var aDate = a.dataZwrotu || a.dataWyp || '';
    var bDate = b.dataZwrotu || b.dataWyp || '';
    return bDate.localeCompare(aDate);
  });
}

function getInitialData() {
  return {
    osoby: getOsoby(),
    narzedzia: getNarzedzia(),
    uszkodzone: getUszkodzone()
  };
}

// ============================================
// OSOBY
// ============================================

function getOsoby() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("osoby");
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* ignore */ }
  }

  var sheet = getSheet(SHEET_OSOBY);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  var osoby = data.map(function (r) {
    return {
      id: String(r[0]),
      imie: String(r[1] || ''),
      telefon: String(r[2] || '')
    };
  }).sort(function (a, b) {
    return a.imie.localeCompare(b.imie);
  });

  try {
    cache.put("osoby", JSON.stringify(osoby), CACHE_TTL);
  } catch (e) { /* za duże - pomijamy cache */ }
  return osoby;
}

function addOsoba(imie, telefon) {
  var sheet = getSheet(SHEET_OSOBY);
  var id = generateId('OS');
  sheet.appendRow([id, imie.trim(), telefon.trim()]);
  CacheService.getScriptCache().remove("osoby");
  return { success: true, id: id };
}

// ============================================
// NARZĘDZIA
// ============================================

function getNarzedzia() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("narzedzia");
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { }
  }

  var sheet = getSheet(SHEET_NARZEDZIA);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  var narzedzia = data.map(function (r) {
    return {
      kod: String(r[COLS_NARZ.KOD]),
      nazwa: String(r[COLS_NARZ.NAZWA]),
      opis: String(r[COLS_NARZ.OPIS] || ''),
      ilosc: Number(r[COLS_NARZ.ILOSC]) || 0,
      jednostka: String(r[COLS_NARZ.JEDNOSTKA] || 'szt.'),
      kategoria: String(r[COLS_NARZ.KATEGORIA] || '')
    };
  }).sort(function (a, b) {
    return a.nazwa.localeCompare(b.nazwa);
  });

  try {
    cache.put("narzedzia", JSON.stringify(narzedzia), CACHE_TTL);
  } catch (e) { }
  return narzedzia;
}

function addNarzedzie(kod, nazwa, ilosc, kategoria) {
  var sheet = getSheet(SHEET_NARZEDZIA);
  var finalKod = kod && kod.trim() ? kod.trim() : generateId('NZ');
  sheet.appendRow([finalKod, nazwa.trim(), '', Number(ilosc) || 1, 'szt.', kategoria || '']);
  CacheService.getScriptCache().remove("narzedzia");
  return { success: true, kod: finalKod };
}

// ============================================
// WYPOŻYCZANIE (BEZ KONFLIKTÓW)
// ============================================

function wypozyczBatch(idOsoby, items) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var narzSheet = ss.getSheetByName(SHEET_NARZEDZIA);
    var wypSheet = ss.getSheetByName(SHEET_WYPOZYCZENIA);
    var osoby = getOsoby();

    var osoba = null;
    for (var o = 0; o < osoby.length; o++) {
      if (osoby[o].id === idOsoby) { osoba = osoby[o]; break; }
    }
    if (!osoba) return { success: false, error: 'Nie znaleziono osoby' };

    var narzData = narzSheet.getDataRange().getValues();

    for (var idx = 0; idx < items.length; idx++) {
      var item = items[idx];
      var kod = item.kod;
      var qty = Number(item.qty) || 1;

      for (var i = 1; i < narzData.length; i++) {
        if (String(narzData[i][COLS_NARZ.KOD]) === kod) {
          var stan = Number(narzData[i][COLS_NARZ.ILOSC]) || 0;
          if (stan < qty) continue;

          narzSheet.getRange(i + 1, COLS_NARZ.ILOSC + 1).setValue(stan - qty);
          narzData[i][COLS_NARZ.ILOSC] = stan - qty;

          wypSheet.appendRow([
            generateId('W'),              // A: idWyp
            kod,                          // B: kod
            narzData[i][COLS_NARZ.NAZWA], // C: nazwa
            narzData[i][COLS_NARZ.OPIS],  // D: opis
            idOsoby,                      // E: idOsoby
            osoba.imie,                   // F: imie
            osoba.telefon,                // G: telefon
            new Date(),                   // H: dataWyp
            "",                           // I: dataZwrotu
            qty                           // J: ilosc
          ]);

          break;
        }
      }
    }

    CacheService.getScriptCache().remove("narzedzia");
    return { success: true };

  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

// ============================================
// ZWROT
// ============================================

function oddajNarzedzie(idWyp) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_WYPOZYCZENIA);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(idWyp)) {
        if (data[i][8] && String(data[i][8]).length > 0) {
          return { success: false, error: 'Już zwrócone' };
        }
        sheet.getRange(i + 1, 9).setValue(new Date());   // kolumna I (dataZwrotu)
        var kod = data[i][1];
        var ilosc = Number(data[i][9]) || 1;              // kolumna J (ilosc)
        zwiekszStanNarzedzia(kod, ilosc);
        CacheService.getScriptCache().remove("narzedzia");
        return { success: true };
      }
    }
    return { success: false, error: 'Nie znaleziono wypożyczenia' };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}
function oddajBatch(ids) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_WYPOZYCZENIA);
    var data = sheet.getDataRange().getValues();
    var count = 0;

    for (var idx = 0; idx < ids.length; idx++) {
      var idWyp = String(ids[idx]);
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === idWyp) {
          if (data[i][8] && String(data[i][8]).length > 0) continue;
          sheet.getRange(i + 1, 9).setValue(new Date());
          var kod = data[i][1];
          var ilosc = Number(data[i][9]) || 1;
          zwiekszStanNarzedzia(kod, ilosc);
          data[i][8] = new Date();
          count++;
          break;
        }
      }
    }

    CacheService.getScriptCache().remove("narzedzia");
    return { success: true, count: count };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}
function mergeDuplicateNarzedzia() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_NARZEDZIA);
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, merged: 0 };

    var map = {};
    var mergedRows = [];
    var toDelete = [];

    for (var i = 1; i < data.length; i++) {
      var kod = String(data[i][COLS_NARZ.KOD]).trim();
      var nazwa = String(data[i][COLS_NARZ.NAZWA]).trim();
      var key = kod ? kod.toLowerCase() : nazwa.toLowerCase();

      if (map[key] !== undefined) {
        var targetIdx = map[key];
        mergedRows[targetIdx].ilosc += Number(data[i][COLS_NARZ.ILOSC]) || 0;
        if (!mergedRows[targetIdx].opis && data[i][COLS_NARZ.OPIS]) {
          mergedRows[targetIdx].opis = String(data[i][COLS_NARZ.OPIS]);
        }
        if (!mergedRows[targetIdx].kategoria && data[i][COLS_NARZ.KATEGORIA]) {
          mergedRows[targetIdx].kategoria = String(data[i][COLS_NARZ.KATEGORIA]);
        }
        toDelete.push(i + 1); // numer wiersza (1-based)
      } else {
        map[key] = i;
        mergedRows[i] = {
          row: i + 1,
          ilosc: Number(data[i][COLS_NARZ.ILOSC]) || 0,
          opis: String(data[i][COLS_NARZ.OPIS] || ''),
          kategoria: String(data[i][COLS_NARZ.KATEGORIA] || '')
        };
      }
    }

    // Aktualizuj ilości w wierszach docelowych
    for (var idx in mergedRows) {
      if (!mergedRows[idx]) continue;
      var m = mergedRows[idx];
      sheet.getRange(m.row, COLS_NARZ.ILOSC + 1).setValue(m.ilosc);
      if (m.opis) {
        sheet.getRange(m.row, COLS_NARZ.OPIS + 1).setValue(m.opis);
      }
      if (m.kategoria) {
        sheet.getRange(m.row, COLS_NARZ.KATEGORIA + 1).setValue(m.kategoria);
      }
    }

    // Usuń duplikaty od dołu żeby nie przesuwać indeksów
    toDelete.sort(function (a, b) { return b - a; });
    for (var d = 0; d < toDelete.length; d++) {
      sheet.deleteRow(toDelete[d]);
    }

    CacheService.getScriptCache().remove("narzedzia");
    return { success: true, merged: toDelete.length };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}
function onEditMergeNarzedzia(e) {
  try {
    var sheet = e.source.getActiveSheet();
    if (sheet.getName() !== SHEET_NARZEDZIA) return;
    mergeDuplicateNarzedzia();
  } catch (err) { }
}
function setupMergeTrigger() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  // Usuń stare triggery tej funkcji
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onEditMergeNarzedzia') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Nowy trigger
  ScriptApp.newTrigger('onEditMergeNarzedzia')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}
function zwiekszStanNarzedzia(kod, ilosc) {
  var sheet = getSheet(SHEET_NARZEDZIA);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COLS_NARZ.KOD]) === String(kod)) {
      var stan = Number(data[i][COLS_NARZ.ILOSC]) || 0;
      sheet.getRange(i + 1, COLS_NARZ.ILOSC + 1).setValue(stan + ilosc);
      break;
    }
  }
  CacheService.getScriptCache().remove("narzedzia");
}

// ============================================
// USZKODZONE
// ============================================

function getUszkodzone() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("uszkodzone");
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { }
  }

  var sheet = getSheet(SHEET_USZKODZONE);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  var uszkodzone = data.map(function (r) {
    return {
      id: String(r[0]),
      kodNarzedzia: String(r[1] || ''),
      nazwaNarzedzia: String(r[2] || ''),
      opisUszkodzenia: String(r[3] || ''),
      data: formatDate(r[4]),
      ilosc: Number(r[5]) || 1
    };
  });

  try {
    cache.put("uszkodzone", JSON.stringify(uszkodzone), CACHE_TTL);
  } catch (e) { }
  return uszkodzone;
}

function addUszkodzone(kodNarzedzia, nazwaNarzedzia, opisUszkodzenia, ilosc) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_USZKODZONE);
    var id = generateId('USZ');
    var qty = Number(ilosc) || 1;
    sheet.appendRow([id, kodNarzedzia, nazwaNarzedzia, opisUszkodzenia, new Date(), qty]);

    // Automatycznie oddaj uszkodzone — zmniejsz stan w magazynie
    zmniejszStanNarzedzia(kodNarzedzia, qty);

    // Automatycznie zwróć aktywne wypożyczenia tego narzędzia (do ilości uszkodzonych)
    autoReturnDamaged(kodNarzedzia, qty);

    CacheService.getScriptCache().remove("uszkodzone");
    CacheService.getScriptCache().remove("narzedzia");
    return { success: true, id: id };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function zmniejszStanNarzedzia(kod, ilosc) {
  var sheet = getSheet(SHEET_NARZEDZIA);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COLS_NARZ.KOD]) === String(kod)) {
      var stan = Number(data[i][COLS_NARZ.ILOSC]) || 0;
      sheet.getRange(i + 1, COLS_NARZ.ILOSC + 1).setValue(Math.max(0, stan - ilosc));
      break;
    }
  }
  CacheService.getScriptCache().remove("narzedzia");
}

function autoReturnDamaged(kod, maxQty) {
  var sheet = getSheet(SHEET_WYPOZYCZENIA);
  var data = sheet.getDataRange().getValues();
  var returned = 0;

  for (var i = 1; i < data.length && returned < maxQty; i++) {
    if (String(data[i][1]) === String(kod) && (!data[i][8] || String(data[i][8]).length === 0)) {
      sheet.getRange(i + 1, 9).setValue(new Date());
      var ilosc = Number(data[i][9]) || 1;
      returned += ilosc;
    }
  }
}