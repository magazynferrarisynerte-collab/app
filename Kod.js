// ============================================
// MAGAZYN / WYPOŻYCZALNIA - BACKEND v3.0
// Nowy model: Katalog + Przesunięcia
// ============================================

const CACHE_TTL = 60 * 5;
const SPREADSHEET_ID = '1OCm_8VZ1Q-z1sXGsthJKUgHS_H13rdXa9uzGKdNrc6A';

const SHEET_KATALOG = "Katalog";
const SHEET_OSOBY = "Osoby";
const SHEET_PRZESUNIECIA = "Przesunięcia";
const DRIVE_FOLDER_NAME = "Magazyn_Zdjecia";

const COLS_KATALOG = {
  ID: 0,
  NAZWA_SYSTEMOWA: 1,
  NAZWA_WYSWIETLANA: 2,
  KATEGORIA: 3,
  SN: 4,
  STAN_POCZATKOWY: 5,
  AKTUALNIE_NA_STANIE: 6
};

const COLS_PRZES = {
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
  OPIS_USZKODZENIA: 11
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
  var match = str.match(/(?:s\/n|sn)[.:;\s]*([^\s,]+)/i);
  if (match) return match[1];
  return str;
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
// OSOBY (bez zmian)
// ============================================

function getOsoby() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("osoby");
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { }
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
  } catch (e) { }
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
// KATALOG
// ============================================

function getKatalog() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("katalog");
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { }
  }

  var sheet = getSheet(SHEET_KATALOG);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

  var katalog = data.map(function (r) {
    var sn = extractSN(r[COLS_KATALOG.SN]);
    var kategoria = String(r[COLS_KATALOG.KATEGORIA] || '');
    if (sn && kategoria !== 'specjalne') kategoria = 'elektronarzedzia';
    return {
      id: String(r[COLS_KATALOG.ID]),
      nazwaSys: String(r[COLS_KATALOG.NAZWA_SYSTEMOWA] || ''),
      nazwaWys: String(r[COLS_KATALOG.NAZWA_WYSWIETLANA] || ''),
      kategoria: kategoria,
      sn: sn,
      stanPoczatkowy: Number(r[COLS_KATALOG.STAN_POCZATKOWY]) || 0,
      aktualnieNaStanie: Number(r[COLS_KATALOG.AKTUALNIE_NA_STANIE]) || 0
    };
  }).sort(function (a, b) {
    return a.nazwaWys.localeCompare(b.nazwaWys);
  });

  try {
    cache.put("katalog", JSON.stringify(katalog), CACHE_TTL);
  } catch (e) { }
  return katalog;
}

function getKatalogGrouped() {
  var katalog = getKatalog();
  var groups = {};

  for (var i = 0; i < katalog.length; i++) {
    var k = katalog[i];
    var key = k.nazwaWys;
    if (!groups[key]) {
      groups[key] = {
        nazwaWys: k.nazwaWys,
        kategoria: k.kategoria,
        totalStock: 0,
        totalPoczatkowy: 0,
        items: []
      };
    }
    groups[key].totalStock += k.aktualnieNaStanie;
    groups[key].totalPoczatkowy += k.stanPoczatkowy;
    groups[key].items.push({
      id: k.id,
      nazwaSys: k.nazwaSys,
      sn: k.sn,
      aktualnieNaStanie: k.aktualnieNaStanie
    });
  }

  var result = [];
  for (var key in groups) {
    result.push(groups[key]);
  }
  return result.sort(function (a, b) {
    return a.nazwaWys.localeCompare(b.nazwaWys);
  });
}

function addKatalogItem(nazwaSys, nazwaWys, kategoria, sn, stanPoczatkowy) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_KATALOG);
    var id = generateId('KT');
    var qty = Number(stanPoczatkowy) || 1;
    sheet.appendRow([id, nazwaSys.trim(), nazwaWys.trim(), kategoria, sn || '', qty, qty]);
    CacheService.getScriptCache().remove("katalog");
    return { success: true, id: id };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function getAvailableSNs(nazwaWys) {
  var katalog = getKatalog();
  var result = [];
  for (var i = 0; i < katalog.length; i++) {
    var k = katalog[i];
    if (k.nazwaWys === nazwaWys && k.aktualnieNaStanie > 0 && k.sn) {
      result.push({
        nazwaSys: k.nazwaSys,
        sn: k.sn,
        aktualnieNaStanie: k.aktualnieNaStanie
      });
    }
  }
  return result;
}

// ============================================
// ALIAS RESOLUTION
// ============================================

function resolveAvailableItem(nazwaWys, preferredSN) {
  var sheet = getSheet(SHEET_KATALOG);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var fallback = null;

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (String(r[COLS_KATALOG.NAZWA_WYSWIETLANA]) !== nazwaWys) continue;
    var stock = Number(r[COLS_KATALOG.AKTUALNIE_NA_STANIE]) || 0;
    if (stock <= 0) continue;

    var item = {
      rowIndex: i + 2,
      id: String(r[COLS_KATALOG.ID]),
      nazwaSys: String(r[COLS_KATALOG.NAZWA_SYSTEMOWA]),
      nazwaWys: String(r[COLS_KATALOG.NAZWA_WYSWIETLANA]),
      kategoria: String(r[COLS_KATALOG.KATEGORIA]),
      sn: extractSN(r[COLS_KATALOG.SN]),
      aktualnieNaStanie: stock
    };

    if (preferredSN && item.sn === preferredSN) {
      return item;
    }
    if (!fallback) {
      fallback = item;
    }
  }

  return fallback;
}

// ============================================
// ZDJĘCIA - GOOGLE DRIVE
// ============================================

function getOrCreateFolder(parent, name) {
  var folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function savePhotoToDrive(base64Data, subfolder, operationId) {
  if (!base64Data) return '';

  try {
    var rootFolders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
    var root = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);

    var sub = getOrCreateFolder(root, subfolder);

    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      'image/jpeg',
      operationId + '_' + Date.now() + '.jpg'
    );

    var file = sub.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return file.getUrl();
  } catch (e) {
    Logger.log('Photo save error: ' + e.toString());
    return '';
  }
}

// ============================================
// WYDANIE (WYPOŻYCZENIE / ZUŻYCIE)
// ============================================

function wydajBatch(idOsoby, items) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var katSheet = ss.getSheetByName(SHEET_KATALOG);
    var przesSheet = ss.getSheetByName(SHEET_PRZESUNIECIA);

    var osobaImie = '';
    if (idOsoby) {
      var osoby = getOsoby();
      for (var o = 0; o < osoby.length; o++) {
        if (osoby[o].id === idOsoby) { osobaImie = osoby[o].imie; break; }
      }
    }

    var katData = katSheet.getDataRange().getValues();
    var errors = [];

    for (var idx = 0; idx < items.length; idx++) {
      var item = items[idx];
      var qty = Number(item.qty) || 1;
      var resolved = null;

      // Find the matching Katalog row
      for (var i = 1; i < katData.length; i++) {
        var r = katData[i];
        if (String(r[COLS_KATALOG.NAZWA_WYSWIETLANA]) !== item.nazwaWys) continue;
        var stock = Number(r[COLS_KATALOG.AKTUALNIE_NA_STANIE]) || 0;
        if (stock < qty) continue;

        if (item.sn) {
          if (extractSN(r[COLS_KATALOG.SN]) === item.sn) {
            resolved = { row: i, data: r, stock: stock };
            break;
          }
        } else {
          resolved = { row: i, data: r, stock: stock };
          break;
        }
      }

      if (!resolved) {
        errors.push('Brak dostępnego: ' + item.nazwaWys);
        continue;
      }

      // Decrement stock
      var newStock = resolved.stock - qty;
      katSheet.getRange(resolved.row + 1, COLS_KATALOG.AKTUALNIE_NA_STANIE + 1).setValue(newStock);
      katData[resolved.row][COLS_KATALOG.AKTUALNIE_NA_STANIE] = newStock;

      // Determine status
      var status = (item.kategoria === 'zuzywalne') ? 'Zuzyte' : 'Wydane';

      // Handle photo for specjalne
      var photoUrl = '';
      if (item.kategoria === 'specjalne' && item.photoBase64) {
        var opId = generateId('OP');
        photoUrl = savePhotoToDrive(item.photoBase64, 'Wydania', opId);
      }

      var opId2 = photoUrl ? opId : generateId('OP');

      // Append to Przesunięcia
      przesSheet.appendRow([
        opId2,                                               // ID_Operacji
        new Date(),                                          // Data_Wydania
        osobaImie,                                           // Osoba
        String(resolved.data[COLS_KATALOG.NAZWA_SYSTEMOWA]), // Nazwa_Systemowa
        String(resolved.data[COLS_KATALOG.SN] || ''),        // SN
        qty,                                                 // Ilosc
        item.kategoria,                                      // Kategoria
        status,                                              // Status
        photoUrl,                                            // Zdjecie_Wydanie_URL
        '',                                                  // Zdjecie_Zwrot_URL
        '',                                                  // Data_Zwrotu
        ''                                                   // Opis_Uszkodzenia
      ]);
    }

    CacheService.getScriptCache().remove("katalog");

    if (errors.length > 0) {
      return { success: true, warnings: errors };
    }
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

function zwrocOperacje(idOperacji, photoBase64) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var przesSheet = ss.getSheetByName(SHEET_PRZESUNIECIA);
    var katSheet = ss.getSheetByName(SHEET_KATALOG);
    var przesData = przesSheet.getDataRange().getValues();

    for (var i = 1; i < przesData.length; i++) {
      if (String(przesData[i][COLS_PRZES.ID_OPERACJI]) !== String(idOperacji)) continue;

      var status = String(przesData[i][COLS_PRZES.STATUS]);
      if (status !== 'Wydane') {
        return { success: false, error: 'Nie można zwrócić — status: ' + status };
      }

      var kategoria = String(przesData[i][COLS_PRZES.KATEGORIA]);

      // Handle photo for specjalne
      if (kategoria === 'specjalne' && photoBase64) {
        var photoUrl = savePhotoToDrive(photoBase64, 'Zwroty', idOperacji);
        przesSheet.getRange(i + 1, COLS_PRZES.ZDJECIE_ZWROT_URL + 1).setValue(photoUrl);
      }

      // Update status and return date
      przesSheet.getRange(i + 1, COLS_PRZES.STATUS + 1).setValue('Zwrocone');
      przesSheet.getRange(i + 1, COLS_PRZES.DATA_ZWROTU + 1).setValue(new Date());

      // Increment stock in Katalog
      var nazwaSys = String(przesData[i][COLS_PRZES.NAZWA_SYSTEMOWA]);
      var ilosc = Number(przesData[i][COLS_PRZES.ILOSC]) || 1;
      incrementStock(katSheet, nazwaSys, ilosc);

      CacheService.getScriptCache().remove("katalog");
      return { success: true };
    }

    return { success: false, error: 'Nie znaleziono operacji' };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function zwrocBatch(ids, photoDataMap) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var przesSheet = ss.getSheetByName(SHEET_PRZESUNIECIA);
    var katSheet = ss.getSheetByName(SHEET_KATALOG);
    var przesData = przesSheet.getDataRange().getValues();
    var count = 0;
    var photoMap = photoDataMap || {};

    for (var idx = 0; idx < ids.length; idx++) {
      var idOp = String(ids[idx]);
      for (var i = 1; i < przesData.length; i++) {
        if (String(przesData[i][COLS_PRZES.ID_OPERACJI]) !== idOp) continue;
        if (String(przesData[i][COLS_PRZES.STATUS]) !== 'Wydane') continue;

        var kategoria = String(przesData[i][COLS_PRZES.KATEGORIA]);

        // Handle photo for specjalne
        if (kategoria === 'specjalne' && photoMap[idOp]) {
          var photoUrl = savePhotoToDrive(photoMap[idOp], 'Zwroty', idOp);
          przesSheet.getRange(i + 1, COLS_PRZES.ZDJECIE_ZWROT_URL + 1).setValue(photoUrl);
        }

        przesSheet.getRange(i + 1, COLS_PRZES.STATUS + 1).setValue('Zwrocone');
        przesSheet.getRange(i + 1, COLS_PRZES.DATA_ZWROTU + 1).setValue(new Date());

        var nazwaSys = String(przesData[i][COLS_PRZES.NAZWA_SYSTEMOWA]);
        var ilosc = Number(przesData[i][COLS_PRZES.ILOSC]) || 1;
        incrementStock(katSheet, nazwaSys, ilosc);

        przesData[i][COLS_PRZES.STATUS] = 'Zwrocone';
        count++;
        break;
      }
    }

    CacheService.getScriptCache().remove("katalog");
    return { success: true, count: count };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function incrementStock(katSheet, nazwaSys, qty) {
  var data = katSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COLS_KATALOG.NAZWA_SYSTEMOWA]) === nazwaSys) {
      var stan = Number(data[i][COLS_KATALOG.AKTUALNIE_NA_STANIE]) || 0;
      katSheet.getRange(i + 1, COLS_KATALOG.AKTUALNIE_NA_STANIE + 1).setValue(stan + qty);
      break;
    }
  }
}

function decrementStock(katSheet, nazwaSys, qty) {
  var data = katSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COLS_KATALOG.NAZWA_SYSTEMOWA]) === nazwaSys) {
      var stan = Number(data[i][COLS_KATALOG.AKTUALNIE_NA_STANIE]) || 0;
      katSheet.getRange(i + 1, COLS_KATALOG.AKTUALNIE_NA_STANIE + 1).setValue(Math.max(0, stan - qty));
      break;
    }
  }
}

// ============================================
// USZKODZENIA
// ============================================

function zglosUszkodzenie(nazwaSys, sn, opis, ilosc) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var przesSheet = ss.getSheetByName(SHEET_PRZESUNIECIA);
    var katSheet = ss.getSheetByName(SHEET_KATALOG);

    var qty = Number(ilosc) || 1;
    var opId = generateId('OP');

    przesSheet.appendRow([
      opId,         // ID_Operacji
      new Date(),   // Data_Wydania
      '',           // Osoba
      nazwaSys,     // Nazwa_Systemowa
      sn || '',     // SN
      qty,          // Ilosc
      '',           // Kategoria
      'Uszkodzone', // Status
      '',           // Zdjecie_Wydanie_URL
      '',           // Zdjecie_Zwrot_URL
      '',           // Data_Zwrotu
      opis || ''    // Opis_Uszkodzenia
    ]);

    decrementStock(katSheet, nazwaSys, qty);

    CacheService.getScriptCache().remove("katalog");
    return { success: true, id: opId };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

// ============================================
// LOG (PRZESUNIĘCIA)
// ============================================

function getLog() {
  var sheet = getSheet(SHEET_PRZESUNIECIA);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();

  return data.map(function (r) {
    return {
      idOp: String(r[COLS_PRZES.ID_OPERACJI]),
      dataWydania: formatDate(r[COLS_PRZES.DATA_WYDANIA]),
      osoba: String(r[COLS_PRZES.OSOBA] || ''),
      nazwaSys: String(r[COLS_PRZES.NAZWA_SYSTEMOWA] || ''),
      sn: String(r[COLS_PRZES.SN] || ''),
      ilosc: Number(r[COLS_PRZES.ILOSC]) || 1,
      kategoria: String(r[COLS_PRZES.KATEGORIA] || ''),
      status: String(r[COLS_PRZES.STATUS] || ''),
      zdjecieWydanieUrl: String(r[COLS_PRZES.ZDJECIE_WYDANIE_URL] || ''),
      zdjecieZwrotUrl: String(r[COLS_PRZES.ZDJECIE_ZWROT_URL] || ''),
      dataZwrotu: formatDate(r[COLS_PRZES.DATA_ZWROTU]),
      opisUszkodzenia: String(r[COLS_PRZES.OPIS_USZKODZENIA] || '')
    };
  }).sort(function (a, b) {
    var aDate = a.dataZwrotu || a.dataWydania || '';
    var bDate = b.dataZwrotu || b.dataWydania || '';
    return bDate.localeCompare(aDate);
  });
}

// ============================================
// RAPORTY
// ============================================

function getSummaryData() {
  var katalog = getKatalog();
  var przesSheet = getSheet(SHEET_PRZESUNIECIA);
  var lastRow = przesSheet.getLastRow();
  var przesData = lastRow >= 2 ? przesSheet.getRange(2, 1, lastRow - 1, 12).getValues() : [];

  // Build map: nazwaSys -> nazwaWys
  var sysToWys = {};
  var groups = {};
  for (var i = 0; i < katalog.length; i++) {
    var k = katalog[i];
    sysToWys[k.nazwaSys] = k.nazwaWys;
    if (!groups[k.nazwaWys]) {
      groups[k.nazwaWys] = {
        nazwaWys: k.nazwaWys,
        kategoria: k.kategoria,
        stanPoczatkowy: 0,
        aktualnieNaStanie: 0,
        wydane: 0,
        zuzyte: 0
      };
    }
    groups[k.nazwaWys].stanPoczatkowy += k.stanPoczatkowy;
    groups[k.nazwaWys].aktualnieNaStanie += k.aktualnieNaStanie;
  }

  // Count active operations
  for (var j = 0; j < przesData.length; j++) {
    var row = przesData[j];
    var nazwaSys = String(row[COLS_PRZES.NAZWA_SYSTEMOWA]);
    var status = String(row[COLS_PRZES.STATUS]);
    var qty = Number(row[COLS_PRZES.ILOSC]) || 1;

    var grupKey = sysToWys[nazwaSys];
    if (grupKey && groups[grupKey]) {
      if (status === 'Wydane') groups[grupKey].wydane += qty;
      if (status === 'Zuzyte') groups[grupKey].zuzyte += qty;
    }
  }

  var result = [];
  for (var key in groups) {
    result.push(groups[key]);
  }
  return result.sort(function (a, b) { return a.nazwaWys.localeCompare(b.nazwaWys); });
}

function getRaportUszkodzonych() {
  var sheet = getSheet(SHEET_PRZESUNIECIA);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  var result = [];

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (String(r[COLS_PRZES.STATUS]) !== 'Uszkodzone') continue;
    result.push({
      nazwaSys: String(r[COLS_PRZES.NAZWA_SYSTEMOWA] || ''),
      sn: String(r[COLS_PRZES.SN] || ''),
      opisUszkodzenia: String(r[COLS_PRZES.OPIS_USZKODZENIA] || ''),
      data: formatDate(r[COLS_PRZES.DATA_WYDANIA]),
      ilosc: Number(r[COLS_PRZES.ILOSC]) || 1
    });
  }

  return result.sort(function (a, b) {
    return (b.data || '').localeCompare(a.data || '');
  });
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
// MIGRACJA (jednorazowa)
// ============================================

function migrateWypozyczenia() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var wypSheet = ss.getSheetByName("Wypożyczenia");
  if (!wypSheet) return { success: false, error: 'Brak arkusza Wypożyczenia' };

  var przesSheet = ss.getSheetByName(SHEET_PRZESUNIECIA);
  if (!przesSheet) return { success: false, error: 'Brak arkusza Przesunięcia' };

  var lastRow = wypSheet.getLastRow();
  if (lastRow < 2) return { success: true, migrated: 0 };

  var data = wypSheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var count = 0;

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var idWyp = String(r[0]);
    var nazwa = String(r[2] || '');
    var imie = String(r[5] || '');
    var dataWyp = r[7];
    var dataZwr = r[8];
    var ilosc = Number(r[9]) || 1;
    var hasReturn = dataZwr && String(dataZwr).length > 0;

    przesSheet.appendRow([
      idWyp,                          // ID_Operacji (zachowujemy stare ID)
      dataWyp || '',                  // Data_Wydania
      imie,                           // Osoba
      nazwa,                          // Nazwa_Systemowa
      '',                             // SN
      ilosc,                          // Ilosc
      '',                             // Kategoria
      hasReturn ? 'Zwrocone' : 'Wydane', // Status
      '',                             // Zdjecie_Wydanie_URL
      '',                             // Zdjecie_Zwrot_URL
      hasReturn ? dataZwr : '',       // Data_Zwrotu
      ''                              // Opis_Uszkodzenia
    ]);
    count++;
  }

  return { success: true, migrated: count };
}
