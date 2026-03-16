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

function parseDateForSort_(val) {
  if (!val) return 0;
  if (val instanceof Date) return val.getTime();
  var str = String(val);
  var m = str.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])).getTime();
  var m2 = str.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m2) return new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3])).getTime();
  return 0;
}

function wydajBatch(idOsoby, items, operator) {
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

    var isRobocza = osobaImie.indexOf('[Robocza]') === 0 || osobaImie.indexOf('[Zamówienia]') === 0;
    var katData = katSheet.getDataRange().getValues();
    var errors = [];

    for (var idx = 0; idx < items.length; idx++) {
      var item = items[idx];
      var qty = Number(item.qty) || 1;

      // Determine status
      var status = (item.kategoria === 'Z') ? 'Zuzyte' : 'Wydane';

      // Handle photo (optional for any category)
      var photoUrl = '';
      var opId = generateId('OP');
      if (item.photoBase64) {
        photoUrl = savePhotoToDrive(item.photoBase64, 'Wydania', opId);
      }

      // Damaged items: log only, no stock change
      if (item.damaged) {
        przesSheet.appendRow([
          opId, new Date(), osobaImie,
          item.nazwaWys, item.sn || '',
          qty, item.kategoria || 'N', 'Uszkodzone', photoUrl, '', '',
          item.opisUszkodzenia || '',
          operator || ''
        ]);
        continue;
      }

      // Custom items: log only, no stock change
      if (item.custom) {
        przesSheet.appendRow([
          opId, new Date(), osobaImie,
          item.nazwaWys, '',
          qty, item.kategoria || 'N', status, photoUrl, '', '', '',
          operator || ''
        ]);
        continue;
      }

      // Robocza list: log only, no stock change
      if (isRobocza) {
        var kod = '';
        for (var ri = 1; ri < katData.length; ri++) {
          if (String(katData[ri][COLS_KATALOG.NAZWA_WYSWIETLANA]) === item.nazwaWys) {
            kod = String(katData[ri][COLS_KATALOG.NAZWA_SYSTEMOWA] || '');
            break;
          }
        }
        przesSheet.appendRow([
          opId, new Date(), osobaImie,
          kod || item.nazwaWys, item.sn || '',
          qty, item.kategoria || 'N', status, photoUrl, '', '', '',
          operator || ''
        ]);
        continue;
      }

      if (item.sn) {
        // ---- SN items: find exact match (single row) ----
        var resolved = null;
        for (var i = 1; i < katData.length; i++) {
          var r = katData[i];
          if (String(r[COLS_KATALOG.NAZWA_WYSWIETLANA]) !== item.nazwaWys) continue;
          var stock = Number(r[COLS_KATALOG.AKTUALNIE_NA_STANIE]) || 0;
          if (stock <= 0) continue;
          if (extractSN(r[COLS_KATALOG.SN]) === item.sn) {
            resolved = { row: i, data: r, stock: stock };
            break;
          }
        }

        if (resolved) {
          var newStock = resolved.stock - qty;
          katSheet.getRange(resolved.row + 1, COLS_KATALOG.AKTUALNIE_NA_STANIE + 1).setValue(newStock);
          katSheet.getRange(resolved.row + 1, COLS_KATALOG.OSTATNIO_WIDZIANE + 1).setValue(new Date());
          katData[resolved.row][COLS_KATALOG.AKTUALNIE_NA_STANIE] = newStock;

          przesSheet.appendRow([
            opId, new Date(), osobaImie,
            String(resolved.data[COLS_KATALOG.NAZWA_SYSTEMOWA] || ''),
            extractSN(resolved.data[COLS_KATALOG.SN]),
            qty, item.kategoria, status, photoUrl, '', '', '',
            operator || ''
          ]);
        } else {
          errors.push('Brak na stanie: ' + (item.nazwaWys || '') + ' SN:' + (item.sn || ''));
        }

      } else {
        // ---- Non-SN items: FIFO across multiple rows ----
        var candidates = [];
        for (var i = 1; i < katData.length; i++) {
          var r = katData[i];
          if (String(r[COLS_KATALOG.NAZWA_WYSWIETLANA]) !== item.nazwaWys) continue;
          var stock = Number(r[COLS_KATALOG.AKTUALNIE_NA_STANIE]) || 0;
          if (stock <= 0) continue;
          candidates.push({
            row: i, data: r, stock: stock,
            dp: String(r[COLS_KATALOG.DATA_PRZESUN] || '')
          });
        }

        // Sort FIFO: oldest transfer first
        candidates.sort(function(a, b) {
          return parseDateForSort_(a.dp) - parseDateForSort_(b.dp);
        });

        var totalAvailable = 0;
        for (var c = 0; c < candidates.length; c++) totalAvailable += candidates[c].stock;

        if (candidates.length > 0 && totalAvailable >= qty) {
          // FIFO distribution across rows
          var remaining = qty;
          for (var c = 0; c < candidates.length && remaining > 0; c++) {
            var take = Math.min(remaining, candidates[c].stock);
            var newStock = candidates[c].stock - take;
            katSheet.getRange(candidates[c].row + 1, COLS_KATALOG.AKTUALNIE_NA_STANIE + 1).setValue(newStock);
            katSheet.getRange(candidates[c].row + 1, COLS_KATALOG.OSTATNIO_WIDZIANE + 1).setValue(new Date());
            katData[candidates[c].row][COLS_KATALOG.AKTUALNIE_NA_STANIE] = newStock;
            remaining -= take;
          }

          // Log one operation using first candidate's data
          var first = candidates[0];
          przesSheet.appendRow([
            opId, new Date(), osobaImie,
            String(first.data[COLS_KATALOG.NAZWA_SYSTEMOWA] || ''),
            '',
            qty, item.kategoria, status, photoUrl, '', '', '',
            operator || ''
          ]);
        } else {
          errors.push('Brak na stanie: ' + (item.nazwaWys || '') + ' (potrzeba: ' + qty + ', dostępne: ' + totalAvailable + ')');
        }
      }
    }

    CacheService.getScriptCache().remove("katalog");

    if (errors.length > 0) {
      if (errors.length === items.length) {
        return { success: false, error: errors.join('; ') };
      }
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

function zwrocOperacje(idOperacji, photoBase64, operator) {
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

      // Handle photo (optional)
      if (photoBase64) {
        var photoUrl = savePhotoToDrive(photoBase64, 'Zwroty', idOperacji);
        przesSheet.getRange(i + 1, COLS_PRZES.ZDJECIE_ZWROT_URL + 1).setValue(photoUrl);
      }

      // Update status, return date, operator
      przesSheet.getRange(i + 1, COLS_PRZES.STATUS + 1).setValue('Zwrocone');
      przesSheet.getRange(i + 1, COLS_PRZES.DATA_ZWROTU + 1).setValue(new Date());
      if (operator) przesSheet.getRange(i + 1, COLS_PRZES.OPERATOR + 1).setValue(operator);

      // Increment stock in Katalog
      var nazwaSys = String(przesData[i][COLS_PRZES.NAZWA_SYSTEMOWA]);
      var ilosc = Number(przesData[i][COLS_PRZES.ILOSC]) || 1;
      incrementStock(katSheet, nazwaSys, ilosc);
      updateOstatnioWidziane(katSheet, nazwaSys);

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

function zwrocBatch(ids, photoDataMap, qtyMap, operator) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var przesSheet = ss.getSheetByName(SHEET_PRZESUNIECIA);
    var katSheet = ss.getSheetByName(SHEET_KATALOG);
    var przesData = przesSheet.getDataRange().getValues();
    var count = 0;
    var photoMap = photoDataMap || {};
    var partialQty = qtyMap || {};

    for (var idx = 0; idx < ids.length; idx++) {
      var idOp = String(ids[idx]);
      for (var i = 1; i < przesData.length; i++) {
        if (String(przesData[i][COLS_PRZES.ID_OPERACJI]) !== idOp) continue;
        if (String(przesData[i][COLS_PRZES.STATUS]) !== 'Wydane') continue;

        var nazwaSys = String(przesData[i][COLS_PRZES.NAZWA_SYSTEMOWA]);
        var totalQty = Number(przesData[i][COLS_PRZES.ILOSC]) || 1;
        var returnQty = partialQty[idOp] ? Number(partialQty[idOp]) : totalQty;
        if (returnQty > totalQty) returnQty = totalQty;
        if (returnQty < 1) returnQty = 1;

        var photoUrl = '';
        if (photoMap[idOp]) {
          photoUrl = savePhotoToDrive(photoMap[idOp], 'Zwroty', idOp);
        }

        if (returnQty < totalQty) {
          // Partial return: reduce qty in existing row, add new row for returned portion
          przesSheet.getRange(i + 1, COLS_PRZES.ILOSC + 1).setValue(totalQty - returnQty);

          // New row for returned portion
          var newRow = [];
          for (var c = 0; c < przesData[i].length; c++) newRow.push(przesData[i][c]);
          newRow[COLS_PRZES.ID_OPERACJI] = idOp + '_z' + Utilities.formatDate(new Date(), 'Europe/Warsaw', 'HHmmss');
          newRow[COLS_PRZES.ILOSC] = returnQty;
          newRow[COLS_PRZES.STATUS] = 'Zwrocone';
          newRow[COLS_PRZES.DATA_ZWROTU] = new Date();
          if (photoUrl) newRow[COLS_PRZES.ZDJECIE_ZWROT_URL] = photoUrl;
          if (operator) newRow[COLS_PRZES.OPERATOR] = operator;
          przesSheet.appendRow(newRow);
        } else {
          // Full return: mark entire row as returned
          if (photoUrl) {
            przesSheet.getRange(i + 1, COLS_PRZES.ZDJECIE_ZWROT_URL + 1).setValue(photoUrl);
          }
          przesSheet.getRange(i + 1, COLS_PRZES.STATUS + 1).setValue('Zwrocone');
          przesSheet.getRange(i + 1, COLS_PRZES.DATA_ZWROTU + 1).setValue(new Date());
          if (operator) przesSheet.getRange(i + 1, COLS_PRZES.OPERATOR + 1).setValue(operator);
        }

        var osobaName = String(przesData[i][COLS_PRZES.OSOBA] || '');
        if (osobaName.indexOf('[Robocza]') !== 0) {
          incrementStock(katSheet, nazwaSys, returnQty);
        }
        updateOstatnioWidziane(katSheet, nazwaSys);

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

function updateOstatnioWidziane(katSheet, nazwaSys) {
  var data = katSheet.getDataRange().getValues();
  var now = new Date();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COLS_KATALOG.NAZWA_SYSTEMOWA]) === nazwaSys) {
      katSheet.getRange(i + 1, COLS_KATALOG.OSTATNIO_WIDZIANE + 1).setValue(now);
      break;
    }
  }
}

// ============================================
// INWENTARYZACJA — aktualizacja stanów
// ============================================

function inwentaryzujBatch(updates) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_KATALOG);
    var data = sheet.getDataRange().getValues();
    var now = new Date();

    var updateMap = {};
    for (var u = 0; u < updates.length; u++) {
      updateMap[updates[u].nazwaSys] = Number(updates[u].nowyStanIlosciowy);
    }

    var updated = 0;
    var unchanged = 0;
    var logRows = [];
    var operator = updates.length > 0 && updates[0].operator ? updates[0].operator : 'System';

    for (var i = 1; i < data.length; i++) {
      var nazwaSys = String(data[i][COLS_KATALOG.NAZWA_SYSTEMOWA]);
      var kategoria = String(data[i][COLS_KATALOG.KATEGORIA]);
      sheet.getRange(i + 1, COLS_KATALOG.OSTATNIO_WIDZIANE + 1).setValue(now);

      if (updateMap.hasOwnProperty(nazwaSys)) {
        var newVal = updateMap[nazwaSys];
        var oldVal = Number(data[i][COLS_KATALOG.AKTUALNIE_NA_STANIE]) || 0;
        if (newVal !== oldVal) {
          sheet.getRange(i + 1, COLS_KATALOG.AKTUALNIE_NA_STANIE + 1).setValue(newVal);
          logRows.push([
            generateId('INW'),
            now,
            'Inwentaryzacja',
            nazwaSys,
            String(data[i][COLS_KATALOG.SN] || ''),
            newVal - oldVal,
            kategoria,
            'Inwentaryzacja',
            '', '', '',
            'Stan: ' + oldVal + ' → ' + newVal,
            operator
          ]);
          updated++;
        } else {
          unchanged++;
        }
      } else {
        unchanged++;
      }
    }

    if (logRows.length) {
      var przesSheet = getSheet(SHEET_PRZESUNIECIA);
      var startRow = przesSheet.getLastRow() + 1;
      przesSheet.getRange(startRow, 1, logRows.length, 13).setValues(logRows);
    }

    CacheService.getScriptCache().remove("katalog");
    return { success: true, updated: updated, unchanged: unchanged };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

// ============================================
// UZUPEŁNIENIE STANU (np. zwrot zużywalnych)
// ============================================

function uzupelnijStan(nazwaWys, qty) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_KATALOG);
    var data = sheet.getDataRange().getValues();
    var updated = 0;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][COLS_KATALOG.NAZWA_WYSWIETLANA]) === nazwaWys) {
        var stan = Number(data[i][COLS_KATALOG.AKTUALNIE_NA_STANIE]) || 0;
        var pocz = Number(data[i][COLS_KATALOG.STAN_POCZATKOWY]) || 0;
        var newStan = Math.min(stan + qty, pocz);
        sheet.getRange(i + 1, COLS_KATALOG.AKTUALNIE_NA_STANIE + 1).setValue(newStan);
        updated += (newStan - stan);
        break;
      }
    }

    CacheService.getScriptCache().remove("katalog");
    return { success: true, updated: updated };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function zwiekszStan(nazwaWys, qty) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_KATALOG);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][COLS_KATALOG.NAZWA_WYSWIETLANA]) === nazwaWys) {
        var pocz = Number(data[i][COLS_KATALOG.STAN_POCZATKOWY]) || 0;
        var stan = Number(data[i][COLS_KATALOG.AKTUALNIE_NA_STANIE]) || 0;
        sheet.getRange(i + 1, COLS_KATALOG.STAN_POCZATKOWY + 1).setValue(pocz + qty);
        sheet.getRange(i + 1, COLS_KATALOG.AKTUALNIE_NA_STANIE + 1).setValue(stan + qty);
        CacheService.getScriptCache().remove("katalog");
        return { success: true, newStan: pocz + qty };
      }
    }
    return { success: false, error: 'Nie znaleziono pozycji' };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

// ============================================
// USZKODZENIA
// ============================================

function zglosUszkodzenie(nazwaSys, sn, opis, ilosc, operator) {
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
      opis || '',   // Opis_Uszkodzenia
      operator || '' // Operator
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
// IMPORT HISTORII Z EXCELA STARGEO
// ============================================

function importExcelHistory() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var katSheet = ss.getSheetByName(SHEET_KATALOG);
    var przesSheet = ss.getSheetByName(SHEET_PRZESUNIECIA);
    var katData = katSheet.getDataRange().getValues();
    var now = new Date();
    var dateWydanie = new Date('2026-02-13');
    var dateZwrot = new Date('2026-02-28');
    var operator = 'Import-Stargeo';
    var results = [];

    // === WYDANE (19 items) - import as Wydane + decrement stock ===
    var wydane = [
      // Okilka Ihor
      {osoba: 'Okilka Ihor', kod: 'SYN/4/974', sn: '', qty: 1, kat: 'N'},
      {osoba: 'Okilka Ihor', kod: 'SYN/5/9125', sn: '', qty: 1, kat: 'N'},
      {osoba: 'Okilka Ihor', kod: 'SYN/1/100586', sn: 'VO500238', qty: 1, kat: 'E'},
      {osoba: 'Okilka Ihor', kod: 'SYN/1/100585', sn: 'VO500237', qty: 1, kat: 'E'},
      // Wdowski Łukasz
      {osoba: 'Wdowski Łukasz', kod: 'SYN/5/9232', sn: 'CN40696', qty: 1, kat: 'E'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/1/100133', sn: 'J422245', qty: 1, kat: 'E'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/7/3', sn: '', qty: 1, kat: 'N'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/4/89', sn: '', qty: 1, kat: 'N'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/4/13', sn: '', qty: 1, kat: 'N'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/4/8', sn: '', qty: 1, kat: 'N'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/4/75', sn: '', qty: 1, kat: 'N'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/6/76', sn: '', qty: 1, kat: 'N'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/6/2353', sn: '', qty: 1, kat: 'N'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/4/34', sn: '', qty: 1, kat: 'N'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/6/6', sn: '', qty: 1, kat: 'N'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/1/100580', sn: 'C655030', qty: 1, kat: 'E'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/4/22', sn: '', qty: 1, kat: 'N'},
      // Sochacki Krzysztof
      {osoba: 'Sochacki Krzysztof', kod: 'SYN/1/100239', sn: 'JO20727', qty: 1, kat: 'E'},
      {osoba: 'Sochacki Krzysztof', kod: 'SYN/1/428', sn: 'J472245', qty: 1, kat: 'E'}
    ];

    // === ZWRÓCONE (5 items) - import as Wydane+Zwrócone, NO stock change ===
    var zwrocone = [
      {osoba: 'Okilka Ihor', kod: 'SYN/7/906', sn: '', qty: 1, kat: 'N'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/4/110', sn: '', qty: 2, kat: 'N'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/4/84', sn: '', qty: 1, kat: 'N'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/6/52', sn: '', qty: 1, kat: 'N'},
      {osoba: 'Wdowski Łukasz', kod: 'SYN/1/100587', sn: 'VO500236', qty: 1, kat: 'E'}
    ];

    // Process WYDANE items
    var wydaneRows = [];
    for (var i = 0; i < wydane.length; i++) {
      var item = wydane[i];
      var opId = generateId('IMP');
      wydaneRows.push([
        opId, dateWydanie, item.osoba,
        item.kod, item.sn, item.qty, item.kat,
        'Wydane', '', '', '', '', operator
      ]);
      // Decrement stock
      decrementStock(katSheet, item.kod, item.qty);
      results.push({op: 'wydane', kod: item.kod, osoba: item.osoba});
    }

    // Process ZWRÓCONE items (wydanie + zwrot = net zero stock impact)
    var zwroconeRows = [];
    for (var j = 0; j < zwrocone.length; j++) {
      var item2 = zwrocone[j];
      var opId2 = generateId('IMP');
      zwroconeRows.push([
        opId2, dateWydanie, item2.osoba,
        item2.kod, item2.sn, item2.qty, item2.kat,
        'Zwrocone', '', '', dateZwrot, '', operator
      ]);
      results.push({op: 'zwrocone', kod: item2.kod, osoba: item2.osoba});
    }

    // Batch write all rows
    var allRows = wydaneRows.concat(zwroconeRows);
    if (allRows.length > 0) {
      var startRow = przesSheet.getLastRow() + 1;
      przesSheet.getRange(startRow, 1, allRows.length, 13).setValues(allRows);
    }

    CacheService.getScriptCache().remove('katalog');
    return {success: true, wydane: wydaneRows.length, zwrocone: zwroconeRows.length, details: results};
  } catch (e) {
    return {success: false, error: e.toString()};
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

function importMissingRecords() {
  var przesSheet = getSheet(SHEET_PRZESUNIECIA);
  var dateWydanie = new Date('2026-02-17');
  var operator = 'Import-Stargeo';
  var opId = generateId('IMP');

  // SYN/1/1051 - Szlifierka katowa Hikoki G18DBL - Okilka Ihor
  // aktStan juz = 0, wiec NIE zmieniamy stanu - tylko dodajemy rekord wydania
  przesSheet.appendRow([
    opId, dateWydanie, 'Okilka Ihor',
    'SYN/1/1051', 'J901230', 1, 'E',
    'Wydane', '', '', '', '', operator
  ]);

  return {success: true, imported: [{kod: 'SYN/1/1051', sn: 'J901230', osoba: 'Okilka Ihor', opId: opId}]};
}

function updateListaQty(idOp, newQty) {
  var sheet = getSheet(SHEET_PRZESUNIECIA);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COLS_PRZES.ID_OPERACJI]) === idOp) {
      var osoba = String(data[i][COLS_PRZES.OSOBA] || '');
      if (osoba.indexOf('[Lista]') !== 0 && osoba.indexOf('[Robocza]') !== 0 && osoba.indexOf('[Zamówienia]') !== 0) {
        return { success: false, error: 'Edycja ilości tylko dla list' };
      }
      var qty = Math.max(1, Math.floor(Number(newQty)));
      sheet.getRange(i + 1, COLS_PRZES.ILOSC + 1).setValue(qty);
      return { success: true, newQty: qty };
    }
  }
  return { success: false, error: 'Nie znaleziono operacji' };
}

function updateOpisUszkodzenia(idOp, opis) {
  var sheet = getSheet(SHEET_PRZESUNIECIA);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COLS_PRZES.ID_OPERACJI]) === idOp) {
      sheet.getRange(i + 1, COLS_PRZES.OPIS_USZKODZENIA + 1).setValue(opis);
      return { success: true };
    }
  }
  return { success: false, error: 'Nie znaleziono operacji' };
}
