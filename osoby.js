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

  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  var osoby = data.map(function (r) {
    return {
      id: String(r[0]),
      imie: String(r[1] || ''),
      telefon: String(r[2] || ''),
      lokalizacja: String(r[3] || ''),
      email: String(r[4] || '')
    };
  }).sort(function (a, b) {
    return a.imie.localeCompare(b.imie);
  });

  try {
    cache.put("osoby", JSON.stringify(osoby), CACHE_TTL);
  } catch (e) { }
  return osoby;
}

function addOsoba(imie, telefon, lokalizacja, email) {
  var sheet = getSheet(SHEET_OSOBY);
  var id = generateId('OS');
  sheet.appendRow([id, imie.trim(), (telefon || '').trim(), (lokalizacja || '').trim(), (email || '').trim()]);
  CacheService.getScriptCache().remove("osoby");
  return { success: true, id: id };
}

function importOsoby(lista) {
  // lista = [{imie: "Jan Kowalski", telefon: "600123456"}, ...]
  var sheet = getSheet(SHEET_OSOBY);
  var lastRow = sheet.getLastRow();
  var existing = {};
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      existing[String(data[i][0]).trim().toLowerCase()] = true;
    }
  }

  var added = 0, skipped = 0;
  var rows = [];
  for (var j = 0; j < lista.length; j++) {
    var name = String(lista[j].imie || '').trim();
    if (!name) continue;
    var key = name.toLowerCase();
    if (existing[key]) { skipped++; continue; }
    existing[key] = true;
    var tel = String(lista[j].telefon || '').trim();
    var lok = String(lista[j].lokalizacja || '').trim();
    var em = String(lista[j].email || '').trim();
    rows.push([generateId('OS'), name, tel, lok, em]);
    added++;
  }

  if (rows.length) {
    sheet.getRange(lastRow + 1, 1, rows.length, 5).setValues(rows);
  }

  CacheService.getScriptCache().remove("osoby");
  return { success: true, added: added, skipped: skipped };
}

function updateOsoba(id, fields) {
  var sheet = getSheet(SHEET_OSOBY);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      if (fields.imie != null) sheet.getRange(i + 1, 2).setValue(fields.imie.trim());
      if (fields.telefon != null) sheet.getRange(i + 1, 3).setValue(fields.telefon.trim());
      if (fields.lokalizacja != null) sheet.getRange(i + 1, 4).setValue(fields.lokalizacja.trim());
      if (fields.email != null) sheet.getRange(i + 1, 5).setValue(fields.email.trim());
      CacheService.getScriptCache().remove("osoby");
      return { success: true };
    }
  }
  return { success: false, error: 'Nie znaleziono osoby' };
}

function renameOsoba(id, newName) {
  newName = (newName || '').trim();
  if (!newName) return { success: false, error: 'Pusta nazwa' };

  // Find old name
  var sheetO = getSheet(SHEET_OSOBY);
  var dataO = sheetO.getDataRange().getValues();
  var oldName = '';
  var rowIdx = -1;
  for (var i = 1; i < dataO.length; i++) {
    if (String(dataO[i][0]) === id) { oldName = String(dataO[i][1]); rowIdx = i; break; }
  }
  if (rowIdx < 0) return { success: false, error: 'Nie znaleziono osoby' };
  if (oldName === newName) return { success: true, updated: 0 };

  // Update Osoby sheet
  sheetO.getRange(rowIdx + 1, 2).setValue(newName);

  // Update all Przesunięcia entries with old name
  var sheetP = getSheet(SHEET_PRZESUNIECIA);
  var lastRow = sheetP.getLastRow();
  var updated = 0;
  if (lastRow >= 2) {
    var col = sheetP.getRange(2, COLS_PRZES.OSOBA + 1, lastRow - 1, 1).getValues();
    for (var j = 0; j < col.length; j++) {
      if (String(col[j][0]) === oldName) {
        sheetP.getRange(j + 2, COLS_PRZES.OSOBA + 1).setValue(newName);
        updated++;
      }
    }
  }

  CacheService.getScriptCache().remove("osoby");
  CacheService.getScriptCache().remove("log");
  return { success: true, updated: updated };
}
