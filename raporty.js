// ============================================
// LOG (PRZESUNIĘCIA)
// ============================================

function getLog() {
  var sheet = getSheet(SHEET_PRZESUNIECIA);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Buduj mapę nazwaSys → nazwaWys z katalogu
  var katSheet = getSheet(SHEET_KATALOG);
  var katLastRow = katSheet.getLastRow();
  var sysToWys = {};
  if (katLastRow >= 2) {
    var katData = katSheet.getRange(2, COLS_KATALOG.NAZWA_SYSTEMOWA + 1, katLastRow - 1, 2).getValues();
    for (var k = 0; k < katData.length; k++) {
      sysToWys[String(katData[k][0])] = String(katData[k][1]);
    }
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

  return data.map(function (r) {
    var ns = String(r[COLS_PRZES.NAZWA_SYSTEMOWA] || '');
    var rawWyd = r[COLS_PRZES.DATA_WYDANIA] ? new Date(r[COLS_PRZES.DATA_WYDANIA]).getTime() || 0 : 0;
    var rawZwr = r[COLS_PRZES.DATA_ZWROTU] ? new Date(r[COLS_PRZES.DATA_ZWROTU]).getTime() || 0 : 0;
    return {
      idOp: String(r[COLS_PRZES.ID_OPERACJI]),
      dataWydania: formatDate(r[COLS_PRZES.DATA_WYDANIA]),
      osoba: String(r[COLS_PRZES.OSOBA] || ''),
      nazwaSys: ns,
      nazwaWys: sysToWys[ns] || ns,
      sn: String(r[COLS_PRZES.SN] || ''),
      ilosc: Number(r[COLS_PRZES.ILOSC]) || 1,
      kategoria: String(r[COLS_PRZES.KATEGORIA] || ''),
      status: String(r[COLS_PRZES.STATUS] || ''),
      zdjecieWydanieUrl: String(r[COLS_PRZES.ZDJECIE_WYDANIE_URL] || ''),
      zdjecieZwrotUrl: String(r[COLS_PRZES.ZDJECIE_ZWROT_URL] || ''),
      dataZwrotu: formatDate(r[COLS_PRZES.DATA_ZWROTU]),
      opisUszkodzenia: String(r[COLS_PRZES.OPIS_USZKODZENIA] || ''),
      operator: String(r[COLS_PRZES.OPERATOR] || ''),
      _sortDate: rawZwr || rawWyd
    };
  }).sort(function (a, b) {
    return b._sortDate - a._sortDate;
  });
}

// ============================================
// RAPORTY
// ============================================

function getSummaryData() {
  var katalog = getKatalog();
  var przesSheet = getSheet(SHEET_PRZESUNIECIA);
  var lastRow = przesSheet.getLastRow();
  var przesData = lastRow >= 2 ? przesSheet.getRange(2, 1, lastRow - 1, 13).getValues() : [];

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

  var data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  var result = [];

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (String(r[COLS_PRZES.STATUS]) !== 'Uszkodzone') continue;
    result.push({
      idOp: String(r[COLS_PRZES.ID_OPERACJI]),
      nazwaSys: String(r[COLS_PRZES.NAZWA_SYSTEMOWA] || ''),
      sn: String(r[COLS_PRZES.SN] || ''),
      opisUszkodzenia: String(r[COLS_PRZES.OPIS_USZKODZENIA] || ''),
      data: formatDate(r[COLS_PRZES.DATA_WYDANIA]),
      ilosc: Number(r[COLS_PRZES.ILOSC]) || 1,
      osoba: String(r[COLS_PRZES.OSOBA] || '')
    });
  }

  return result.sort(function (a, b) {
    return (b.data || '').localeCompare(a.data || '');
  });
}

function edytujOpisUszkodzenia(idOp, nowyOpis) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_PRZESUNIECIA);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Brak danych' };

    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === idOp) {
        sheet.getRange(i + 2, COLS_PRZES.OPIS_USZKODZENIA + 1).setValue(nowyOpis);
        return { success: true };
      }
    }
    return { success: false, error: 'Nie znaleziono operacji' };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function przeniesNaOsobe(ids, nowaOsoba, operator) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_PRZESUNIECIA);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Brak danych' };

    var idArr = Array.isArray(ids) ? ids : [ids];
    var idSet = {};
    for (var k = 0; k < idArr.length; k++) idSet[String(idArr[k])] = true;

    var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var count = 0;
    for (var i = 0; i < data.length; i++) {
      var opId = String(data[i][COLS_PRZES.ID_OPERACJI]);
      if (idSet[opId] && String(data[i][COLS_PRZES.STATUS]) === 'Wydane') {
        sheet.getRange(i + 2, COLS_PRZES.OSOBA + 1).setValue(nowaOsoba);
        if (operator) sheet.getRange(i + 2, COLS_PRZES.OPERATOR + 1).setValue(operator);
        count++;
      }
    }
    return { success: true, count: count };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function edytujIloscUszkodzenia(idOp, nowaIlosc) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_PRZESUNIECIA);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Brak danych' };

    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === idOp) {
        sheet.getRange(i + 2, COLS_PRZES.ILOSC + 1).setValue(nowaIlosc);
        return { success: true };
      }
    }
    return { success: false, error: 'Nie znaleziono operacji' };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function przeniesUszkodzone(idOp, nowaOsoba, operator) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_PRZESUNIECIA);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Brak danych' };

    var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][COLS_PRZES.ID_OPERACJI]) === idOp && String(data[i][COLS_PRZES.STATUS]) === 'Uszkodzone') {
        sheet.getRange(i + 2, COLS_PRZES.OSOBA + 1).setValue(nowaOsoba);
        if (operator) sheet.getRange(i + 2, COLS_PRZES.OPERATOR + 1).setValue(operator);
        return { success: true };
      }
    }
    return { success: false, error: 'Nie znaleziono operacji' };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}
