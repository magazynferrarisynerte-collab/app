// ============================================
// INWENTARYZACJA - ROZPOZNAWANIE NARZĘDZI (GEMINI)
// ============================================

// Dekodowanie klucza (XOR + charCode)
function _dk(d) {
  var r = '';
  for (var i = 0; i < d.length; i++) r += String.fromCharCode(d[i] ^ ((i % 7) + 3));
  return r;
}

// Klucze Gemini (XOR-encoded) — rotacja round-robin
var _GKEYS = [
  [66,77,127,103,84,113,77,114,74,93,86,93,123,108,100,67,83,69,115,125,58,49,49,111,116,104,66,101,59,85,78,110,118,101,101,89,52,106,101],
  [66,77,127,103,84,113,77,51,54,82,71,64,126,97,105,102,100,114,54,82,112,100,83,118,119,96,75,123,98,110,104,78,79,123,64,75,125,117,67]
];

function _getGeminiKey() {
  var props = PropertiesService.getScriptProperties();
  var idx = parseInt(props.getProperty('GEMINI_KEY_IDX') || '0', 10);
  var key = _dk(_GKEYS[idx % _GKEYS.length]);
  props.setProperty('GEMINI_KEY_IDX', String((idx + 1) % _GKEYS.length));
  return key;
}

// Krok 1: Zapis zdjęcia na Dysk (osobno od analizy — retry nie tworzy duplikatów)
function invSavePhoto(base64) {
  var rootFolders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  var root = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
  var sub = getOrCreateFolder(root, 'Inwentaryzacja');
  var ts = Utilities.formatDate(new Date(), 'Europe/Warsaw', 'yyyyMMdd_HHmmss_') + Math.random().toString(36).substr(2, 4);
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg', 'INV_' + ts + '.jpg');
  var file = sub.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { driveUrl: file.getUrl(), driveFileId: file.getId() };
}

// Krok 2: Analiza Gemini + zapis do arkusza (można bezpiecznie ponawiać)
function invAnalyze(base64, driveFileId, driveUrl) {
  var apiKey = _getGeminiKey();

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=' + apiKey;

  var prompt = 'Jesteś ekspertem od narzędzi budowlanych, warsztatowych i ręcznych. '
    + 'Przeanalizuj zdjęcie BARDZO DOKŁADNIE. Masz DWA zadania:\n'
    + 'ZADANIE 1 — IDENTYFIKACJA I LICZENIE NARZĘDZI:\n'
    + 'Zidentyfikuj KAŻDE narzędzie. Policz PRECYZYJNIE ile sztuk każdego widzisz. '
    + 'Różne rozmiary = osobne wpisy (np. klucz 13mm qty=1, klucz 15mm qty=1). '
    + 'Identyczne = jeden wpis z qty. Opis po polsku z rozmiarem/modelem.\n'
    + 'ZADANIE 2 — ODCZYT KODÓW I OZNACZEŃ (PRIORYTET!):\n'
    + 'Odczytaj WSZYSTKIE widoczne teksty, kody, numery, oznaczenia: '
    + 'napisy markerem, naklejki, numery seryjne (SN), '
    + 'kody inwentarzowe, numery katalogowe, nazwy producenta, '
    + 'numery modeli, KAŻDY tekst widoczny na narzędziu lub etykiecie. '
    + 'Nawet jeśli tekst jest częściowo nieczytelny — brakujące/nieczytelne znaki zastąp znakiem ?. '
    + 'Każdy odczytany tekst = osobny element w tablicy "texts". '
    + 'Ignoruj napis "synerte" — nie dodawaj go do wyników.\n'
    + 'Odpowiedz WYŁĄCZNIE czystym JSON (bez markdown, bez komentarzy, bez tekstu przed/po): '
    + '{"items":[{"name":"Klucz płaski 13mm","qty":3}],"texts":["ABC-123","SN: XY456","Bosch"]}. '
    + 'Jeśli nie widzisz narzędzi: {"items":[],"texts":[]}.';

  var payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: 'image/jpeg', data: base64 } }
      ]
    }],
    generationConfig: { temperature: 0.2, maxOutputTokens: 1024 }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response, code, body;
  for (var attempt = 0; attempt < 3; attempt++) {
    response = UrlFetchApp.fetch(url, options);
    code = response.getResponseCode();
    body = response.getContentText();
    if (code === 429) {
      Utilities.sleep(2000 * (attempt + 1));
      continue;
    }
    break;
  }

  if (code !== 200) {
    Logger.log('Gemini API error: ' + code + ' ' + body);
    throw new Error(code === 429 ? 'Zbyt wiele zapytań — odczekaj chwilę i spróbuj ponownie' : 'Błąd Gemini API (' + code + ')');
  }

  var json = JSON.parse(body);
  var text = '';
  try {
    text = json.candidates[0].content.parts[0].text;
  } catch (e) {
    throw new Error('Nieprawidłowa odpowiedź z Gemini');
  }

  var cleaned = text.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
  var jsonMatch = cleaned.match(/\{[\s\S]*\}/);
  if (!jsonMatch) throw new Error('Gemini nie zwrócił prawidłowego JSON');

  var jsonStr = jsonMatch[0].replace(/,\s*([}\]])/g, '$1');
  var result;
  try {
    result = JSON.parse(jsonStr);
  } catch (e) {
    Logger.log('JSON parse error: ' + e.message + ' | raw: ' + jsonStr.substring(0, 500));
    throw new Error('Gemini zwrócił nieprawidłowy JSON');
  }

  var items = result.items || [];
  var texts = result.texts || [];
  var descText = '';
  if (items.length) {
    descText = items.map(function (it) { return (it.qty > 1 ? it.qty + 'x ' : '') + it.name; }).join(', ');
    if (texts.length) descText += ' | Oznaczenia: ' + texts.join(', ');
    if (driveFileId) {
      try { DriveApp.getFileById(driveFileId).setDescription(descText); } catch (e) { }
    }
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var invSheet = ss.getSheetByName(SHEET_INWENTARYZACJA);
    if (!invSheet) {
      invSheet = ss.insertSheet(SHEET_INWENTARYZACJA);
      invSheet.appendRow(['Timestamp', 'Sesja', 'DriveURL', 'DriveFileID', 'Items', 'Texts', 'Opis']);
    }
    var sesja = Utilities.formatDate(new Date(), 'Europe/Warsaw', 'yyyy-MM-dd');
    invSheet.appendRow([
      new Date(),
      sesja,
      driveUrl || '',
      driveFileId || '',
      JSON.stringify(items),
      JSON.stringify(texts),
      descText
    ]);
  } catch (e) { Logger.log('Inv sheet save error: ' + e.toString()); }

  return {
    items: items,
    texts: texts,
    driveUrl: driveUrl || '',
    driveFileId: driveFileId || ''
  };
}

// ============================================
// INWENTARYZACJA — aktualizacja opisu
// ============================================

function updateInwentaryzacja(driveFileId, items, texts) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_INWENTARYZACJA);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Brak danych inwentaryzacji');
  var data = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === driveFileId) {
      var row = i + 2;
      var opis = items.map(function (it) { return (it.qty > 1 ? it.qty + 'x ' : '') + it.name; }).join(', ');
      if (texts.length) opis += ' | Oznaczenia: ' + texts.join(', ');
      sheet.getRange(row, 5).setValue(JSON.stringify(items));
      sheet.getRange(row, 6).setValue(JSON.stringify(texts));
      sheet.getRange(row, 7).setValue(opis);
      try { DriveApp.getFileById(driveFileId).setDescription(opis); } catch (e) { }
      return { success: true };
    }
  }
  throw new Error('Nie znaleziono skanu');
}

// ============================================
// INWENTARYZACJA — zapis do Dostawa / Wyniki / Braki
// ============================================

function _invGetSheet(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var headers = SCHEMA[name];
    if (headers) { sheet.appendRow(headers); sheet.setFrozenRows(1); sheet.getRange('1:1').setFontWeight('bold'); }
  }
  return sheet;
}

function _invSesja() {
  return Utilities.formatDate(new Date(), 'Europe/Warsaw', 'yyyy-MM-dd');
}

function invSaveDostawa(osoba, typ, opis, driveFileId, itemsJson) {
  var sheet = _invGetSheet(SHEET_INV_DOSTAWA);
  sheet.appendRow([new Date(), _invSesja(), osoba || 'Magazyn', typ || 'skan', opis || '', driveFileId || '', itemsJson || '']);
  return { success: true };
}

function invSaveWynik(osoba, nazwaSys, nazwaWys, sn, ilosc, aiNazwa) {
  var sheet = _invGetSheet(SHEET_INV_WYNIKI);
  var sesja = _invSesja();
  // Check for duplicate — update if exists
  if (sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 3).getValues(); // Sesja, Osoba, Nazwa_Systemowa
    for (var i = 0; i < data.length; i++) {
      var rowSesja = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], 'Europe/Warsaw', 'yyyy-MM-dd') : String(data[i][0]);
      if (rowSesja === sesja && String(data[i][2]) === nazwaSys) {
        // Update qty and timestamp
        var row = i + 2;
        sheet.getRange(row, 1).setValue(new Date());
        sheet.getRange(row, 7).setValue(ilosc);
        sheet.getRange(row, 8).setValue(aiNazwa || '');
        return { success: true, updated: true };
      }
    }
  }
  sheet.appendRow([new Date(), sesja, osoba || 'Magazyn', nazwaSys, nazwaWys || nazwaSys, sn || '', ilosc || 1, aiNazwa || '']);
  return { success: true };
}

function invSaveBrak(osoba, nazwaAI, ilosc) {
  var sheet = _invGetSheet(SHEET_INV_BRAKI);
  var sesja = _invSesja();
  // Check for duplicate — accumulate qty
  if (sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 3).getValues(); // Sesja, Osoba, Nazwa_AI
    for (var i = 0; i < data.length; i++) {
      var rowSesja = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], 'Europe/Warsaw', 'yyyy-MM-dd') : String(data[i][0]);
      if (rowSesja === sesja && String(data[i][2]).toLowerCase() === (nazwaAI || '').toLowerCase()) {
        var row = i + 2;
        var oldQty = Number(sheet.getRange(row, 5).getValue()) || 0;
        sheet.getRange(row, 1).setValue(new Date());
        sheet.getRange(row, 5).setValue(oldQty + (ilosc || 1));
        return { success: true, updated: true };
      }
    }
  }
  sheet.appendRow([new Date(), sesja, osoba || 'Magazyn', nazwaAI, ilosc || 1]);
  return { success: true };
}

function invClearSheets(sesja) {
  var sheets = [SHEET_INV_DOSTAWA, SHEET_INV_WYNIKI, SHEET_INV_BRAKI];
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  for (var s = 0; s < sheets.length; s++) {
    var sheet = ss.getSheetByName(sheets[s]);
    if (!sheet || sheet.getLastRow() < 2) continue;
    var data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      var val = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], 'Europe/Warsaw', 'yyyy-MM-dd') : String(data[i][0]);
      if (val === sesja) sheet.deleteRow(i + 2);
    }
  }
  return { success: true };
}

// ============================================
// INWENTARYZACJA — AI matching braków z katalogiem
// ============================================

function invAiMatch(brakiNames) {
  if (!brakiNames || !brakiNames.length) return [];

  // Build compact catalog list
  var katalog = getKatalog();
  var katalogList = [];
  for (var i = 0; i < katalog.length; i++) {
    var k = katalog[i];
    var entry = k.nazwaWys || k.nazwaSys;
    if (k.sn) entry += ' [SN:' + k.sn + ']';
    katalogList.push({ idx: i, text: entry, nazwaSys: k.nazwaSys, nazwaWys: k.nazwaWys || k.nazwaSys, sn: k.sn || '' });
  }

  // Keep catalog compact — max 300 entries for token budget
  var catalogStr = katalogList.slice(0, 300).map(function (c, i) { return i + ':' + c.text; }).join('\n');

  var prompt = 'Jesteś asystentem magazynu narzędzi.\n'
    + 'KATALOG (id:nazwa):\n' + catalogStr + '\n\n'
    + 'NIEZNALEZIONE POZYCJE (z rozpoznawania AI ze zdjęcia):\n'
    + brakiNames.map(function (b, i) { return i + ':' + b; }).join('\n') + '\n\n'
    + 'Dla każdej nieznalezionej pozycji znajdź NAJLEPSZE dopasowanie z katalogu. '
    + 'Jeśli nie ma dobrego dopasowania — zwróć -1.\n'
    + 'Odpowiedz WYŁĄCZNIE JSON (bez markdown): [{"brak":0,"katalog":15},{"brak":1,"katalog":-1}]\n'
    + 'Gdzie "brak" = indeks pozycji, "katalog" = indeks z katalogu (-1 = brak dopasowania).';

  var apiKey = _getGeminiKey();
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=' + apiKey;

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.1, maxOutputTokens: 512 }
  };

  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) return [];

  var json = JSON.parse(response.getContentText());
  var text = '';
  try { text = json.candidates[0].content.parts[0].text; } catch (e) { return []; }

  var cleaned = text.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
  var arrMatch = cleaned.match(/\[[\s\S]*\]/);
  if (!arrMatch) return [];

  try {
    var matches = JSON.parse(arrMatch[0]);
    var result = [];
    for (var i = 0; i < matches.length; i++) {
      var m = matches[i];
      var bIdx = m.brak, kIdx = m.katalog;
      if (kIdx < 0 || kIdx >= katalogList.length) continue;
      var kat = katalogList[kIdx];
      result.push({
        brakIdx: bIdx,
        brakName: brakiNames[bIdx] || '',
        nazwaSys: kat.nazwaSys,
        nazwaWys: kat.nazwaWys,
        sn: kat.sn
      });
    }
    return result;
  } catch (e) {
    return [];
  }
}

// ============================================
// INWENTARYZACJA — odczyt / czyszczenie
// ============================================

function getInwentaryzacja(sesja) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_INWENTARYZACJA);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var rawSesja = data[i][1];
    var rowSesja = rawSesja instanceof Date ? Utilities.formatDate(rawSesja, 'Europe/Warsaw', 'yyyy-MM-dd') : String(rawSesja);
    if (sesja && rowSesja !== sesja) continue;
    var items = [], texts = [];
    try { items = JSON.parse(data[i][4]); } catch (e) { }
    try { texts = JSON.parse(data[i][5]); } catch (e) { }
    result.push({
      timestamp: data[i][0] ? Utilities.formatDate(new Date(data[i][0]), 'Europe/Warsaw', 'HH:mm:ss') : '',
      sesja: rowSesja,
      driveUrl: String(data[i][2] || ''),
      driveFileId: String(data[i][3] || ''),
      items: items,
      texts: texts,
      opis: String(data[i][6] || '')
    });
  }
  return result;
}

function clearInwentaryzacja(sesja) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_INWENTARYZACJA);
  if (!sheet || sheet.getLastRow() < 2) return { success: true };
  if (!sesja) {
    // Wyczyść wszystko poza nagłówkiem
    sheet.deleteRows(2, sheet.getLastRow() - 1);
    return { success: true };
  }
  var data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    var raw = data[i][0];
    var val = raw instanceof Date ? Utilities.formatDate(raw, 'Europe/Warsaw', 'yyyy-MM-dd') : String(raw);
    if (val === sesja) sheet.deleteRow(i + 2);
  }
  return { success: true };
}
