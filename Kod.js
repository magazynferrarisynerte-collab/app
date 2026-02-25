// ============================================
// MAGAZYN / WYPOŻYCZALNIA - BACKEND v3.0
// Nowy model: Katalog + Przesunięcia
// ============================================

const CACHE_TTL = 60 * 5;
const SPREADSHEET_ID = '1bGjJ4NYfrdtKcqX2GIWIheo_KQQHdhH6keDurrZEHxM';

const SHEET_KATALOG = "Katalog";
const SHEET_OSOBY = "Osoby";
const SHEET_PRZESUNIECIA = "Przesunięcia";
const SHEET_INWENTARYZACJA = "Inwentaryzacja";
const SHEET_INV_DOSTAWA = "Inv_Dostawa";
const SHEET_INV_WYNIKI = "Inv_Wyniki";
const SHEET_INV_BRAKI = "Inv_Braki";
const DRIVE_FOLDER_NAME = "Magazyn_Zdjecia";

const COLS_KATALOG = {
  ID: 0,
  NAZWA_SYSTEMOWA: 1,
  NAZWA_WYSWIETLANA: 2,
  KATEGORIA: 3,
  SN: 4,
  STAN_POCZATKOWY: 5,
  AKTUALNIE_NA_STANIE: 6,
  FLAGA: 7,
  TAGI: 8
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
  OPIS_USZKODZENIA: 11,
  OPERATOR: 12
};

// ============================================
// SETUP — uruchom raz z edytora skryptów
// ============================================

function createNewDatabase() {
  var ss = SpreadsheetApp.create('Magazyn Wypożyczalnia - Baza');
  var id = ss.getId();

  // Katalog
  var kat = ss.getSheets()[0];
  kat.setName(SHEET_KATALOG);
  kat.appendRow(['ID', 'Nazwa Systemowa', 'Nazwa Wyświetlana', 'Kategoria', 'S/N', 'Stan Początkowy', 'Aktualnie Na Stanie', 'Flaga', 'Tagi']);
  kat.setFrozenRows(1);
  kat.getRange('1:1').setFontWeight('bold');

  // Osoby
  var os = ss.insertSheet(SHEET_OSOBY);
  os.appendRow(['ID', 'Imię i Nazwisko', 'Telefon']);
  os.setFrozenRows(1);
  os.getRange('1:1').setFontWeight('bold');

  // Przesunięcia
  var prz = ss.insertSheet(SHEET_PRZESUNIECIA);
  prz.appendRow(['ID Operacji', 'Data Wydania', 'Osoba', 'Nazwa Systemowa', 'S/N', 'Ilość', 'Kategoria', 'Status', 'Zdjęcie Wydanie URL', 'Zdjęcie Zwrot URL', 'Data Zwrotu', 'Opis Uszkodzenia', 'Operator']);
  prz.setFrozenRows(1);
  prz.getRange('1:1').setFontWeight('bold');

  // Inwentaryzacja
  ss.insertSheet(SHEET_INWENTARYZACJA);
  ss.insertSheet(SHEET_INV_DOSTAWA);
  ss.insertSheet(SHEET_INV_WYNIKI);
  ss.insertSheet(SHEET_INV_BRAKI);

  // LocalBackup
  var bk = ss.insertSheet(SHEET_LOCAL_BACKUP);
  bk.appendRow(['Timestamp', 'Operator', 'Type', 'Data_JSON']);
  bk.setFrozenRows(1);
  bk.getRange('1:1').setFontWeight('bold');

  Logger.log('=== NOWY ARKUSZ UTWORZONY ===');
  Logger.log('ID: ' + id);
  Logger.log('URL: ' + ss.getUrl());
  Logger.log('Skopiuj ID powyżej i wklej do SPREADSHEET_ID w Kod.js');

  return { id: id, url: ss.getUrl() };
}

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
    rows.push([generateId('OS'), name, tel]);
    added++;
  }

  if (rows.length) {
    sheet.getRange(lastRow + 1, 1, rows.length, 3).setValues(rows);
  }

  CacheService.getScriptCache().remove("osoby");
  return { success: true, added: added, skipped: skipped };
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

  var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  var katalog = data.map(function (r) {
    var sn = extractSN(r[COLS_KATALOG.SN]);
    var kategoria = String(r[COLS_KATALOG.KATEGORIA] || '').trim().toUpperCase();
    if (sn) kategoria = 'E';
    else if (kategoria !== 'E' && kategoria !== 'Z') kategoria = 'N';
    var flagVal = r[COLS_KATALOG.FLAGA];
    return {
      id: String(r[COLS_KATALOG.ID]),
      nazwaSys: String(r[COLS_KATALOG.NAZWA_SYSTEMOWA] || ''),
      nazwaWys: String(r[COLS_KATALOG.NAZWA_WYSWIETLANA] || ''),
      kategoria: kategoria,
      sn: sn,
      stanPoczatkowy: Number(r[COLS_KATALOG.STAN_POCZATKOWY]) || 0,
      aktualnieNaStanie: Number(r[COLS_KATALOG.AKTUALNIE_NA_STANIE]) || 0,
      flaga: flagVal === true || flagVal === 1 || String(flagVal) === '1',
      tagi: String(r[COLS_KATALOG.TAGI] || '')
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
        flaga: false,
        totalStock: 0,
        totalPoczatkowy: 0,
        items: []
      };
    }
    if (k.flaga) groups[key].flaga = true;
    groups[key].totalStock += k.aktualnieNaStanie;
    groups[key].totalPoczatkowy += k.stanPoczatkowy;
    groups[key].items.push({
      id: k.id,
      nazwaSys: k.nazwaSys,
      sn: k.sn,
      aktualnieNaStanie: k.aktualnieNaStanie,
      flaga: k.flaga
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
    sheet.appendRow([id, nazwaSys.trim(), nazwaWys.trim(), kategoria, sn || '', qty, qty, 0]);
    CacheService.getScriptCache().remove("katalog");
    return { success: true, id: id };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function changeKategoria(idKatalog, nowaKategoria) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_KATALOG);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Katalog pusty' };

    var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][COLS_KATALOG.ID]) === idKatalog) {
        sheet.getRange(i + 2, COLS_KATALOG.KATEGORIA + 1).setValue(nowaKategoria);
        CacheService.getScriptCache().remove("katalog");
        return { success: true };
      }
    }
    return { success: false, error: 'Nie znaleziono pozycji' };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function toggleFlaga(idKatalog) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_KATALOG);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Katalog pusty' };

    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === idKatalog) {
        var cell = sheet.getRange(i + 2, COLS_KATALOG.FLAGA + 1);
        var current = cell.getValue();
        var newVal = (current === 1 || current === true || String(current) === '1') ? 0 : 1;
        cell.setValue(newVal);
        CacheService.getScriptCache().remove("katalog");
        return { success: true, flaga: newVal === 1 };
      }
    }
    return { success: false, error: 'Nie znaleziono pozycji' };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function saveTagi(idKatalog, tagi) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_KATALOG);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Katalog pusty' };
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === idKatalog) {
        sheet.getRange(i + 2, COLS_KATALOG.TAGI + 1).setValue(tagi);
        CacheService.getScriptCache().remove("katalog");
        return { success: true };
      }
    }
    return { success: false, error: 'Nie znaleziono pozycji' };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function autoTagKatalog() {
  var TAG_RULES = [
    // === KLUCZE ===
    { match: /KLUCZ\s+P[ŁL]ASKO/i, tags: 'klucz, klucz płasko-oczkowy, klucz oczkowy, klucz płaski, oczko, płaski, widlak' },
    { match: /KLUCZ.*GRZECHOTK/i, tags: 'klucz, grzechotka, ratchet, klucz z grzechotką, trzeszczotka' },
    { match: /KLUCZ.*DYNAMOMETR|TORQUE\s*WRENCH/i, tags: 'klucz, klucz dynamometryczny, dynamometr, moment, momentowy, torque, dokręcanie' },
    { match: /KLUCZ.*PODBIJAN|RING\s*IMPACT/i, tags: 'klucz, klucz do podbijania, klucz udarowy, podbijak, ring impact' },
    { match: /KLUCZ.*IMBUSOW|HEX\s*KEY|IMBUS/i, tags: 'klucz, imbus, hex, allen, klucz imbusowy, szesciokatny, sześciokątny' },
    { match: /KLUCZ.*TORX|KOMPLET\s+KLUCZY\s+TORX/i, tags: 'klucz, torx, klucz torx, gwiazdka' },
    { match: /KLUCZ\s+Z\s+GRZECHOTK|RATCHET\s*WRENCH/i, tags: 'klucz, grzechotka, ratchet, pokrętło, trzeszczotka' },
    { match: /KLUCZ/i, tags: 'klucz' },
    // === NASADKI ===
    { match: /NASADK.*UDAROW|IMPACT\s*SOCKET/i, tags: 'nasadka, nasadka udarowa, klucz, socket' },
    { match: /NASADK.*IMBUS/i, tags: 'nasadka, imbus, hex, klucz' },
    { match: /NASADK.*TORX/i, tags: 'nasadka, torx, klucz' },
    { match: /NASADEK|NASADK/i, tags: 'nasadka, klucz nasadowy, socket' },
    { match: /PRZEDŁUŻKA|EXTENSION\s*BAR/i, tags: 'przedłużka, adapter, nasadka, klucz' },
    { match: /REDUKCJA\s+ADAPTER/i, tags: 'redukcja, adapter, nasadka, przejściówka' },
    // === ZAKRĘTARKI ===
    { match: /ZAKR[ĘE]TARK.*K[ĄA]TOW/i, tags: 'zakrętarka, zakrętarka kątowa, klucz udarowy, impact' },
    { match: /ZAKR[ĘE]TARK/i, tags: 'zakrętarka, klucz udarowy, impact wrench, dokręcanie, maszynka, pistolecik' },
    // === WKRĘTAKI / ŚRUBOKRĘTY / BITY ===
    { match: /WKR[ĘE]TAK|SCREWDRIVER\b/i, tags: 'wkrętak, śrubokręt, screwdriver' },
    { match: /[ŚS]RUBOKR[ĘE]T/i, tags: 'śrubokręt, wkrętak, screwdriver' },
    { match: /BIT\s+IMBUS|HEX\s*BIT/i, tags: 'bit, imbus, hex, wkrętak, końcówka' },
    { match: /ADAPTER\s+DO\s+BIT/i, tags: 'adapter, bit, wkrętak, wkrętarka' },
    // === WIERTŁA ===
    { match: /WIERT[ŁL]O\s+DO\s+MET|DRILL\s+BIT\s+FOR\s+METAL/i, tags: 'wiertło, wiertło do metalu, drill bit, metal' },
    { match: /WIERT[ŁL]O.*KOBALT|COBALT\s*DRILL/i, tags: 'wiertło, kobalt, wiertło do metalu, drill bit, inox' },
    { match: /WIERT[ŁL]O\s+SDS\s*MAX/i, tags: 'wiertło, sds max, wiertło udarowe, beton' },
    { match: /WIERT[ŁL]O\s+SDS\+/i, tags: 'wiertło, sds+, sds plus, wiertło udarowe, beton' },
    { match: /WIERT[ŁL]O\s+TREP|ANNULAR\s*CUTTER/i, tags: 'wiertło, wiertło trepanacyjne, koronka, otwornica, annular cutter' },
    { match: /WIERT[ŁL]/i, tags: 'wiertło, drill bit' },
    { match: /OTWORNIC/i, tags: 'otwornica, koronka, wiertło, hole saw' },
    { match: /KOMPLET\s+WIERTE[ŁL]/i, tags: 'wiertło, komplet, zestaw wierteł, drill bit set' },
    // === WIERTARKI ===
    { match: /WIERTARK.*MAGNETYCZN/i, tags: 'wiertarka, wiertarka magnetyczna, mag drill, wiercenie' },
    { match: /WIERTARK.*UDAROW.*SDS|SDS\+.*HIKOKI|SDS\+.*HITACHI|ROTARY\s*HAMMER/i, tags: 'wiertarka, wiertarka udarowa, sds+, młotowiertarka, perforator, młot, kucie' },
    { match: /WIERTARK.*UDAROW/i, tags: 'wiertarka, wiertarka udarowa, drill, impact' },
    { match: /WIERTARK/i, tags: 'wiertarka, drill, wiercenie' },
    // === WKRĘTARKI AKU ===
    { match: /WKR[ĘE]TARK.*AKU/i, tags: 'wkrętarka, wkrętarka akumulatorowa, drill, akumulatorowa, aku' },
    { match: /WKR[ĘE]TARK/i, tags: 'wkrętarka, drill, wiertarko-wkrętarka' },
    // === MŁOTY / MŁOTKI ===
    { match: /M[ŁL]OT\s+UDAROW/i, tags: 'młot, młot udarowy, perforator, sds max, kucie, burzenie, wyburzenie, dłutowanie' },
    { match: /M[ŁL]OTEK.*GUMOW/i, tags: 'młotek, młotek gumowy, guma, rubber hammer, kopyto' },
    { match: /M[ŁL]OTEK.*TEFLON/i, tags: 'młotek, młotek teflonowy, teflon, bezodrzutowy, biały' },
    { match: /M[ŁL]OTEK.*KOWALSKI/i, tags: 'młotek, młotek kowalski, teflon' },
    { match: /M[ŁL]OTEK/i, tags: 'młotek, hammer' },
    // === SZLIFIERKI ===
    { match: /SZLIFIERK.*K[ĄA]TOW.*230/i, tags: 'szlifierka, szlifierka kątowa, flex, kątówka, duża, 230, grinder, cięcie, tarcza' },
    { match: /SZLIFIERK.*K[ĄA]TOW.*125/i, tags: 'szlifierka, szlifierka kątowa, flex, kątówka, mała, 125, grinder, cięcie, tarcza' },
    { match: /SZLIFIERK.*K[ĄA]TOW/i, tags: 'szlifierka, szlifierka kątowa, flex, kątówka, grinder, cięcie, tarcza' },
    { match: /SZLIFIERK.*PROST/i, tags: 'szlifierka, szlifierka prosta, prościutka, trzpień, frez, szlifowanie' },
    { match: /SZLIFIERK.*OSCYLAC/i, tags: 'szlifierka, szlifierka oscylacyjna, delta, szlifowanie' },
    { match: /SZLIFIERK/i, tags: 'szlifierka, grinder, szlifowanie' },
    // === TARCZE ===
    { match: /TARCZ.*CI[ĘE]CI.*INOX|TARCZ.*INOX/i, tags: 'tarcza, tarcza tnąca, inox, nierdzewka, cięcie' },
    { match: /TARCZ.*CI[ĘE]CI|TARCZ.*METAL/i, tags: 'tarcza, tarcza tnąca, cięcie, metal' },
    { match: /TARCZ.*SZLIFOW|LAMELKA/i, tags: 'tarcza, tarcza szlifierska, lamelka, szlifowanie' },
    { match: /TARCZ.*FIBROWY|DYSK\s+FIBROWY/i, tags: 'tarcza, dysk fibrowy, szlifowanie' },
    { match: /TARCZ.*RZEP|MULTIDYSK/i, tags: 'tarcza, rzep, multidysk, szlifowanie' },
    { match: /TARCZ/i, tags: 'tarcza, szlifierka' },
    // === PIŁY ===
    { match: /PI[ŁL]A\s+SZABLAST|RECIPRO/i, tags: 'piła, piła szablasta, szablówka, szablastka, lisica, recipro, cięcie' },
    { match: /PI[ŁL]A\s+TARCZOW/i, tags: 'piła, piła tarczowa, cięcie, metal' },
    { match: /PI[ŁL]A\s+HM/i, tags: 'piła, piła tarczowa, tarcza, cięcie' },
    { match: /BRZESZCZOT/i, tags: 'brzeszczot, piła, cięcie, ostrze' },
    { match: /WYRZYNARK|JIGSAW/i, tags: 'wyrzynarka, jigsaw, piła, cięcie' },
    { match: /UKO[ŚS]NIC/i, tags: 'ukośnica, piła, cięcie, kątowe' },
    // === OŚWIETLENIE ===
    { match: /LAMP.*LED|LED.*LAMP|LEDOW/i, tags: 'lampa, led, halogen, oświetlenie, reflektor, światło, robocza, naświetlacz, budowlana' },
    { match: /LAMP.*HALOGEN|HALOGEN/i, tags: 'halogen, lampa, led, oświetlenie, reflektor, światło, robocza, naświetlacz, budowlana' },
    { match: /TA[ŚS]MA\s+LED/i, tags: 'taśma led, led, oświetlenie, światło, pasek' },
    { match: /LATARK|FLASHLIGHT/i, tags: 'latarka, oświetlenie, led, światło, czołówka, czołowa' },
    // === POMIARY ===
    { match: /MIARA\s+ZWIJAN/i, tags: 'miara, metrówka, taśma miernicza, miarka, pomiar, roleta, zwijana' },
    { match: /POZIOMIC.*LASER/i, tags: 'poziomica, laser, pomiar, poziom' },
    { match: /POZIOMIC.*MASZYNOW/i, tags: 'poziomica, maszynowa, precyzyjna, pomiar, poziom' },
    { match: /POZIOMIC/i, tags: 'poziomica, libella, level, pomiar, poziom' },
    { match: /NIWELATOR.*LASER/i, tags: 'niwelator, laser, pomiar, poziomica, rotacyjny' },
    { match: /NIWELATOR/i, tags: 'niwelator, pomiar, poziomica, geodezja' },
    { match: /TEODOLIT/i, tags: 'teodolit, pomiar, geodezja, kąt' },
    { match: /[ŁL]ATA\s+POMIAR/i, tags: 'łata pomiarowa, pomiar, niwelator, geodezja' },
    { match: /SUWMIARK|CALIPER/i, tags: 'suwmiarka, pomiar, kaliber, caliper' },
    { match: /MULTIMETR/i, tags: 'multimetr, miernik, pomiar, elektryka, tester' },
    { match: /PION\s+MURARS|PION\s+PRECYZ|PLUMB/i, tags: 'pion, pion murarski, pomiar, pion budowlany' },
    { match: /SZNUR.*MURARS|SZNUREK.*TRASER|CHALK\s*LINE/i, tags: 'sznurek, sznur murarski, trasowanie, pomiar' },
    { match: /K[ĄA]TOWNIK/i, tags: 'kątownik, kątomierz, pomiar, kąt, liniał' },
    // === SPAWANIE ===
    { match: /SPAWARK.*SPARTUS|SPAWARK.*SPEEDTEC/i, tags: 'spawarka, spawanie, welder, mma, mig, mag, spaw' },
    { match: /SPAWARK/i, tags: 'spawarka, spawanie, welder, spaw' },
    { match: /PRZYŁBICA\s+SPAWAL|WELDING\s*HELMET/i, tags: 'przyłbica, spawanie, maska spawalnicza, bhp, ochrona' },
    { match: /FARTUCH\s+SPAWAL/i, tags: 'fartuch, spawanie, bhp, ochrona, skórzany' },
    { match: /KAPTUR\s+SPAWAL/i, tags: 'kaptur, spawanie, bhp, ochrona' },
    { match: /R[ĘE]KAW.*SPAWAL/i, tags: 'rękaw, spawanie, bhp, ochrona' },
    { match: /R[ĘE]KAWIC.*SPAWAL|WELDING\s*GLOVES/i, tags: 'rękawice, rękawice spawalnicze, spawanie, bhp, tig' },
    { match: /KOC\s+SPAWAL/i, tags: 'koc spawalniczy, spawanie, ochrona, ognioodporny' },
    { match: /KURTYN.*SPAWAL/i, tags: 'kurtyna spawalnicza, spawanie, ochrona, ekran' },
    { match: /STOJAK.*KURTYN/i, tags: 'stojak, kurtyna spawalnicza, spawanie' },
    { match: /UCHWYT\s+SPAWAL|UCHWYT.*MB/i, tags: 'uchwyt, uchwyt spawalniczy, spawanie, palnik, mig' },
    { match: /UCHWYT\s+TIG/i, tags: 'uchwyt, tig, spawanie' },
    { match: /DYSZA.*MB|MB-\d|MB\d/i, tags: 'dysza, spawanie, mig, części spawarki' },
    { match: /REDUKTOR.*CO2|REDUKTOR.*GAZ/i, tags: 'reduktor, gaz, co2, argon, spawanie' },
    { match: /W[ĄA][ŻZ].*GAZ.*TLEN|WELDING.*HOSE/i, tags: 'wąż, gaz, tlen, acetylen, spawanie' },
    { match: /DRUT.*SPAWAL|WELDING\s*WIRE/i, tags: 'drut, drut spawalniczy, spawanie, sg3' },
    { match: /SUSZARK.*ELEKTROD/i, tags: 'suszarka, elektrody, spawanie' },
    { match: /SPRAY\s+SPAWMIX|ANTYPRZYCZ/i, tags: 'spray, spawanie, antyodpryskowy, spawmix' },
    { match: /[ŚS]RODEK.*WAD\s+SPAWAL/i, tags: 'spray, spawanie, kontrola, wady, penetrant' },
    { match: /K[ĄA]TOWNIK\s+SPAWAL|MAGNETYCZN/i, tags: 'kątownik, kątownik spawalniczy, magnes, spawanie' },
    // === GWINTOWANIE ===
    { match: /GWINTOWNIK\s+MASZYNOW|TAPPING/i, tags: 'gwintownik, gwintownik maszynowy, gwintowanie, gwint' },
    { match: /GWINTOWNIK/i, tags: 'gwintownik, gwintowanie, gwint' },
    { match: /NARZYNK/i, tags: 'narzynka, gwintowanie, gwint' },
    { match: /KOMPLET\s+GWINTOWNIK/i, tags: 'gwintownik, narzynka, gwintowanie, komplet' },
    { match: /UCHWYT\s+DO\s+GWINTOWNIK|UCHWYT\s+DO\s+NARZYN|OPRAWKA\s+DO\s+NARZYN/i, tags: 'oprawka, uchwyt, gwintownik, narzynka, gwintowanie' },
    // === PRZEDŁUŻACZE / ELEKTRYKA ===
    { match: /PRZED[ŁL]U[ŻZ]ACZ.*400V/i, tags: 'przedłużacz, 400v, trójfazowy, elektryka, kabel' },
    { match: /PRZED[ŁL]U[ŻZ]ACZ/i, tags: 'przedłużacz, kabel, elektryka, 230v' },
    { match: /ROZDZIELNIC/i, tags: 'rozdzielnica, rozdzielnia, elektryka, prąd, budowlana' },
    { match: /LISTWA\s+PR[ĄA]DOW/i, tags: 'listwa, listwa prądowa, elektryka, gniazdko' },
    { match: /GNIAZDO\s+PRZEMYS|GNIAZDO.*IP/i, tags: 'gniazdo, wtyczka, elektryka, przemysłowe' },
    { match: /WTYCZKA/i, tags: 'wtyczka, gniazdo, elektryka, złącze' },
    { match: /CZUJNIK\s+TLENU/i, tags: 'czujnik, tlen, detektor, bezpieczeństwo, gaz' },
    // === SZCZYPCE / KOMBINERKI ===
    { match: /KOMBINERKI/i, tags: 'kombinerki, szczypce, pliers, chwytanie, kombi' },
    { match: /SZCZYPCE\s+MORS/i, tags: 'szczypce, morsa, chwytanie, zacisk, lock grip, samozaciskowe' },
    { match: /SZCZYPCE\s+SEGER/i, tags: 'szczypce, segera, seeger, pierścienie, osadcze' },
    { match: /SZCZYPCE/i, tags: 'szczypce, obcęgi, pliers' },
    // === NOŻYCE / NOŻE ===
    { match: /NO[ŻZ]YCE.*BLACH/i, tags: 'nożyce, nożyce do blachy, blacha, cięcie' },
    { match: /NO[ŻZ]YCE.*KABL/i, tags: 'nożyce, nożyce do kabli, kable, elektryka, cięcie' },
    { match: /NO[ŻZ]YCE.*BLACHY\s+FALIST/i, tags: 'nożyce, blacha falista, cięcie' },
    { match: /N[ÓO][ŻZ]\s+TAPICERSK|RETRACTABLE\s*BLADE/i, tags: 'nóż, nóż tapicerski, cutter, ostrze' },
    { match: /OSTRZA\s+DO\s+NO/i, tags: 'ostrze, nóż, cutter, wymienne' },
    // === PODNOSZENIE / TRANSPORT ===
    { match: /WCI[ĄA]GNIK.*[ŁL]A[ŃN]CUCH.*ELEKTR/i, tags: 'wciągnik, elektryczny, łańcuchowy, podnoszenie, dźwig' },
    { match: /WCI[ĄA]GNIK.*D[ŹZ]WIGN|LEVER\s*CHAIN/i, tags: 'wciągnik, dźwigniowy, łańcuchowy, podnoszenie, lewar, żaba, żabka' },
    { match: /WCI[ĄA]GNIK|CHAIN\s*HOIST/i, tags: 'wciągnik, łańcuchowy, podnoszenie, dźwig' },
    { match: /PODNO[ŚS]NIK.*HYDRAUL/i, tags: 'podnośnik, hydrauliczny, lewar, podnoszenie' },
    { match: /ZAWIESIE\s+[ŁL]A[ŃN]CUCH/i, tags: 'zawiesia, łańcuch, podnoszenie, hak' },
    { match: /ZAWIESIE\s+PASOW|ZAWIESIE\s+W[ĘE][ŻZ]OW/i, tags: 'zawiesia, pas, wąż, podnoszenie' },
    { match: /SZEKL|SHACKLE/i, tags: 'szekla, podnoszenie, mocowanie, zawiesia' },
    { match: /HAK\s+[ŁL]ADUNK/i, tags: 'hak, ładunkowy, podnoszenie, widły' },
    { match: /UCHWYT\s+KLAMROW|BEAM\s*CLAMP/i, tags: 'uchwyt klamrowy, podnoszenie, zacisk, belka' },
    { match: /[ŁL]A[ŃN]CUCH|CHAIN\b/i, tags: 'łańcuch, podnoszenie' },
    { match: /ROLKI\s+TRANSPORT|DOLLIES|DOLLY/i, tags: 'rolki, rolki transportowe, transport, przesuwanie, wózki, podjazdowe' },
    { match: /W[ÓO]ZEK.*PALETOW|PALLET\s*TRUCK/i, tags: 'wózek, wózek paletowy, paleciak, paleciarz, transport' },
    { match: /W[ÓO]ZEK.*DEMAG|HOIST\s*TROLLEY/i, tags: 'wózek, wózek jezdny, demag, suwnica' },
    { match: /TRANSPORTER.*W[ÓO]ZEK|W[ÓO]Z\s+Z\s+DYSZL/i, tags: 'wózek, transport, przyczepa' },
    { match: /PAS\s+TRANSPORT/i, tags: 'pas transportowy, transport, mocowanie, ratchet, napinacz, sjorka' },
    { match: /[ŻZ]URAW.*WYS[ĄA]G|DZIOBAK/i, tags: 'żuraw, wysięgnik, dziobak, podnoszenie, widły' },
    // === DRABINY / PODESTY ===
    { match: /DRABIN/i, tags: 'drabina, ladder, wysokość' },
    { match: /PODEST/i, tags: 'podest, rusztowanie, wysokość, platforma' },
    // === BHP / OCHRONA ===
    { match: /R[ĘE]KAWIC.*ROBOCZ/i, tags: 'rękawice, rękawice robocze, bhp, ochrona' },
    { match: /R[ĘE]KAWIC/i, tags: 'rękawice, bhp, ochrona' },
    { match: /KASK\b/i, tags: 'kask, bhp, ochrona, głowa' },
    { match: /OKULARY\s+OCHRONN/i, tags: 'okulary, ochronne, bhp, ochrona' },
    { match: /OS[ŁL]ONA\s+TWARZY|PLASTIK\s+DO\s+MASEK/i, tags: 'osłona, twarzy, szybka, bhp, ochrona' },
    { match: /NAKOLANNIK/i, tags: 'nakolanniki, bhp, ochrona, kolana' },
    { match: /KAMIZELK.*OSTRZEG/i, tags: 'kamizelka, odblaskowa, bhp, ochrona' },
    { match: /MASKA.*PRZECIWPY|FFP/i, tags: 'maska, ffp2, przeciwpyłowa, bhp, ochrona' },
    { match: /ZATYCZK.*USZ/i, tags: 'zatyczki, uszy, bhp, ochrona, słuch' },
    { match: /SZELKI\s+BEZPIECZ|FALL\s*PROTECTION/i, tags: 'szelki, uprząż, bhp, praca na wysokości, ochrona, zabezpieczenie, alpinistyczne' },
    { match: /URZĄDZENIE\s+SAMOHAMOWN|SELF.*RETRACT/i, tags: 'urządzenie samohamowne, bhp, praca na wysokości, linka' },
    { match: /LINA\s+ASYKUR|LINA.*[ŻZ]YCIA/i, tags: 'lina, lina asekuracyjna, bhp, praca na wysokości' },
    { match: /KURTKA\s+PRZECIWDESZCZ/i, tags: 'kurtka, deszcz, bhp, ochrona, ubranie' },
    { match: /APTECZK/i, tags: 'apteczka, pierwsza pomoc, bhp' },
    // === PPOŻ ===
    { match: /GA[ŚS]NIC.*CO2/i, tags: 'gaśnica, co2, ppoż, pożar, ogień' },
    { match: /GA[ŚS]NIC.*PIANOW/i, tags: 'gaśnica, pianowa, ppoż, pożar, ogień' },
    { match: /GA[ŚS]NIC.*PROSZKOW/i, tags: 'gaśnica, proszkowa, ppoż, pożar, ogień' },
    { match: /GA[ŚS]NIC/i, tags: 'gaśnica, ppoż, pożar, ogień' },
    { match: /KOC\s+GA[ŚS]NIC/i, tags: 'koc gaśniczy, ppoż, ogień, ochrona' },
    // === PRZECINAK / DŁUTO / BRECHA ===
    { match: /PRZECINAK/i, tags: 'przecinak, dłuto, chisel, kucie' },
    { match: /BRECHA|CROWBAR/i, tags: 'brecha, łom, łapka, crowbar, wyważanie, łamak, gwoździówka' },
    { match: /PUNKTAK/i, tags: 'punktak, wybijak, nakiełek, metal' },
    { match: /KOMPLET\s+WYBIJAK/i, tags: 'wybijak, punktak, drift punch' },
    // === PILNIKI / FREZY ===
    { match: /PILNIK\s+OKR[ĄA]G/i, tags: 'pilnik, pilnik okrągły, piłowanie, metal' },
    { match: /PILNIK\s+P[ŁL]ASK/i, tags: 'pilnik, pilnik płaski, piłowanie, metal' },
    { match: /PILNIK/i, tags: 'pilnik, piłowanie, metal' },
    { match: /FREZ\s+PILNIK/i, tags: 'frez, pilnik, szlifierka prosta, obróbka' },
    // === ŚCISKI / IMADŁO ===
    { match: /IMAD[ŁL]O/i, tags: 'imadło, ścisk, vice, mocowanie' },
    { match: /[ŚS]CISK\s+STOLARS/i, tags: 'ścisk, ścisk stolarski, zacisk, clamp, f-clamp' },
    // === ŚCIĄGACZ ===
    { match: /[ŚS]CI[ĄA]GACZ.*[ŁL]O[ŻZ]YSK/i, tags: 'ściągacz, łożyska, puller, demontaż' },
    { match: /[ŚS]CI[ĄA]GACZ/i, tags: 'ściągacz, puller, demontaż' },
    // === NITOWNICA ===
    { match: /NITOWNIC.*NITONAKR/i, tags: 'nitownica, nitonakrętka, rivnut, nitowanie' },
    { match: /NITOWNIC/i, tags: 'nitownica, nity, riveter, nitowanie' },
    // === ODKURZACZ ===
    { match: /ODKURZACZ/i, tags: 'odkurzacz, vacuum, ssanie, czyszczenie, przemysłowy' },
    // === GIĘTARKA / OBCINACZ RUR ===
    { match: /GI[ĘE]TARK/i, tags: 'giętarka, gięcie, rury, tube bender' },
    { match: /OBCINA[CK]|PIPECUTTER/i, tags: 'obcinacz, obcinak, rury, cięcie rur' },
    { match: /K[ÓO][ŁL]KO\s+TN[ĄA]C/i, tags: 'kółko tnące, obcinacz, rury, rems' },
    { match: /GRATOWNIK/i, tags: 'gratownik, rury, obróbka' },
    // === MALOWANIE ===
    { match: /P[ĘE]DZEL/i, tags: 'pędzel, malowanie, farba' },
    { match: /WA[ŁL]EK\s+MALARS/i, tags: 'wałek, malowanie, farba' },
    { match: /KUWETA\s+MALARS/i, tags: 'kuweta, malowanie, wałek, farba' },
    { match: /ZESTAW\s+DO\s+MALOWA/i, tags: 'malowanie, zestaw, pędzel, wałek, farba' },
    { match: /SPRAY\s+AKRYL|SPRAY\s+ZNAKUJ/i, tags: 'spray, farba, znakowanie, malowanie' },
    // === PISTOLET DO MAS ===
    { match: /PISTOLET.*HILTI|PISTOLET.*BOSTIK|WYCISKACZ/i, tags: 'pistolet, wyciskacz, silikon, klej, masa, bostik' },
    { match: /KOSTKA\s+DO\s+BOSTIK|KO[ŃN]C[ÓO]WKA.*BOSTIK/i, tags: 'bostik, klej, masa, końcówka' },
    // === NAGRZEWNICA ===
    { match: /NAGRZEWNIC/i, tags: 'nagrzewnica, grzanie, ogrzewanie, heater, dmuchawa, ciepło, piecyk' },
    // === WENTYLATOR ===
    { match: /WENTYLATOR/i, tags: 'wentylator, wiatrak, wentylacja, dmuchawa' },
    // === SPRAYE TECHNICZNE ===
    { match: /SPRAY\s+WD|WD-?40/i, tags: 'wd-40, spray, odrdzewiacz, smarowanie' },
    { match: /SPRAY\s+ODRDZ|ROST.OFF/i, tags: 'spray, odrdzewiacz, wd-40, chemia' },
    { match: /SPRAY\s+DO\s+GWINT|REMS/i, tags: 'spray, gwintowanie, rems, smarowanie' },
    { match: /ACETON/i, tags: 'aceton, rozpuszczalnik, czyszczenie, chemia' },
    { match: /ODMRA[ŻZ]ACZ/i, tags: 'odmrażacz, zamek, zima, chemia' },
    // === ŁOPATY / MIOTŁY / CZYSZCZENIE ===
    { match: /[ŁL]OPAT|SHOVEL/i, tags: 'łopata, szpadel, kopanie' },
    { match: /MIOT[ŁL]|SZCZOTK.*ULICZN/i, tags: 'miotła, szczotka, zamiatanie, czyszczenie' },
    { match: /ZMIOTK/i, tags: 'zmiotka, szufelka, czyszczenie' },
    { match: /WIADRO/i, tags: 'wiadro, budowlane, woda' },
    { match: /CZY[ŚS]CIWO/i, tags: 'czyściwo, szmaty, czyszczenie' },
    // === SKRZYNKI NARZĘDZIOWE ===
    { match: /SKRZYNI.*NARZ.*METALOW/i, tags: 'skrzynka, narzędziowa, metalowa, przechowywanie' },
    { match: /SKRZYNI.*NARZ.*SIATK/i, tags: 'skrzynka, narzędziowa, siatkowa, przechowywanie' },
    { match: /SKRZYNI.*NARZ/i, tags: 'skrzynka, narzędziowa, przechowywanie' },
    // === TAŚMY ===
    { match: /TA[ŚS]MA\s+OSTRZEG/i, tags: 'taśma, ostrzegawcza, bhp, oznakowanie' },
    { match: /TA[ŚS]MA.*MALARS|PAINTER/i, tags: 'taśma, taśma malarska, malowanie, maskowanie' },
    { match: /TA[ŚS]MA.*SREBRNA|REINFORCED\s*TAPE/i, tags: 'taśma, taśma srebrna, ducktape, mocna, naprawa' },
    // === LINY ===
    { match: /LINA\s+JUTOW|LINA\s+KONOPL/i, tags: 'lina, jutowa, konopna, sznur' },
    // === BARIERKI / OZNAKOWANIE ===
    { match: /BARIERKA|SŁUPEK.*BIAŁO/i, tags: 'barierka, słupek, oznakowanie, bhp, ogrodzenie' },
    // === KRÓTKOFALÓWKA ===
    { match: /KR[ÓO]TKOFAL[ÓO]WK|MOTOROLA/i, tags: 'krótkofalówka, radio, komunikacja, motorola, walkie talkie, radiotelefon, cb' },
    // === OSTRZAŁKA ===
    { match: /OSTRZA[ŁL]K.*WIERTE/i, tags: 'ostrzałka, wiertła, ostrzenie, regeneracja' },
    // === POGŁĘBIACZ ===
    { match: /POG[ŁL][ĘE]BIAC|GZYMEK/i, tags: 'pogłębiacz, gzymek, stożkowy, fazowanie' },
    // === SMAROWNICA ===
    { match: /SMAROWNIC/i, tags: 'smarownica, smarowanie, towot' },
    // === POMPKA ===
    { match: /POMPK.*DYBLOW/i, tags: 'pompka, dyblowanie, klej, dozowanie' },
    // === OSADZAK ===
    { match: /OSADZAK/i, tags: 'osadzak, kotwy, montaż, beton' },
    // === SZCZOTKI DRUCIANE ===
    { match: /SZCZOTK.*DRUCIAN|WIRE\s*BRUSH/i, tags: 'szczotka, druciana, czyszczenie, rdza, metal' },
    { match: /SZCZOTK.*TARCZOW/i, tags: 'szczotka, tarczowa, szlifierka, czyszczenie, metal' },
    { match: /SZCZOTK.*TRZPIEN|SZCZOTK.*P[ĘE]DZELK/i, tags: 'szczotka, trzpieniowa, szlifierka prosta, czyszczenie' },
    // === NACIĄG ===
    { match: /NACI[ĄA]G\s+DO\s+STRUN/i, tags: 'naciąg, struny, napinanie' },
    // === PAPIER ŚCIERNY ===
    { match: /PAPIER\s+[ŚS]CIERN/i, tags: 'papier ścierny, szlifowanie, ścierny' },
    // === FOLIA ===
    { match: /FOLIA\s+STRETCH/i, tags: 'folia, stretch, pakowanie, owijanie' },
    // === MATERIAŁY BIUROWE ===
    { match: /MATERIA[ŁL]Y\s+BIUROW/i, tags: 'biuro, materiały biurowe, segregator, długopis' },
    { match: /MARKER|CIENKOPIS/i, tags: 'marker, pisak, oznaczanie, cienkopis' },
    { match: /NOTESY/i, tags: 'notes, notatnik, biuro' },
    { match: /DRUKARK/i, tags: 'drukarka, biuro, drukowanie' },
    { match: /TUSZ\s+DRUKARK/i, tags: 'tusz, drukarka, wkład' },
    { match: /LAMINATOR/i, tags: 'laminator, biuro, laminowanie' },
    { match: /TABLICA\s+MAGNETYCZN/i, tags: 'tablica, magnetyczna, biuro' },
    { match: /P[ÓO][ŁL]KA\s+NA\s+DOKUMENT/i, tags: 'półka, dokumenty, biuro' },
    // === AGD ===
    { match: /CZAJNIK/i, tags: 'czajnik, kuchnia, agd' },
    { match: /KUCHENKA\s+MIKROFAL/i, tags: 'kuchenka, mikrofalowa, kuchnia, agd' },
    { match: /LOD[ÓO]WK/i, tags: 'lodówka, kuchnia, agd' },
    // === ELEKTRONIKA ===
    { match: /HUB\s+MULTIPORT/i, tags: 'hub, usb, komputer, adapter' },
    { match: /MYSZ\s+BEZPRZEW/i, tags: 'mysz, komputer, bezprzewodowa' },
    { match: /KABEL\s+USB/i, tags: 'kabel, usb, ładowanie' },
    // === OPASKA ZACISKOWA ===
    { match: /OPASK.*ZACISK/i, tags: 'opaska, zaciskowa, trytytka, mocowanie' },
    // === WKRĘTY ===
    { match: /WKR[ĘE]TY\b/i, tags: 'wkręty, śruby, mocowanie' },
    // === WORKI ===
    { match: /WORKI\s+BIG\s*BAG/i, tags: 'worki, big bag, transport, gruz' },
    { match: /WORKI\s+NA\s+[ŚS]MIEC/i, tags: 'worki, śmieci, sprzątanie' },
    { match: /WORKI\s+DO\s+ODKURZ/i, tags: 'worki, odkurzacz, wymienne' },
    // === K[Ł]ÓDKA ===
    { match: /K[ŁL][ÓO]DK/i, tags: 'kłódka, zamek, zabezpieczenie' },
    // === KANISTER ===
    { match: /KANISTER/i, tags: 'kanister, zbiornik, paliwo' },
    // === MATA ===
    { match: /MATA\s+REMONT/i, tags: 'mata, remontowa, ochrona, podłoga' },
    // === WŁÓKNINA ===
    { match: /W[ŁL][ÓO]KNIN/i, tags: 'włóknina, szlifowanie, ścierny' },
    // === PROFIL DREWNIANY ===
    { match: /PROFIL\s+DREWN/i, tags: 'profil, drewno, łata, budowlany' },
    // === PŁYTA ===
    { match: /P[ŁL]YTA\s+PIL[ŚS]N/i, tags: 'płyta, pilśniowa, budowlana' },
    // === DORNIK ===
    { match: /DORNIK/i, tags: 'dornik, trzpień, kalibrowanie, rury' }
  ];

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(60000);
    var sheet = getSheet(SHEET_KATALOG);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Katalog pusty' };

    var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    var count = 0;

    for (var i = 0; i < data.length; i++) {
      var nazwa = String(data[i][COLS_KATALOG.NAZWA_WYSWIETLANA] || '') + ' ' + String(data[i][COLS_KATALOG.NAZWA_SYSTEMOWA] || '');
      var existing = String(data[i][COLS_KATALOG.TAGI] || '').trim();
      var tagSet = {};

      // Zachowaj istniejące ręczne tagi
      if (existing) {
        existing.split(',').forEach(function (t) {
          var tt = t.trim().toLowerCase();
          if (tt) tagSet[tt] = true;
        });
      }

      // Dopasuj reguły
      for (var r = 0; r < TAG_RULES.length; r++) {
        if (TAG_RULES[r].match.test(nazwa)) {
          TAG_RULES[r].tags.split(',').forEach(function (t) {
            var tt = t.trim().toLowerCase();
            if (tt) tagSet[tt] = true;
          });
        }
      }

      // Dodaj też rozmiary jako tagi (np. "13mm", "1/2")
      var sizes = nazwa.match(/\b(\d+(?:[.,]\d+)?)\s*mm\b/gi);
      if (sizes) {
        for (var s = 0; s < sizes.length; s++) tagSet[sizes[s].toLowerCase().replace(/\s/g, '')] = true;
      }
      var inches = nazwa.match(/\b(\d+\/\d+"?)\b/g);
      if (inches) {
        for (var s = 0; s < inches.length; s++) tagSet[inches[s]] = true;
      }

      var tagiStr = Object.keys(tagSet).sort().join(', ');
      if (tagiStr !== existing) {
        sheet.getRange(i + 2, COLS_KATALOG.TAGI + 1).setValue(tagiStr);
        count++;
      }
    }

    // Najpierw wyczyść nazwy wyświetlane
    var cleanResult = cleanDisplayNames();
    var cleaned = cleanResult.success ? cleanResult.count : 0;

    CacheService.getScriptCache().remove("katalog");
    return { success: true, count: count, total: data.length, cleaned: cleaned };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function cleanDisplayNames() {
  var BRANDS = ['HIKOKI','HITACHI','MAKITA','BOSCH','DEWALT','MILWAUKEE','METABO','STANLEY',
    'HILTI','WURTH','IRWIN','YATO','PROLINE','KARCHER','BRENNENSTUHL','STALIMET',
    'PROJAAHN','SPARTUS','SPEEDTEC','MOTOROLA','BROTHER','ARGO','DEDRA','TOPCON',
    'LAMIGO','PREXISO','SOUTH','NIVEL','SMART','UNICRAFT','RHINO','MIPROMET',
    'ALFRA','METALLKRAFT','DEMAG','VANDER','MASTER','FARAONE','CORMAK','KOWALSKI',
    'MARELD','MEMOBE','REMS','SCELL-IT','KING TONY','M7','EPM','KASTOM',
    'HEFAJSTOS','BORG','SIGMA','NEOS STRADA'];
  var ABBREVS = [
    [/\bMET\.\b/gi, 'METALU'],
    [/\bSZAB\.\b/gi, 'SZABLASTEJ'],
    [/\bDŁU\b/gi, 'DŁUGIE'],
    [/\bNIWELATO\b/gi, 'NIWELATOR'],
    [/\bpcs\b/gi, 'SZT'],
    [/\bI\.T\.P\b/gi, 'ITP'],
    [/\bST\.\b/gi, 'STOPNI'],
    [/\bSZT\.\b/gi, 'SZT'],
    [/\b(\d+)MM\b/gi, '$1 MM'],
    [/\bOBROT\.\b/gi, 'OBROTÓW']
  ];

  function simplify(name) {
    if (!name) return '';
    var s = String(name).trim();
    // Usuń angielski tekst w nawiasach
    s = s.replace(/\s*\([^)]*[A-Z]{3,}[^)]*\)\s*/g, ' ');
    // Popraw skróty
    for (var i = 0; i < ABBREVS.length; i++) {
      s = s.replace(ABBREVS[i][0], ABBREVS[i][1]);
    }
    // Usuń numery modeli (ale zachowaj rozmiary jak 125MM, 230, 1/2'')
    var words = s.split(/\s+/);
    var result = [];
    var brandFound = false;
    for (var i = 0; i < words.length; i++) {
      var w = words[i];
      var wUp = w.toUpperCase().replace(/[,;]/g, '');
      // Zachowaj markę (pierwszą znalezioną)
      var isBrand = false;
      for (var b = 0; b < BRANDS.length; b++) {
        if (wUp === BRANDS[b] || (BRANDS[b].indexOf(' ') === -1 && wUp === BRANDS[b])) {
          isBrand = true; break;
        }
      }
      if (isBrand) { if (!brandFound) { result.push(w); brandFound = true; } continue; }
      // Usuń numery modeli alfanumeryczne (DS18DE, G13SB4YGZ, CR18DBWJZ, WR18DH, MB501D itd.)
      if (/^[A-Z]{1,4}\d{2,}[A-Z]*[-]?\d*[A-Z]*\d*$/i.test(wUp) && wUp.length > 4) continue;
      if (/^[A-Z]{2,}\d+-\w+$/i.test(wUp)) continue;
      // Usuń kody artykułów (30-497, 30-457)
      if (/^\d{2,}-\d{2,}$/.test(wUp)) continue;
      // Usuń puste myślniki
      if (/^[-–—]+$/.test(w)) continue;
      // Usuń kolory (ŻÓŁTA, CZERWONY, NIEBIESKI, CZARNY itp) - opcjonalnie
      // Zachowaj rozmiary, ilości, nazwy
      result.push(w);
    }
    s = result.join(' ').trim();
    // Wyczyść podwójne spacje, trailing commas
    s = s.replace(/\s{2,}/g, ' ').replace(/\s*,\s*$/, '').replace(/\s+\)/, ')').replace(/\(\s+/, '(');
    // Usuwaj puste nawiasy
    s = s.replace(/\(\s*\)/g, '');
    return s;
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(60000);
    var sheet = getSheet(SHEET_KATALOG);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Katalog pusty' };

    var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    var count = 0;
    var changes = [];

    for (var i = 0; i < data.length; i++) {
      var old = String(data[i][COLS_KATALOG.NAZWA_WYSWIETLANA] || '');
      var cleaned = simplify(old);
      if (cleaned && cleaned !== old) {
        sheet.getRange(i + 2, COLS_KATALOG.NAZWA_WYSWIETLANA + 1).setValue(cleaned);
        changes.push(old.substring(0, 40) + ' → ' + cleaned.substring(0, 40));
        count++;
      }
    }

    CacheService.getScriptCache().remove("katalog");
    return { success: true, count: count, total: data.length, examples: changes.slice(0, 15) };
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

      // Determine status
      var status = (item.kategoria === 'Z') ? 'Zuzyte' : 'Wydane';

      // Handle photo (optional for any category)
      var photoUrl = '';
      var opId = generateId('OP');
      if (item.photoBase64) {
        photoUrl = savePhotoToDrive(item.photoBase64, 'Wydania', opId);
      }

      if (resolved) {
        // Decrement stock
        var newStock = resolved.stock - qty;
        katSheet.getRange(resolved.row + 1, COLS_KATALOG.AKTUALNIE_NA_STANIE + 1).setValue(newStock);
        katData[resolved.row][COLS_KATALOG.AKTUALNIE_NA_STANIE] = newStock;

        przesSheet.appendRow([
          opId, new Date(), osobaImie,
          String(resolved.data[COLS_KATALOG.NAZWA_SYSTEMOWA] || ''),
          extractSN(resolved.data[COLS_KATALOG.SN]),
          qty, item.kategoria, status, photoUrl, '', '', '',
          operator || ''
        ]);
      } else {
        // Custom item — nie ma w katalogu, wpisz bezpośrednio
        przesSheet.appendRow([
          opId, new Date(), osobaImie,
          item.nazwaWys,
          item.sn || '',
          qty, item.kategoria || 'N', status, photoUrl, '', '', '',
          operator || ''
        ]);
      }
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

        incrementStock(katSheet, nazwaSys, returnQty);

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
// LOG (PRZESUNIĘCIA)
// ============================================

function getLog() {
  var sheet = getSheet(SHEET_PRZESUNIECIA);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

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
      opisUszkodzenia: String(r[COLS_PRZES.OPIS_USZKODZENIA] || ''),
      operator: String(r[COLS_PRZES.OPERATOR] || '')
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
      idOp: String(r[COLS_PRZES.ID_OPERACJI]),
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

// Jedna funkcja: zapis na Dysk + analiza Gemini (base64 przesyłany raz z klienta)
function invScanAndSave(base64) {
  // 1. Zapisz na Dysk
  var driveUrl = '', driveFileId = '';
  try {
    var rootFolders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
    var root = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
    var sub = getOrCreateFolder(root, 'Inwentaryzacja');
    var ts = Utilities.formatDate(new Date(), 'Europe/Warsaw', 'yyyyMMdd_HHmmss');
    var blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg', 'INV_' + ts + '.jpg');
    var file = sub.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    driveUrl = file.getUrl();
    driveFileId = file.getId();
  } catch (e) {
    Logger.log('Drive save error: ' + e.toString());
  }

  // 2. Analiza Gemini (z tego samego base64 — już w pamięci serwera)
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

  // 3. Aktualizuj opis pliku na Dysku
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

  // 4. Zapisz do arkusza Inwentaryzacja
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
      driveUrl,
      driveFileId,
      JSON.stringify(items),
      JSON.stringify(texts),
      descText
    ]);
  } catch (e) { Logger.log('Inv sheet save error: ' + e.toString()); }

  return {
    items: items,
    texts: texts,
    driveUrl: driveUrl,
    driveFileId: driveFileId
  };
}

// ============================================
// BACKUP
// ============================================

function createBackup() {
  var file = DriveApp.getFileById(SpreadsheetApp.openById(SPREADSHEET_ID).getId());
  var folderId = "19maggYLBrsxFEZRvSPDGdLTSDNgO3bmG";
  var timestamp = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd_HH-mm");
  var backupName = "BACKUP_" + file.getName() + "_" + timestamp;
  var copy = file.makeCopy(backupName);
  var folder = DriveApp.getFolderById(folderId);
  folder.addFile(copy);
  DriveApp.getRootFolder().removeFile(copy);
  return { success: true, name: backupName };
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
  Katalog: ['ID', 'Nazwa_Systemowa', 'Nazwa_Wyswietlana', 'Kategoria', 'SN', 'Stan_Poczatkowy', 'Aktualnie_Na_Stanie', 'Flaga', 'Tagi'],
  Osoby: ['ID', 'Imie', 'Telefon'],
  'Przesunięcia': ['ID_Operacji', 'Data_Wydania', 'Osoba', 'Nazwa_Systemowa', 'SN', 'Ilosc', 'Kategoria', 'Status', 'Zdjecie_Wydanie_URL', 'Zdjecie_Zwrot_URL', 'Data_Zwrotu', 'Opis_Uszkodzenia', 'Operator'],
  Inwentaryzacja: ['Timestamp', 'Sesja', 'DriveURL', 'DriveFileID', 'Items', 'Texts', 'Opis'],
  Inv_Dostawa: ['Timestamp', 'Sesja', 'Osoba', 'Typ', 'Opis', 'DriveFileID', 'Items_JSON'],
  Inv_Wyniki: ['Timestamp', 'Sesja', 'Osoba', 'Nazwa_Systemowa', 'Nazwa_Wyswietlana', 'SN', 'Ilosc', 'AI_Nazwa'],
  Inv_Braki: ['Timestamp', 'Sesja', 'Osoba', 'Nazwa_AI', 'Ilosc']
};

function migrateSchema() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var log = [];
  for (var name in SCHEMA) {
    var headers = SCHEMA[name];
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      sheet.getRange('1:1').setFontWeight('bold');
      log.push('Utworzono: ' + name);
    } else {
      // Sprawdź brakujące kolumny
      var existing = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
      var added = 0;
      for (var i = 0; i < headers.length; i++) {
        var found = false;
        for (var j = 0; j < existing.length; j++) {
          if (String(existing[j]).trim() === headers[i]) { found = true; break; }
        }
        if (!found) {
          var col = existing.length + added + 1;
          sheet.getRange(1, col).setValue(headers[i]).setFontWeight('bold');
          added++;
        }
      }
      if (added) log.push(name + ': dodano ' + added + ' kolumn');
      else log.push(name + ': OK');
    }
  }
  return { success: true, log: log };
}

// ============================================
// CRUD — generyczne operacje na arkuszach
// ============================================

function getSheetData(sheetName, startRow, maxRows) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Brak arkusza: ' + sheetName };
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { success: true, headers: [], rows: [], total: 0 };

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  var total = Math.max(0, lastRow - 1);
  var from = Math.max(2, (startRow || 0) + 2);
  var count = maxRows ? Math.min(maxRows, lastRow - from + 1) : lastRow - from + 1;
  if (count < 1) return { success: true, headers: headers, rows: [], total: total };

  var data = sheet.getRange(from, 1, count, lastCol).getValues();
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    var row = {};
    row._row = from + i; // numer wiersza w arkuszu (1-based)
    for (var j = 0; j < headers.length; j++) {
      var v = data[i][j];
      row[headers[j]] = v instanceof Date ? Utilities.formatDate(v, 'Europe/Warsaw', 'yyyy-MM-dd HH:mm') : String(v != null ? v : '');
    }
    rows.push(row);
  }
  return { success: true, headers: headers, rows: rows, total: total };
}

function insertSheetRow(sheetName, rowData) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Brak arkusza: ' + sheetName };
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var newRow = [];
  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i]);
    newRow.push(rowData[key] != null ? rowData[key] : '');
  }
  sheet.appendRow(newRow);
  CacheService.getScriptCache().removeAll(["katalog", "osoby"]);
  return { success: true };
}

function updateSheetRow(sheetName, rowNum, rowData) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Brak arkusza: ' + sheetName };
  if (rowNum < 2) return { success: false, error: 'Nieprawidłowy wiersz' };
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i]);
    if (rowData[key] !== undefined) {
      sheet.getRange(rowNum, i + 1).setValue(rowData[key]);
    }
  }
  CacheService.getScriptCache().removeAll(["katalog", "osoby"]);
  return { success: true };
}

function deleteSheetRow(sheetName, rowNum) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Brak arkusza: ' + sheetName };
  if (rowNum < 2) return { success: false, error: 'Nie można usunąć nagłówka' };
  if (rowNum > sheet.getLastRow()) return { success: false, error: 'Wiersz nie istnieje' };
  sheet.deleteRow(rowNum);
  CacheService.getScriptCache().removeAll(["katalog", "osoby"]);
  return { success: true };
}

function getSheetNames() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheets().map(function (s) { return s.getName(); });
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
    var kod = String(r[1] || '');              // stary KOD narzędzia
    var nazwa = String(r[2] || '');
    var opis = String(r[3] || '');             // stary OPIS (może zawierać SN)
    var imie = String(r[5] || '');
    var dataWyp = r[7];
    var dataZwr = r[8];
    var ilosc = Number(r[9]) || 1;
    var hasReturn = dataZwr && String(dataZwr).length > 0;

    // Nazwa_Systemowa: oryginalna nazwa
    var nazwaSys = nazwa || kod;

    // Próba wyciągnięcia SN z opisu
    var sn = extractSN(opis);
    // Jeśli extractSN zwróciło cały opis (brak prefiksu sn/s/n), nie traktuj jako SN
    if (sn === opis.trim()) sn = '';

    przesSheet.appendRow([
      idWyp,                          // ID_Operacji (zachowujemy stare ID)
      dataWyp || '',                  // Data_Wydania
      imie,                           // Osoba
      nazwaSys,                       // Nazwa_Systemowa (KOD — NAZWA)
      sn,                             // SN (wyciągnięty z opisu)
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

// Migracja starych Narzędzi → nowy Katalog
// Stary format: KOD | NAZWA | OPIS | ILOSC | JEDNOSTKA | KATEGORIA
function migrateNarzedzia() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var narzSheet = ss.getSheetByName("Narzędzia");
  if (!narzSheet) return { success: false, error: 'Brak arkusza Narzędzia' };

  var katSheet = ss.getSheetByName(SHEET_KATALOG);
  if (!katSheet) return { success: false, error: 'Brak arkusza Katalog' };

  var lastRow = narzSheet.getLastRow();
  if (lastRow < 2) return { success: true, migrated: 0 };

  var data = narzSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var count = 0;

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var kod = String(r[0] || '').trim();
    var nazwa = String(r[1] || '').trim();
    var opis = String(r[2] || '').trim();
    var ilosc = Number(r[3]) || 1;
    var kategoria = String(r[5] || '').trim().toLowerCase();

    if (!kod && !nazwa) continue;

    // Nazwa_Systemowa: oryginalna pełna nazwa
    var nazwaSys = nazwa || kod;
    var nazwaWys = simplifyToolName(nazwa) || nazwa || kod;

    // Wyciągnij SN z opisu
    var sn = extractSN(opis);
    if (sn === opis) sn = '';

    // Normalizuj kategorię do E/N/Z
    if (kategoria === 'zużywalne' || kategoria === 'zuzywalne' || kategoria === 'z') kategoria = 'Z';
    else if (kategoria === 'elektronarzędzia' || kategoria === 'elektronarzedzia' || kategoria === 'specjalne' || kategoria === 'e') kategoria = 'E';
    else if (kategoria === 'stałe' || kategoria === 'stale' || kategoria === 'n') kategoria = 'N';
    else if (sn) kategoria = 'E';
    else kategoria = 'N';

    var id = generateId('K');

    katSheet.appendRow([
      id,           // ID
      nazwaSys,     // Nazwa_Systemowa
      nazwaWys,     // Nazwa_Wyswietlana
      kategoria,    // Kategoria
      sn,           // SN
      ilosc,        // Stan_Poczatkowy
      ilosc         // Aktualnie_Na_Stanie
    ]);
    count++;
  }

  CacheService.getScriptCache().remove("katalog");
  return { success: true, migrated: count };
}

// ============================================
// MIGRACJA KATEGORII (jednorazowa)
// Podmienia stare nazwy kategorii na E/N/Z
// ============================================
function migrateKategorie() {
  var map = {
    'elektronarzedzia': 'E', 'elektronarzędzia': 'E', 'specjalne': 'E',
    'stale': 'N', 'stałe': 'N',
    'zuzywalne': 'Z', 'zużywalne': 'Z'
  };

  var changed = 0;

  // Katalog — kolumna D (index 3)
  var kat = getSheet(SHEET_KATALOG);
  var katLast = kat.getLastRow();
  if (katLast >= 2) {
    var katData = kat.getRange(2, COLS_KATALOG.KATEGORIA + 1, katLast - 1, 1).getValues();
    for (var i = 0; i < katData.length; i++) {
      var v = String(katData[i][0] || '').trim().toLowerCase();
      if (map[v]) {
        kat.getRange(i + 2, COLS_KATALOG.KATEGORIA + 1).setValue(map[v]);
        changed++;
      }
    }
  }

  // Przesunięcia — kolumna G (index 6)
  var prz = getSheet(SHEET_PRZESUNIECIA);
  var przLast = prz.getLastRow();
  if (przLast >= 2) {
    var przData = prz.getRange(2, COLS_PRZES.KATEGORIA + 1, przLast - 1, 1).getValues();
    for (var j = 0; j < przData.length; j++) {
      var w = String(przData[j][0] || '').trim().toLowerCase();
      if (map[w]) {
        prz.getRange(j + 2, COLS_PRZES.KATEGORIA + 1).setValue(map[w]);
        changed++;
      }
    }
  }

  CacheService.getScriptCache().remove("katalog");
  return { success: true, changed: changed };
}

// ============================================
// LOCAL DATA BACKUP / SYNC
// ============================================

const SHEET_LOCAL_BACKUP = "LocalBackup";

function saveLocalBackup(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_LOCAL_BACKUP);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOCAL_BACKUP);
    sheet.appendRow(['Timestamp', 'Operator', 'Type', 'Data_JSON']);
    sheet.setFrozenRows(1);
    sheet.getRange('1:1').setFontWeight('bold');
  }

  var timestamp = new Date().toISOString();
  var operator = data.operator || 'unknown';
  var type = data.type || 'manual'; // manual | auto
  var jsonStr = JSON.stringify(data.payload);

  sheet.appendRow([timestamp, operator, type, jsonStr]);

  // Keep max 50 backups — remove oldest
  var lastRow = sheet.getLastRow();
  if (lastRow > 51) {
    sheet.deleteRows(2, lastRow - 51);
  }

  return { success: true, timestamp: timestamp };
}

function getLocalBackup() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_LOCAL_BACKUP);
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: false, message: 'Brak backupów' };
  }

  var lastRow = sheet.getLastRow();
  var row = sheet.getRange(lastRow, 1, 1, 4).getValues()[0];

  return {
    success: true,
    timestamp: row[0],
    operator: row[1],
    type: row[2],
    payload: JSON.parse(row[3])
  };
}

function getLocalBackupList() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_LOCAL_BACKUP);
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: true, backups: [] };
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var list = data.map(function(r, i) {
    return { index: i + 2, timestamp: r[0], operator: r[1], type: r[2] };
  }).reverse();

  return { success: true, backups: list.slice(0, 20) };
}

function getLocalBackupByRow(rowIndex) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_LOCAL_BACKUP);
  if (!sheet) return { success: false };

  var row = sheet.getRange(rowIndex, 1, 1, 4).getValues()[0];
  return {
    success: true,
    timestamp: row[0],
    operator: row[1],
    type: row[2],
    payload: JSON.parse(row[3])
  };
}
