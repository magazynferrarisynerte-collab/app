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

  var data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

  var katalog = data.map(function (r) {
    var sn = extractSN(r[COLS_KATALOG.SN]);
    var kategoria = String(r[COLS_KATALOG.KATEGORIA] || '').trim().toUpperCase();
    if (sn) kategoria = 'E';
    else if (kategoria !== 'E' && kategoria !== 'Z') kategoria = 'N';
    var flagVal = r[COLS_KATALOG.FLAGA];
    var owVal = r[COLS_KATALOG.OSTATNIO_WIDZIANE];
    var ostatnioWidziane = '';
    if (owVal instanceof Date) {
      ostatnioWidziane = Utilities.formatDate(owVal, 'Europe/Warsaw', 'dd.MM.yyyy HH:mm');
    } else if (owVal) {
      try { ostatnioWidziane = Utilities.formatDate(new Date(owVal), 'Europe/Warsaw', 'dd.MM.yyyy HH:mm'); } catch(e) {}
    }
    return {
      id: String(r[COLS_KATALOG.ID]),
      nazwaSys: String(r[COLS_KATALOG.NAZWA_SYSTEMOWA] || ''),
      nazwaWys: String(r[COLS_KATALOG.NAZWA_WYSWIETLANA] || ''),
      kategoria: kategoria,
      sn: sn,
      stanPoczatkowy: Number(r[COLS_KATALOG.STAN_POCZATKOWY]) || 0,
      aktualnieNaStanie: Number(r[COLS_KATALOG.AKTUALNIE_NA_STANIE]) || 0,
      flaga: flagVal === true || flagVal === 1 || String(flagVal) === '1',
      tagi: String(r[COLS_KATALOG.TAGI] || ''),
      ostatnioWidziane: ostatnioWidziane,
      opis: String(r[COLS_KATALOG.OPIS] || ''),
      przesun: String(r[COLS_KATALOG.PRZESUN] || ''),
      dataPrzesun: r[COLS_KATALOG.DATA_PRZESUN] instanceof Date
        ? Utilities.formatDate(r[COLS_KATALOG.DATA_PRZESUN], 'Europe/Warsaw', 'dd.MM.yyyy')
        : (r[COLS_KATALOG.DATA_PRZESUN] ? String(r[COLS_KATALOG.DATA_PRZESUN]) : '')
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
        ostatnioWidziane: '',
        items: []
      };
    }
    if (k.flaga) groups[key].flaga = true;
    groups[key].totalStock += k.aktualnieNaStanie;
    groups[key].totalPoczatkowy += k.stanPoczatkowy;
    if (k.ostatnioWidziane && (!groups[key].ostatnioWidziane || k.ostatnioWidziane > groups[key].ostatnioWidziane)) {
      groups[key].ostatnioWidziane = k.ostatnioWidziane;
    }
    groups[key].items.push({
      id: k.id,
      nazwaSys: k.nazwaSys,
      sn: k.sn,
      aktualnieNaStanie: k.aktualnieNaStanie,
      flaga: k.flaga,
      ostatnioWidziane: k.ostatnioWidziane || ''
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

function addKatalogItem(nazwaSys, nazwaWys, kategoria, sn, stanPoczatkowy, opis, przesun) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getSheet(SHEET_KATALOG);
    var id = generateId('KT');
    var qty = Number(stanPoczatkowy) || 1;
    sheet.appendRow([id, nazwaSys.trim(), nazwaWys.trim(), kategoria, sn || '', qty, qty, 0, '', '', opis || '', przesun || '', '']);
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
