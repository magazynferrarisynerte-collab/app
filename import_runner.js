// Importuje dane z Excela do arkusza Katalog
// Uruchom tę funkcję raz z edytora Apps Script
function importNarzedzia() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_KATALOG);
  
  var lastRow = sheet.getLastRow();
  Logger.log('Katalog: aktualnie ' + lastRow + ' wierszy (z nagłówkiem)');
  
  var allData = [].concat(getImportData1(), getImportData2(), getImportData3());
  Logger.log('Do importu: ' + allData.length + ' wierszy');
  
  // Dopisz poniżej istniejących danych
  var startRow = lastRow + 1;
  sheet.getRange(startRow, 1, allData.length, 9).setValues(allData);
  
  Logger.log('Import zakończony. Dodano ' + allData.length + ' wierszy od wiersza ' + startRow);
  SpreadsheetApp.getUi().alert('Import zakończony! Dodano ' + allData.length + ' pozycji.');
}
// ============================================
// IMPORT OSÓB z pliku 20.02.2026_Osoby z numerami.xlsx
// Kolumna B → Imię (kolumna B w Osoby)
// Kolumna D → Telefon (kolumna C w Osoby)
// ============================================

function importOsobyZPliku() {
  var lista = getImportOsoby();
  var result = importOsoby(lista);
  
  Logger.log('Import osób zakończony:');
  Logger.log('  Dodano: ' + result.added);
  Logger.log('  Pominięto (duplikaty): ' + result.skipped);
  SpreadsheetApp.getUi().alert('Import osób zakończony!\nDodano: ' + result.added + '\nPominięto: ' + result.skipped);
}