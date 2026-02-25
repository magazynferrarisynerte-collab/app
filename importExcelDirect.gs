// ============================================
// IMPORT OSÓB BEZPOŚREDNIO Z PLIKU EXCEL NA DRIVE
// Wklej tę funkcję do Kod.js w Google Apps Script
// ============================================

/**
 * Importuje osoby z pliku Excel na Google Drive bezpośrednio do arkusza Osoby
 * Szuka pliku o nazwie zawierającej "Osoby z numerami"
 * Kolumna B → Imie (kolumna B w Osoby)
 * Kolumna D → Telefon (kolumna C w Osoby)
 */
function importOsobyZExcelaGoogleDrive() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var osbySheet = ss.getSheetByName(SHEET_OSOBY);
    
    // Szukaj pliku Excel z "Osoby z numerami"
    var files = DriveApp.getFilesByName("20.02.2026_Osoby z numerami.xlsx");
    
    if (!files.hasNext()) {
      throw new Error('Plik "20.02.2026_Osoby z numerami.xlsx" nie znaleziony na Drive');
    }
    
    var file = files.next();
    var blob = file.getBlob();
    
    // Konwertuj do Sheets (tymczasowo)
    var resource = {
      title: "TEMP_IMPORT_" + Date.now(),
      mimeType: "application/vnd.google-apps.spreadsheet"
    };
    
    var tempSheet = Drive.Files.insert(resource, blob, { convert: true });
    var tempSS = SpreadsheetApp.openById(tempSheet.id);
    var tempWs = tempSS.getSheets()[0];
    
    // Wczytaj dane: kolumna B i D
    var lastRow = tempWs.getLastRow();
    var data = [];
    
    for (var i = 2; i <= lastRow; i++) {
      var imie = String(tempWs.getRange(i, 2).getValue()).trim();
      var telefon = String(tempWs.getRange(i, 4).getValue()).trim();
      
      if (imie && imie.length > 0 && imie !== "null") {
        data.push({ imie: imie, telefon: telefon === "null" ? "" : telefon });
      }
    }
    
    // Usuń tymczasowy arkusz
    Drive.Files.remove(tempSheet.id);
    
    // Importuj dane do Osoby
    var result = importOsoby(data);
    
    Logger.log("Import osób z Excela zakończony:");
    Logger.log("  - Dodano: " + result.added);
    Logger.log("  - Pominięto (duplikaty): " + result.skipped);
    Logger.log("  - Razem w bazie: " + (osbySheet.getLastRow() - 1));
    
    SpreadsheetApp.getUi().alert(
      'Import z Excela zakończony!\n\n' +
      'Dodano: ' + result.added + '\n' +
      'Pominięto: ' + result.skipped + '\n' +
      'Razem osób: ' + (osbySheet.getLastRow() - 1)
    );
    
    return { success: true, added: result.added, skipped: result.skipped };
  } catch (e) {
    Logger.log("Błąd: " + e.toString());
    SpreadsheetApp.getUi().alert("Błąd importu: " + e.toString());
    return { success: false, error: e.toString() };
  }
}
