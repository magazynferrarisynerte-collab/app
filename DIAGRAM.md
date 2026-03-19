# Magazyn / Wypożyczalnia — Diagram Architektury

## 1. Architektura ogólna

```
┌─────────────────────────────────────────────────────────────────────┐
│                        PRZEGLĄDARKA (telefon/PC)                    │
│                                                                     │
│  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐ │
│  │   LOG    │ │  OSOBY   │ │NARZĘDZIA │ │ RAPORTY  │ │INWENTARY-│ │
│  │          │ │          │ │          │ │          │ │  ZACJA   │ │
│  └────┬─────┘ └────┬─────┘ └────┬─────┘ └────┬─────┘ └────┬─────┘ │
│       └──────┬─────┴──────┬─────┴──────┬─────┘            │       │
│              ▼            ▼            ▼                   ▼       │
│  ┌─────────────────────────────────────────┐  ┌──────────────────┐ │
│  │          MODAL WYDANIE / ZWROT          │  │  KAMERA + AI     │ │
│  │  (wybór osoby → narzędzia → zdjęcia)   │  │  (Gemini Vision) │ │
│  └───────────────────┬─────────────────────┘  └────────┬─────────┘ │
│                      │                                  │          │
│  ┌───────────────────▼──────────────────────────────────▼────────┐ │
│  │              RELIABLE OPERATION QUEUE                         │ │
│  │  ┌──────────┐  ┌──────────────────┐  ┌────────────────────┐  │ │
│  │  │IndexedDB │  │ fetch+keepalive  │  │google.script.run   │  │ │
│  │  │ (backup) │  │  (przeżywa lock) │  │(odpowiedź + retry) │  │ │
│  │  └──────────┘  └──────────────────┘  └────────────────────┘  │ │
│  │  batchId dedup ─── 3 warstwy dostarczania ─── auto-retry    │ │
│  └──────────────────────────┬────────────────────────────────────┘ │
└─────────────────────────────┼───────────────────────────────────────┘
                              │ HTTPS (POST / google.script.run)
                              ▼
┌─────────────────────────────────────────────────────────────────────┐
│                    GOOGLE APPS SCRIPT (V8)                          │
│                                                                     │
│  ┌─────────┐  ┌───────────┐  ┌──────────┐  ┌─────────┐            │
│  │ core.js │  │operacje.js│  │osoby.js  │  │katalog.js│           │
│  │ doGet() │  │wydajBatch │  │getOsoby  │  │getKatalog│           │
│  │ doPost()│  │zwrocBatch │  │addOsoba  │  │autoTag   │           │
│  │ dedup   │  │inwentaryz │  │rename    │  │groupKat  │           │
│  └────┬────┘  └─────┬─────┘  └────┬─────┘  └────┬─────┘           │
│       │              │             │              │                 │
│  ┌────┴──────┐  ┌────┴────┐  ┌────┴────┐  ┌─────┴─────┐           │
│  │raporty.js │  │admin.js │  │zamow.js │  │inwent.js  │           │
│  │getLog     │  │backup   │  │matrix   │  │AI match   │           │
│  │przenies*  │  │CRUD     │  │email    │  │scan+save  │           │
│  │scoring    │  │format   │  │trigger  │  │Gemini API │           │
│  └───────────┘  └─────────┘  └─────────┘  └───────────┘           │
│                              │                                      │
└──────────────────────────────┼──────────────────────────────────────┘
                               │
            ┌──────────────────┼──────────────────┐
            ▼                  ▼                  ▼
  ┌──────────────┐  ┌──────────────┐  ┌──────────────────┐
  │Google Sheets │  │ Google Drive │  │  Google Mail     │
  │              │  │              │  │                  │
  │• Katalog     │  │Magazyn_      │  │MailApp.send()   │
  │• Przesunięcia│  │  Zdjecia/    │  │• raporty tyg.   │
  │• Osoby       │  │  ├─Wydania/  │  │• zamówienia     │
  │• Inwentaryz. │  │  └─Zwroty/   │  │• CC: magazyn@   │
  │• Inv_*       │  │              │  │                  │
  │• LocalBackup │  │Backup copies │  │Trigger: Pt 12:00│
  └──────────────┘  └──────────────┘  └──────────────────┘
```

## 2. Flow: Wydanie narzędzi

```
  Użytkownik                    Frontend                         Backend (GAS)
      │                            │                                  │
      │  klik "+ Wydaj"            │                                  │
      ├───────────────────────────►│                                  │
      │                            │  openWypozyczModal()             │
      │  wybór osoby               │                                  │
      ├───────────────────────────►│  wypSelectOsoba()                │
      │                            │                                  │
      │  wybór narzędzi + ilości   │                                  │
      ├───────────────────────────►│  wypAddToCart()                  │
      │                            │  (opcja: zdjęcie, DMG)          │
      │                            │                                  │
      │  klik "Wydaj"             │                                  │
      ├───────────────────────────►│                                  │
      │                            │  ① _oqSave(batchId, data)       │
      │                            │     → IndexedDB                  │
      │  ◄── toast "Wysyłam..."   │                                  │
      │  ◄── modal zamknięty      │                                  │
      │                            │  ② _fetchKeepalive(POST)────────►│ doPost()
      │                            │     (przeżywa lock telefonu)     │  ├─ _isDuplicateBatch?
      │                            │                                  │  │  NIE → wydajBatch()
      │                            │  ③ google.script.run ───────────►│  │   ├─ decrementStock()
      │                            │     (z odpowiedzią)              │  │   ├─ appendRow(Przesunięcia)
      │                            │                                  │  │   ├─ savePhotoToDrive()
      │                            │  ◄─── {success, batchId} ───────┤  │   └─ return {success}
      │                            │  _oqRemove(batchId)              │  │
      │  ◄── toast "Wydano!"      │                                  │  └─ TAK → {deduplicated}
      │                            │  loadLogBackground()             │
      │                            │  loadKatalogBackground()         │
```

## 3. Flow: Zwrot narzędzi

```
  Użytkownik                    Frontend                         Backend (GAS)
      │                            │                                  │
      │  klik na osobę             │                                  │
      ├───────────────────────────►│  openOsobaDetail()               │
      │                            │  (zapamiętaj tab + scroll)       │
      │                            │                                  │
      │  zaznacz pozycje ☑         │                                  │
      ├───────────────────────────►│  toggleReturnCart()               │
      │                            │  (opcja: zmień ilość zwrotu)     │
      │                            │                                  │
      │  klik "Zwróć"             │                                  │
      ├───────────────────────────►│                                  │
      │                            │  (opcja: zdjęcia zwrotu)        │
      │                            │                                  │
      │                            │  ① _oqSave(batchId)             │
      │  ◄── modal zamknięty      │  ② _fetchKeepalive(POST) ──────►│ zwrocBatch()
      │                            │  ③ google.script.run ──────────►│  ├─ status='Zwrocone'
      │                            │                                  │  ├─ incrementStock()
      │                            │  ◄─── {success, count} ────────┤  └─ (partial: split row)
      │  ◄── toast "Zwrócono!"    │  _oqRemove(batchId)              │
```

## 4. Flow: Przekazanie narzędzi

```
  Osoba A ──[narzędzia]──► Osoba B

  Użytkownik                    Frontend                         Backend
      │                            │                                  │
      │  otwórz osobę A           │                                  │
      │  zaznacz pozycje           │                                  │
      │  klik "Przekaż"          │                                  │
      ├───────────────────────────►│  transferSelected()              │
      │                            │  → lista osób + search           │
      │  wybierz osobę B          │                                  │
      ├───────────────────────────►│  executeTransfer()               │
      │                            │  ① _oqSave ② fetch ③ g.s.run──►│ przeniesNaOsobe()
      │                            │  (lokalny update natychmiast)    │  └─ zmień OSOBA w arkuszu
      │  ◄── toast "Przekazano"   │                                  │
```

## 5. Reliable Queue — 3 warstwy

```
┌──────────────────────────────────────────────────────────────────┐
│                    OPERACJA (wydaj/zwrot/przekaż)                │
│                                                                  │
│  batchId = "b_1711036800000_x7k2m9"                             │
│                                                                  │
│  WARSTWA 1: IndexedDB                                            │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │ Zapisz przed wysłaniem. Przeżywa: crash, restart, reload  │  │
│  │ Retry: _oqRetryPending() przy kolejnym załadowaniu strony │  │
│  │ TTL: 24h (potem auto-discard)                              │  │
│  └────────────────────────────────────────────────────────────┘  │
│                                                                  │
│  WARSTWA 2: fetch + keepalive                                    │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │ POST na APP_URL, mode: 'no-cors'                          │  │
│  │ Przeżywa: zamknięcie zakładki, blokada telefonu           │  │
│  │ Limit: body < 64KB                                         │  │
│  └────────────────────────────────────────────────────────────┘  │
│                                                                  │
│  WARSTWA 3: google.script.run                                    │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │ Standardowe wywołanie GAS z callbackami                    │  │
│  │ onSuccess → _oqRemove(batchId) → toast                    │  │
│  │ onFailure → zostaw w queue (retry przy reload)             │  │
│  └────────────────────────────────────────────────────────────┘  │
│                                                                  │
│  SERWER: CacheService dedup (batchId → 1h TTL)                  │
│  → Nawet jeśli 2-3 warstwy dotrą = 1 wykonanie                 │
└──────────────────────────────────────────────────────────────────┘
```

## 6. Struktura plików

```
app/
├── core.js              ← doGet, doPost, stałe, routing (30+ actions)
├── operacje.js          ← wydajBatch, zwrocBatch, inwentaryzujBatch, zdjęcia
├── osoby.js             ← CRUD osoby, rename, import
├── katalog.js           ← CRUD katalog, autoTag, grupowanie, SN
├── raporty.js           ← log, summary, uszkodzone, przeniesienia, scoring search
├── admin.js             ← backup, CRUD arkuszy, email, formatowanie
├── zamowienia.js        ← macierz pozycji×lokalizacje, raport email, trigger
├── inwentaryzacja.js    ← skan AI (Gemini Vision), dopasowanie, sesje
│
├── index.html           ← główny template, 5 tabów, modale, nawigacja
├── styles.html          ← CSS (~870 linii)
├── js_globals.html      ← zmienne, helpers, queue, modal tracking
├── js_osoby.html        ← renderLog, renderOsoby, vCard
├── js_katalog.html      ← renderNarzedzia, search listeners, grupowanie
├── js_operacje.html     ← modal wydania, koszyk, zdjęcia
├── js_narzedzia.html    ← modal detalu narzędzia, historia
├── js_raporty.html      ← raporty, scoring search, synonimy, macierz UI
├── js_inwentaryzacja.html ← kamera, AI match, sesje
├── js_admin.html        ← data manager UI
├── js_init.html         ← preload, version check, auto-sync
├── js_zamowienia.html   ← email zamówień, macierz edycja
├── preview_lista.html   ← podgląd wydruku listy
└── preview_zam.html     ← podgląd wydruku zamówienia
```

## 7. Arkusze Google (model danych)

```
┌─────────────────────────────────────────────────────────────────┐
│ KATALOG (13 kolumn)                                             │
│ ID │ NazwaSys │ NazwaWys │ Kat │ SN │ StanPocz │ AktStan │    │
│ Flaga │ Tagi │ OstatnioWidz │ Opis │ Przesun │ DataPrzesun │   │
│                                                                 │
│ Kategorie: E=ewidencjonowane(SN) │ N=normalne │ Z=zużywalne    │
├─────────────────────────────────────────────────────────────────┤
│ PRZESUNIĘCIA (13 kolumn) — log wszystkich operacji              │
│ ID_Op │ DataWyd │ Osoba │ NazwaSys │ SN │ Ilość │ Kat │       │
│ Status │ ZdjWyd │ ZdjZwr │ DataZwr │ OpisUszkodz │ Operator │  │
│                                                                 │
│ Statusy: Wydane │ Zwrocone │ Uszkodzone │ Zużyte │ Inwentaryz. │
├─────────────────────────────────────────────────────────────────┤
│ OSOBY (5 kolumn)                                                │
│ ID │ Imię │ Telefon │ Lokalizacja │ Email                      │
│                                                                 │
│ Pseudo-osoby: [Lista] │ [Robocza] │ [Zamówienia]               │
│ → nie zmieniają stanu magazynowego (oprócz [Lista])             │
└─────────────────────────────────────────────────────────────────┘
```

## 8. Wyszukiwarka (scored search)

```
  Query: "wiert"                    Algorytm scoreWord():
      │                             ┌────────────────────────────┐
      ▼                             │ 1.0  exact match           │
  normalize()                       │ 0.9  starts with query     │
  "wiert" (bez ąęś...)             │ 0.88 word prefix match     │
      │                             │ 0.85 contains query        │
      ▼                             │ 0.7  Levenshtein ≤ maxLev  │
  expandWithSynonyms()              │ 0.6  prefix-Levenshtein    │
  ["wiert", "wiertarko-            │ 0.0  no match              │
   wkretarka", "wkretarka"]         └────────────────────────────┘
      │
      ▼                             maxLev:
  scoreItem() per field             ├─ słowo ≤2 zn: 0 (brak tolerancji)
  [nazwaSys, nazwaWys,              ├─ słowo 3-4 zn: 1 literówka
   kategoria, sn, tagi]             └─ słowo 5+ zn:  2 literówki
      │
      ▼                             Synonimy (16 grup):
  sort by score DESC                szlifierka ↔ flex ↔ grinder
  return top matches                taśma ↔ miara ↔ metr ↔ miarka
                                    pilarka ↔ piła ↔ pilarki
  Debounce: 150ms                   kombinerki ↔ szczypce ↔ obcęgi
```

## 9. Integracje zewnętrzne

```
┌──────────┐     ┌──────────────┐     ┌──────────────┐
│ WhatsApp │     │ Google Mail  │     │ Gemini Vision│
│ (wa.me)  │     │ (MailApp)    │     │    (AI)      │
│          │     │              │     │              │
│ Stan     │     │ Raport tyg.  │     │ Rozpoznaj   │
│ narzędzi │     │ (Pt 12:00)  │     │ narzędzia   │
│ osoby    │     │              │     │ ze zdjęcia  │
│          │     │ Zamówienia   │     │              │
│ Lista    │     │ akceptacja   │     │ → dopasuj   │
│ zwrotów  │     │              │     │   do katalog│
└──────────┘     │ CC: magazyn@ │     └──────────────┘
                 └──────────────┘

┌──────────────┐     ┌──────────────┐
│ Google Drive │     │ localStorage │
│              │     │ + IndexedDB  │
│ Zdjęcia:    │     │              │
│ /Wydania/   │     │ operator     │
│ /Zwroty/    │     │ invOsoba     │
│              │     │ invHidden    │
│ Backup      │     │ OpQueue      │
│ arkusza     │     │ (reliable)   │
└──────────────┘     └──────────────┘
```
