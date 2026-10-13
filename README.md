# Avtomatizacija obdelave bankovnih izpisov in Google Sheets

Ta projekt služi za avtomatizirano izvlekanje transakcij iz PDF izpisov kartičnih računov in njihovo urejeno shranjevanje v Excel datoteke ali neposredno v Google Sheets.

## Struktura map in datotek

- **`izpiski/`**: V tej mapi se nahajajo vsi originalni PDF izpisi, ki jih je treba obdelati.
- **`process.py`**: Glavni Python skript za lokalno obdelavo in shranjevanje v Excel.
- **`process_sheets.py`**: Skript za avtomatizirano dodajanje transakcij neposredno v Google Sheets.
- **`merge_to_excel.py`**: Pomožni skript za dodajavanje transakcij iz JSON posrednika v Excel.
- **`Transakcije (5).xlsx`**: Glavna Excel datoteka za lokalno shranjevanje.
- **`processed.txt`** / **`processed_sheets.txt`**: Datoteki, ki shrani imena že obdelanih PDF datotek.
- **`credentials.json`**: Google API ključi za dostop do Google Sheets.

## Kako uporabiti

### 1. Lokalna obdelava (Excel)

1.  **Priprava**: V mapi `izpiski/` postavite nove PDF izpise.
2.  **Izvedba**: V terminalu zaženite:
    ```bash
    python process.py
    ```
3.  **Rezultat**: Transakcije bodo samodejno dodane v `Transakcije (5).xlsx` na tab `transakcije`, urejene po datumu.

### 2. Obdelava v Google Sheets

1.  **Priprava**: 
    - Postavite `credentials.json` v glavno mapo.
    - Dodajte email service accounta v Share Google tabele.
    - V mapi `izpiski/` postavite nove PDF izpise.
2.  **Izvedba**: V terminalu zaženite:
    ```bash
    python process_sheets.py
    ```
3.  **Rezultat**: Transakcije bodo dodane neposredno v Google tabelo.

## Tehnologije

- **Python**: Glavna programska logika.
- **pypdf**: Za branje in izvlekanje besedila iz PDF datotek.
- **openpyxl**: Za branje in pisanje v Excel datoteke (.xlsx).
- **gspread**: Za interakcijo z Google Sheets API.
- **oauth2client**: Za avtentikacijo z Google API.