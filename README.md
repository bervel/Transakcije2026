# Avtomatizacija obdelave bankovnih izpisov

Ta projekt služi za avtomatizirano izvlekanje transakcij iz PDF izpisov kartičnih računov in njihovo urejeno shranjevanje v Excel datoteko.

## Struktura map in datotek

- **`izpiski/`**: V tej mapi se nahajajo vsi originalni PDF izpisi, ki jih je treba obdelati.
- **`process.py`**: Glavni Python skript, ki opravlja naslednja opravila:
    - Branje in parsanje PDF datotek.
    - Kategorizacijo transakcij (npr. Avto, Dom, Šport itd.).
    - Deduplikacijo (preprečuje ponovljen vnos istih transakcij).
    - Čiščenje podatkov (odstranjevanje povzetekov in skupnih zneskov, npr. VISA ali Mastercard povzemki).
    - Kronološko vstavljanje transakcij v obstoječo Excel datoteko.
- **`merge_to_excel.py`**: Pomožni skript za dodajanje transakcij iz JSON posrednika v Excel.
- **`Transakcije (5).xlsx`**: Glavna Excel datoteka, kjer so shranjene vse očiščene in urejene transakcije.
- **`processed.txt`**: Datoteka, ki shrani imena že obdelanih PDF datotek, da se ne obdelajo ponovno.

## Kako uporabiti

1.  **Priprava**: V mapi `izpiski/` postavite nove PDF izpise.
2.  **Izvedba**: V terminalu zaženite:
    ```bash
    python process.py
    ```
3.  **Rezultat**: Transakcije bodo samodejno dodane v `Transakcije (5).xlsx` na tab `transakcije`, urejene po datumu.

## Tehnologije

- **Python**: Glavna programska logika.
- **pypdf**: Za branje in izvlekanje besedila iz PDF datotek.
- **openpyxl**: Za branje in pisanje v Excel datoteke (.xlsx).
