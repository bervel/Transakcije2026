# -*- coding: utf-8 -*-
"""
process_sheets.py — Samodejna obdelava PDF-jev in vnos v Google Sheets
====================================================================
Uporaba:
  1. Postavi credentials.json (iz Google Cloud)
  2. Dodaj email service accounta v Share Google tabele
  3. Kopiraj PDF v mapo 'izpiski'
  4. Pritisni:  python process_sheets.py
"""
import re, os, sys, unicodedata, shutil, traceback
from datetime import date, timedelta
from collections import Counter
from pypdf import PdfReader
import gspread
import openpyxl
from oauth2client.service_account import ServiceAccountCredentials
import warnings; warnings.filterwarnings("ignore")

DIR          = os.path.dirname(os.path.abspath(__file__))
os.chdir(DIR)
sys.stdout.reconfigure(encoding="utf-8")

# ── Konfiguracija ──────────────────────────────────────────────────────────
PDF_PREFIX     = "Izpisek_"
SUBDIR         = "izpiski"
STATE_FILE     = "processed_sheets.txt"
SHEET_NAME     = "transakcije"      # Ime taba v Google tabeli
GOOGLE_SHEET_ID = "1JIAfmNICTokIzNeQ5OV17nnVhopUkESG9GnSgywGUGQ"
CREDENTIALS_FILE = "credentials.json"

# Konfiguracija za Excel (osnovni račun)
EXCEL_PATTERN  = r"Transakcije \(\d+\)\.xlsx"
# ────────────────────────────────────────────────────────────────────────────

# ── Pomožne funkcije (Parsing & Kategorizacija) ────────────────────────────

def load_processed():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding="utf-8") as f:
            return set(line.strip() for line in f if line.strip())
    return set()

def save_processed(s):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        for p in sorted(s):
            f.write(p + "\n")

def fd(d):  return f"{d.day}. {d.month}. {d.year}"

def pd(v):
    if not v: return None
    p = [int(x) for x in re.findall(r"\d+", str(v))]
    return date(p[2], p[1], p[0]) if len(p) >= 3 else None

def norm(s):
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").lower()
    return re.sub(r"[^a-z0-9]", "", s)

# ── PDF parsanje ───────────────────────────────────────────────────────────

RX = re.compile(r"^(\d{2}\.\d{2}\.\d{4})\s+(.+?)\s+(\d{1,3}(?:\.\d{3})*,\d{2})(?:[+\-])?\s*$")
SKIP = ["BERNARD","VISAPREMIUM","VISA_PREMIUM","474819","OŽIJEK","OZIJEK","PORABA",
        "PLAČA","PLACAN","ZARAČUN","ZARACUN","ZNESEK","KARTIC","SKUPAJ","POVZETEK",
        "PLAČATEL","PLAČILO","NARAVNA","STORITVE","VRACANJE","KREDITNO","DEBITNO",
        "OBLIČJE","VRSTA","KARTICA","MASTERCARD","SKUPAJ ZNESEK PORABE"]
FX = {"DKK","SEK","CHF","USD","GBP","PLN","CZK","HUF","CNY","JPY","EUR"}

def parse_pdf(path):
    out = []
    for pg in PdfReader(path).pages:
        for ln in pg.extract_text().splitlines():
            ln = re.sub(" +", " ", ln.strip())
            if not ln or any(k in ln for k in SKIP): continue
            m = RX.match(ln)
            if not m: continue
            dt, desc, val = m.groups()
            d_s, mo_s, y_s = dt.split(".")
            out.append({
                "d":    date(int(y_s), int(mo_s), int(d_s)),
                "desc": re.sub(r"\s+", " ", re.sub(r"\bEUR\b", "", desc)).strip().rstrip(","),
                "v":    float(val.replace(".", "").replace(",", ".")),
                "cur":  next((t for t in desc.replace(",", " ").upper().split() if t in FX), "EUR"),
                "neg":  ln.rstrip().endswith("-"),
            })
    return out

def parse_excel(path):
    """Parsa Excel datoteko s transakcijami računa."""
    out = []
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet = wb.active
        # Preberi vse podatke iz aktivnega lista
        for row in sheet.iter_rows(min_row=2, values_only=True):
            # Struktura: A=Datum, B=Opis, C=Znesek (glede na prejšnje razgovore)
            # row[0]=Datum, row[1]=Opis, row[2]=Znesek
            if not row or row[0] is None: continue
            
            try:
                # Datumska obdelava (Excel lahko vrne datetime objekt)
                # Excel stolpci: A=Namen, B=Kategorija, C=+/- , D=Znesek, E=Valuta, F=Datum
                dt_val = row[5]
                if isinstance(dt_val, date):
                    d = dt_val
                else:
                    # Če je string, poskušaj parsati (npr. 23. 3. 2026)
                    p = [int(x) for x in re.findall(r"\d+", str(dt_val))]
                    d = date(p[2], p[1], p[0]) if len(p) >= 3 else None
                
                if not d:
                    print(f"    [DEBUG] Neznan datum: {dt_val}")
                    continue

                desc = str(row[0]).strip() if row[0] else ""
                v_val = row[3]
                if v_val is None:
                    print(f"    [DEBUG] Prazn znesek v vrstici: {row}")
                    continue
                v = float(str(v_val).replace(",", "."))
                
                out.append({
                    "d":    d,
                    "desc": desc,
                    "v":    v,
                    "cur":  "EUR",  # Excel običajno nima valute v stolpcu, predpostavljamo EUR
                    "neg":  v < 0,  # Če je znesek negativ way, je to izdatki
                })
            except Exception as e:
                print(f"    [DEBUG] Napaka v vrstici {row}: {e}")
                continue
    except Exception as e:
        print(f"  ☒  Napaka pri branju Excelja {path}: {e}")
    return out

# ── Kategorije ─────────────────────────────────────────────────────────────

CAT = [
    ("Šport & Prosti čas",    ["extremevital","steam","eventim","sportvision","sport"]),
    ("Avto & Prevozi",        ["petrol","inamarket","bsmol","tamoil","zminj","autopark",
                               "petrolio","span","drivr","uber","wallismotor","sixt",
                               "sixinz","europeanpark","petrolm","euroam","motocenter"]),
    ("Dom",                   ["suncontract","rtvslovenija","teleslovenija","telemach",
                               "komunalne","slodkovec","sodis"]),
    ("Finance & Zavarovanja", ["zavarovalnica","zavarov"]),
    ("Kavarne & Restavracije",["kavarna","brunch","restavracija","restaurant","boreka",
                               "laturka","burek","lahocolate","chutis","hoodburger",
                               "domacapeka","nyhavn38","ubereats","wolt","foetex",
                               "belacenter","pressbyran","skastnmalmo","twoforks",
                               "heinemann","tubik","gastro"]),
]

def categorize(desc):
    d = norm(desc)
    for cat, kws in CAT:
        if any(k in d for k in kws):
            return cat
    return "Nakupi"

# ── Google Sheets Logika ───────────────────────────────────────────────────

def get_gsheet_client():
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
    if not os.path.exists(CREDENTIALS_FILE):
        print(f"☒  Ni datoteke {CREDENTIALS_FILE}! Postavi jo iz Google Cloud konzole.")
        sys.exit(1)
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    return gspread.authorize(creds)

def build_existing_gsheet(ws):
    all_values = ws.get_all_values()
    keyed = []
    amt_counter = Counter()
    for row in all_values[1:]:
        if len(row) < 6: continue
        d_val = pd(row[5])
        v_val = row[3]
        if d_val and v_val:
            try:
                v_float = float(v_val.replace(" ", "").replace(",", "."))
                desc_norm = norm(row[0][:60])
                keyed.append((d_val.month, round(v_float, 2), desc_norm))
                amt_counter[(d_val.month, round(v_float, 2))] += 1
            except ValueError:
                continue
    return keyed, amt_counter

def already_exists_gsheet(t, keyed, amt_counter):
    m, a = t["d"].month, round(t["v"], 2)
    c = amt_counter.get((m, a), 0)
    if c == 0: return False
    if c >= 3: return True
    short = norm(t["desc"])[:10]
    if len(short) >= 3:
        return any(mm == m and aa == a and short in nm for mm, aa, nm in keyed)
    return False

def insert_all_gsheet(ws, items):
    n = 0
    # 1 · Preberi vse obstoječe podatke enkrat
    all_rows = ws.get_all_values()
    
    if not all_rows:
        if not items:
            return 0
        final_data = [make_row_list(t) for t in sorted(items, key=lambda x: x["d"])]
        ws.update('A1', final_data)
        return len(items)

    header = all_rows[0]
    data = all_rows[1:]

    # 2 · Pripravi obstoječe vrstice z datumi za sortiranje
    rows_with_dates = []
    for row in data:
        d_val = pd(row[5]) if len(row) > 5 else None
        if not d_val:
            d_val = date(1, 1, 1)
        rows_with_dates.append((d_val, row))

    # 3 · Pripravi nove vrstice z datumi
    new_rows_with_dates = []
    for t in items:
        new_rows_with_dates.append((t["d"], make_row_list(t)))

    # 4 · Združi in razvrsti vse po datumu
    combined = rows_with_dates + new_rows_with_dates
    combined.sort(key=lambda x: x[0])

    # 5 · Izvzemi samo vrednosti vrstic
    final_data = [header] + [item[1] for item in combined]

    # 6 · Ponastavi tabelo in vnesi vse podatke z enim zahtevkom (Batch Update)
    ws.clear()
    ws.update('A1', final_data)
    
    return len(items)

def make_row_list(t):
    d, v, cur, desc = t["d"], t["v"], t["cur"], t["desc"]
    txt  = f"NAKUP {desc} {v:,.2f} {cur}".replace(",", " ")
    sign = "+" if t["neg"] else "-"
    # Vrstica vsebuje: [Opis, Kategorija, Znak, Znesek, Valuta, Datum, + praznih stolpcev]
    return [txt, t["cat"], sign, round(v, 2), cur, fd(d), "", "", "", "", "", "", "", "", "", ""]

# ── Glavni program ─────────────────────────────────────────────────────────

def main():
    print("\n" + "=" * 52)
    print(" KARTIČNI IZPISKI — GOOGLE SHEETS")
    print("=" * 52)

    client = get_gsheet_client()

    # 3 · Odpri Excel (Google Sheets)
    try:
        print(f"  ... poskušam odpreti tabelo preko ID-ja {GOOGLE_SHEET_ID} ...")
        spreadsheet = client.open_by_key(GOOGLE_SHEET_ID)
        print(f"  ✓  Tabela odprta.")
        
        # Izpiši vse razpoložljive tabe (worksheets)
        all_worksheets = [sh.title for sh in spreadsheet.worksheets()]
        print(f"  ℹ  Razpoložljivi tabi v tej tabeli: {all_worksheets}")
        
        try:
            ws = spreadsheet.worksheet(SHEET_NAME)
            print(f"  ✓  Tab '{SHEET_NAME}' najden.")
        except Exception as e:
            print(f"  ☒  Napaka pri iskanju taba '{SHEET_NAME}': {e}")
            print(f"     Preveri, če je ime taba v seznamu zgoraj pravilno.")
            return

    except Exception as e:
        print(f"☒  Napaka pri dostopu do Google tabele: {e}")
        traceback.print_exc()
        print(f"     Preveri, če je ID '{GOOGLE_SHEET_ID}' pravilno zapisan in ali si Service Accountu dal dostop (Share).")
        return

    all_pdfs = sorted(f for f in os.listdir(os.path.join(DIR, SUBDIR))
                      if f.lower().endswith(".pdf") and f.lower().startswith(PDF_PREFIX.lower()))
    processed = load_processed()
    new_pdfs  = [p for p in all_pdfs if p not in processed]

    if not new_pdfs:
        print("  ✓  Ni novih PDF-jev za obdelavo.")
        all_txns = []
    else:
        print(f"  PDF-jev skupno : {len(all_pdfs)}")
        print(f"  Novih (za obdelavo): {len(new_pdfs)}")
        all_txns = []
        # 4 · PDF-ji
        for p in new_pdfs:
            txns = parse_pdf(os.path.join(DIR, SUBDIR, p))
            print(f"  ✓  {p}: {len(txns)} transakcij")
            all_txns.extend(txns)

    # 4b · Excel (osnovni račun)
    all_excels = sorted(f for f in os.listdir(DIR)
                        if f.lower().endswith(".xlsx") and re.match(EXCEL_PATTERN, f))
    processed_excels = [e for e in all_excels if e not in processed]

    if processed_excels:
        print(f"  Excel datotek skupno: {len(all_excels)}")
        print(f"  Novih (za obdelavo): {len(processed_excels)}")
        for e in processed_excels:
            txns = parse_excel(os.path.join(DIR, e))
            print(f"  ✓  {e}: {len(txns)} transakcij")
            all_txns.extend(txns)

    if not all_txns:
        # Update processed for both PDFs and Excels if nothing new was found
        processed.update(new_pdfs)
        processed.update(processed_excels)
        save_processed(processed)
        print("\n  ⚠  Nobene nove transakcije ni bilo mogoče najti (PDF ali Excel).")
        return

    # 5 · Dedup
    keyed, amt_counter = build_existing_gsheet(ws)
    to_add, skipped = [], 0
    for t in all_txns:
        t["cat"] = categorize(t["desc"])
        if already_exists_gsheet(t, keyed, amt_counter):
            skipped += 1
        else:
            to_add.append(t)
    print(f"  ✓  Dedup: preskočenih {skipped} | za vnos {len(to_add)}")

    if not to_add:
        processed.update(new_pdfs)
        save_processed(processed)
        print("\n  Vse transakcije iz novih PDF-jev so že v Google Excel-u.")
        return

    print("  ... vstavljanje v Google Sheets ...")
    n_insert = insert_all_gsheet(ws, to_add)
    print(f"  ✓  Vstavljenih vrstic: {n_insert}")

    processed.update(new_pdfs)
    save_processed(processed)

    print("\n" + "=" * 52)
    print("  KONČANO!")
    print("=" * 52 + "\n")

if __name__ == "__main__":
    main()
