# -*- coding: utf-8 -*-
"""
process.py — Samodejna obdelava novih kartičnih PDF izpisov
=============================================================
Uporaba:
  1. Kopiraj novi PDF izpis (Izpisek_*.pdf) v to mapo
  2. Pritisni:  python process.py
  3. Končano — novi nakupi v Excel, VISA_PREMIUM odstranjen

Resete state (če si napačno označil PDF kot obdelan):
  → izbriši datoteko  processed.txt  v tej mapi
"""
import re, os, sys, unicodedata, shutil
from datetime import date, timedelta
from collections import Counter
from pypdf import PdfReader
import openpyxl
import warnings; warnings.filterwarnings("ignore")

DIR          = os.path.dirname(os.path.abspath(__file__))
os.chdir(DIR)
sys.stdout.reconfigure(encoding="utf-8")

# ── Konfiguracija ──────────────────────────────────────────────────────────
PDF_PREFIX   = "Izpisek_"       # samo PDF z tem imenom se obravnavajo
SUBDIR       = "izpiski"        # mapa, kjer so PDF izpisi
STATE_FILE   = "processed.txt"  # seznam že obdelanih PDF-jev (ime datoteke)
SHEET_NAME   = "transakcije"   # ime taba v Excel
# ────────────────────────────────────────────────────────────────────────────

# ── Pomožne funkcije ───────────────────────────────────────────────────────

def find_xlsx():
    files = [f for f in os.listdir(DIR)
             if f.lower().endswith(".xlsx") and "backup" not in f.lower()]
    if len(files) == 0:
        print("☒  V mapi ni nobene .xlsx datoteke (razen backup).")
        sys.exit(1)
    if len(files) > 1:
        print(f"☒  V mapi je  {len(files)}  xlsx (razen backup): {files}")
        print("   Pusti samo eno ali preimenuj odvečne.")
        sys.exit(1)
    return files[0]

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

def pd(v):  # "3. 2. 2026" → date
    p = [int(x) for x in re.findall(r"\d+", str(v))]
    return date(p[2], p[1], p[0]) if len(p) >= 3 else None

def norm(s):
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").lower()
    return re.sub(r"[^a-z0-9]", "", s)

# ── PDF parsanje ───────────────────────────────────────────────────────────

# Regex za iskanje vrstic: Datum (DD.MM.YYYY) + Opis + Znesek (npr. 1.234,56)
RX   = re.compile(r"^(\d{2}\.\d{2}\.\d{4})\s+(.+?)\s+(\d{1,3}(?:\.\d{3})*,\d{2})(?:[+\-])?\s*$")

# Ključne besede, ki jih preskočimo (povzemki, naslovi, itd.)
SKIP = ["BERNARD","VISAPREMIUM","VISA_PREMIUM","474819","OŽIJEK","OZIJEK","PORABA",
        "PLAČA","PLACAN","ZARAČUN","ZARACUN","ZNESEK","KARTIC","SKUPAJ","POVZETEK",
        "PLAČATEL","PLAČILO","NARAVNA","STORITVE","VRACANJE","KREDITNO","DEBITNO",
        "OBLIČJE","VRSTA","KARTICA","MASTERCARD","SKUPAJ ZNESEK PORABE"]

# Podprte valute
FX = {"DKK","SEK","CHF","USD","GBP","PLN","CZK","HUF","CNY","JPY","EUR"}

def parse_pdf(path):
    \"\"\"Prebere PDF datoteko in izvzleka transakcije v list dict-ov.\"\"\"
    out = []
    for pg in PdfReader(path).pages:
        for ln in pg.extract_text().splitlines():
            ln = re.sub(" +", " ", ln.strip())
            # Preskoči prazne vrstice ali vrstice s prepovedanimi besedami
            if not ln or any(k in ln for k in SKIP):
                continue
            m = RX.match(ln)
            if not m:
                continue
            dt, desc, val = m.groups()
            d_s, mo_s, y_s = dt.split(".")
            out.append({
                "d":    date(int(y_s), int(mo_s), int(d_s)),
                "desc": re.sub(r"\s+", " ", re.sub(r"\bEUR\b", "", desc)).strip().rstrip(","),
                "v":    float(val.replace(".", "").replace(",", ".")),
                "cur":  next((t for t in desc.replace(",", " ").upper().split() if t in FX), "EUR"),
                "neg":  ln.rstrip().endswith("-"), # Preveri, če je transakcija odobritev/plačilo (znamenje -)
            })
    return out

# ── Kategorije ─────────────────────────────────────────────────────────────

# Seznam kategorij in ključnih besed za njihovo prepoznavanje
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
    \"\"\"Deli transakcijo v kategorije na podlagi opisa.\"\"\"
    d = norm(desc)
    for cat, kws in CAT:
        if any(k in d for k in kws):
            return cat
    return "Nakupi"

# ── Deduplikacija ──────────────────────────────────────────────────────────

def build_existing(ws):
    \"\"\"Zbere obstoječe transakcije iz Excela za preverjanje duplikata.\"\"\"
    keyed = []
    for r in ws.iter_rows(min_row=2):
        v0 = r[0].value
        if v0 is None and r[3].value is None: continue
        d = pd(r[5].value)
        if d and isinstance(r[3].value, (int, float)):
            # Shranimo mesec, znesek in del opisa za lažje ujemanje
            keyed.append((d.month, round(float(r[3].value), 2), norm(str(v0 or "")[:60])))
    return keyed, Counter((m, a) for m, a, _ in keyed)

def already_exists(t, keyed, amt_counter):
    \"\"\"Preveri, ali je transakcija že vstavljena v Excel.\"\"\"
    m, a = t["d"].month, round(t["v"], 2)
    c = amt_counter.get((m, a), 0)
    if c == 0:    return False
    if c >= 3:    return True # Če imamo že 3 ista zneska v mesecu, verjetno je duplikat
    short = norm(t["desc"])[:10]
    # Če ima opis dovolj znakov, preveri, če se ujema z obstoječimi
    if len(short) >= 3:
        return any(mm == m and aa == a and short in nm for mm, aa, nm in keyed)
    return False

# ── Vstavljanje ────────────────────────────────────────────────────────────

def find_insert(ws, newd):
    \"\"\"Poišče primeren indeks za vstavljanje, da ohranimo kronološki red.\"\"\"
    for i in range(2, ws.max_row + 1):
        d = pd(ws.cell(i, 6).value)
        if d and d < newd:
            return i
    return ws.max_row + 1

def make_row(t):
    \"\"\"Ustvari seznam vrednosti za novo vrstico v Excelu.\"\"\"
    d, v, cur, desc = t["d"], t["v"], t["cur"], t["desc"]
    txt  = f"NAKUP {desc} {v:,.2f} {cur}".replace(",", " ")
    sign = "+" if t["neg"] else "-"
    return [txt, t["cat"], sign, round(v, 2), cur,
            fd(d), "","","","","","","",fd(d + timedelta(days=2)), "", ""]

def insert_all(ws, items):
    \"\"\"Vstavi vse nove transakcije v Excel po datumu.\"\"\"
    n = 0
    for t in sorted(items, key=lambda x: x["d"]):
        idx = find_insert(ws, t["d"])
        ws.insert_rows(idx, 1)
        for c, val in enumerate(make_row(t), start=1):
            ws.cell(idx, c).value = val
        n += 1
    return n

# ── VISA_PREMIUM odstranitev ───────────────────────────────────────────────

def delete_visa_premium(ws):
    \"\"\"Odstrani vrstice, ki so povzemki (npr. VISA_PREMIUM).\"\"\"
    to_del = [i for i in range(2, ws.max_row + 1)
              if ws.cell(i,1).value and
                 str(ws.cell(i,1).value).replace(" ","").replace("_","").upper() == "VISAPREMIUM"]
    for i in sorted(to_del, reverse=True):
        ws.delete_rows(i, 1)
    return len(to_del)

# ── Verifikacija ───────────────────────────────────────────────────────────

def verify(ws):
    \"\"\"Preveri, ali so datumi v Excelu pravilno razvrščeni (kronološko).\"\"\"
    dates = [pd(ws.cell(i,6).value) for i in range(2, ws.max_row+1)]
    dates = [d for d in dates if d]
    violations = sum(1 for i in range(len(dates)-1) if dates[i] < dates[i+1])
    return len(dates), violations

# ── Glavni program ─────────────────────────────────────────────────────────

def main():
    \"\"\"Glavna funkcija, ki vodi celoten proces.\"\"\"
    XLSX      = find_xlsx()
    XLSX_PATH = os.path.join(DIR, XLSX)

    print("\n" + "=" * 52)
    print(" KARTIČNI IZPISKI — SAMODEJNA OBDELAVA")
    print("=" * 52)
    print(f"  Mapa : {DIR}")
    print(f"  Excel: {XLSX}\n")

    # 1 · Poišči nove PDF-e
    all_pdfs = sorted(f for f in os.listdir(os.path.join(DIR, SUBDIR))
                      if f.lower().endswith(".pdf") and f.lower().startswith(PDF_PREFIX.lower()))
    processed = load_processed()
    new_pdfs  = [p for p in all_pdfs if p not in processed]

    if not all_pdfs:
        print(f"  Ni PDF izpisov v mapi {SUBDIR}. Kopiraj Izpisek_*.pdf in poskusi znova.")
        return
    if not new_pdfs:
        print(f"  ✓  Vsi PDF izpiski ({len(all_pdfs)}) so že obdelani. Ni novih podatkov.")
        print(f"     (Če želiš reset: izbriši {STATE_FILE})")
        return

    print(f"  PDF-jev skupno : {len(all_pdfs)}")
    print(f"  Novih (za obdelavo): {len(new_pdfs)}")
    for p in new_pdfs:
        print(f"    +  {p}")
    print()

    # 2 · Backup
    backup = XLSX.replace(".xlsx", " - backup.xlsx")
    shutil.copy(XLSX_PATH, os.path.join(DIR, backup))
    print(f"  ✓  Backup → {os.path.basename(backup)}")

    # 3 · Odpri Excel
    wb = openpyxl.load_workbook(XLSX_PATH)
    ws = wb[SHEET_NAME]
    before = ws.max_row - 1
    print(f"  ✓  Excel odprt: {before} vrstic podatkov")

    # 4 · Parsaj PDFe
    all_txns = []
    for p in new_pdfs:
        txns = parse_pdf(os.path.join(DIR, SUBDIR, p))
        print(f"  ✓  {p}: {len(txns)} transakcij")
        all_txns.extend(txns)

    if not all_txns:
        processed.update(new_pdfs)
        save_processed(processed)
        print("\n  ⚠  Nobene transakcije ni bilo mogoče razbrati iz PDF-jev.")
        print("     Preveri format izpisa ali posodobi regex v process.py.")
        return

    # 5 · Dedup
    keyed, amt_counter = build_existing(ws)
    to_add, skipped = [], 0
    for t in all_txns:
        t["cat"] = categorize(t["desc"])
        if already_exists(t, keyed, amt_counter):
            skipped += 1
        else:
            to_add.append(t)
    print(f"  ✓  Dedup: preskočenih {skipped} | za vnos {len(to_add)}")

    if not to_add:
        processed.update(new_pdfs)
        save_processed(processed)
        print("\n  Vse transakcije iz novih PDF-jev so že v Excel-u.")
        return

    # 6 · Vstavi
    n_insert = insert_all(ws, to_add)
    print(f"  ✓  Vstavljenih vrstic: {n_insert}")

    # 7 · Odstrani VISA_PREMIUM (skupni mesčni zneski)
    n_del = delete_visa_premium(ws)
    if n_del:
        print(f"  ✓  Odstranjenih VISA_PREMIUM vrstic: {n_del}")

    # 8 · Shrani
    wb.save(XLSX_PATH)
    print(f"  ✓  Excel shranjen")

    # 9 · Verifikacija
    rows, violations = verify(openpyxl.load_workbook(XLSX_PATH)[SHEET_NAME])
    print(f"  ✓  Verifikacija: {rows} vrstic, sort-napake: {violations}")

    # 10 · Označi PDF-e kot obdelane
    processed.update(new_pdfs)
    save_processed(processed)

    # Povzetek
    print()
    print("=" * 52)
    print(f"  KONČANO")
    print(f"    Pred   : {before} vrstic")
    print(f"    Dodano : +{n_insert}  |  Odstranjeno: −{n_del}")
    print(f"    Po     : {before + n_insert - n_del} vrstic")
    print(f"    Sort   : {'✓  brez prekršitev' if violations == 0 else f'⚠  {violations} prekršitev!'}")
    print("=" * 52 + "\n")

if __name__ == "__main__":
    main()
