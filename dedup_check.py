# -*- coding: utf-8 -*-
import re, os, sys
os.chdir(r"d:\BVDoc\Transakcije2026")
sys.stdout.reconfigure(encoding="utf-8")
import unicodedata
from pypdf import PdfReader
import warnings; warnings.filterwarnings("ignore")
import openpyxl
from datetime import date, timedelta

AMT=re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}$")
RX=re.compile(r"^(\d{2}\.\d{2}\.\d{4})\s+(.+?)\s+(\d{1,3}(?:\.\d{3})*,\d{2})(?:[+\-])?\s*$")
SKIP=["BERNARD","VISAPREMIUM","VISA_PREMIUM","474819","OŽIJEK","OZIJEK","PORABA","PLAČA","PLACAN","ZARAČUN","ZARACUN","ZNESEK","KARTIC","SKUPAJ","POVZETEK","PLAČATEL","PLAČILO","NARAVNA","STORITVE","VRACANJE","KREDITNO","DEBITNO","OBLIČJE","VRSTA","KARTICA"]
FX={"DKK","SEK","CHF","USD","GBP","PLN","CZK","HUF","CNY","JPY","EUR"}
PDFS=["Izpisek_2026_10_11_12_56_25.pdf","Izpisek_2026_10_11_12_56_25 (1).pdf","Izpisek_2026_10_11_12_56_25 (2).pdf","Izpisek_2026_10_11_12_56_25 (3).pdf","Izpisek_2026_10_11_12_56_25 (4).pdf","Izpisek_2026_10_11_13_05_21.pdf","Izpisek_2026_10_11_12_56_25 (5).pdf","Izpisek_2026_10_11_12_56_25 (6).pdf","Izpisek_2026_10_11_12_56_25 (7).pdf"]

def parse_pdf(path):
    out=[]
    for pg in PdfReader(path).pages:
        for ln in pg.extract_text().splitlines():
            ln=re.sub(" +"," ",ln.strip())
            if not ln or any(k in ln for k in SKIP): continue
            m=RX.match(ln)
            if not m: continue
            dt,desc,val=m.groups()
            d,mo,y=[int(x) for x in dt.split(".")]
            v=float(val.replace(".","").replace(",","."))
            neg=ln.rstrip().endswith("-")
            cur="EUR"
            toks=desc.replace(","," ").upper().split()
            for t in toks:
                if t in FX: cur=t; break
            desc=re.sub(r"\bEUR\b","",desc)
            desc=re.sub(r"\s+"," ",desc).strip().rstrip(",")
            out.append({"d":date(y,mo,d),"desc":desc,"v":v,"cur":cur,"neg":neg})
    return out

def norm(s):
    s="".join(c for c in unicodedata.normalize("NFD",s) if unicodedata.category(c)!="Mn").lower()
    return re.sub(r"[^a-z0-9]","",s)

CAT=[
 ("Šport & Prosti čas", ["extremevital","steam","eventim","sportvision","sport"]),
 ("Avto & Prevozi", ["petrol","inamarket","bsmol","tamoyl","tamoil","zminj","autopark","petrolio","span","drivr","uber","wallismotor","sixt","sixinz","europeanpark","petrolm","euroam","motocenter","span"]),
 ("Dom", ["suncontract","rtvslovenija","teleslovenija","telemach","komunalne","slodkovec","sodis"]),
 ("Finance & Zavarovanja", ["zavarovalnica","zavarov"]),
 ("Kavarne & Restavracije", ["kavarna","brunch","restavracija","restaurant","boreka","laturka","latureka","burek","lahocolate","chutis","hoodburger","domacapeka","nyhavn38","ubereats","wolt","foetex","belacenter","pressbyran","skastnmalmo","twoforks","heinemann","tubik","gastro"]),
]
def cat(desc):
    d=norm(desc)
    for c,kws in CAT:
        for k in kws:
            if k in d: return c
    return "Nakupi"

def fd(d): return f"{d.day}. {d.month}. {d.year}"

allt=[]
for p in PDFS: allt.extend(parse_pdf(p))
print("total pdf txns:",len(allt))

wb=openpyxl.load_workbook(r"Transakcije (5).xlsx")
ws=wb["transakcije"]
existing=[]
for r in ws.iter_rows(min_row=2):
    v0=r[0].value
    if v0 is None and r[3].value is None: continue
    existing.append({"txt":str(v0),"num":r[3].value,"sign":r[2].value,"cat":r[1].value,"date":r[5].value})
def pd(s):
    p=[int(x) for x in re.findall(r"\d+",str(s))]
    return date(p[2],p[1],p[0]) if len(p)>=3 else None
keyed=[]
for e in existing:
    d=pd(e["date"])
    if d and isinstance(e["num"],(int,float)):
        keyed.append((d.month,round(float(e["num"]),2),norm((e["txt"] or "")[:60])))
print("sheet rows:",len(keyed))
from collections import Counter
mc=Counter((m,a) for m,a,n in keyed)
print("dup amts in month:",[(k,v) for k,v in mc.items() if v>=3][:10])

def already_exists(t):
    m,a=t["d"].month,round(t["v"],2)
    c=mc.get((m,a),0)
    if c==0: return False
    if c>=3: return True
    # 1 or 2: require merchant overlap
    for mm,aa,nm in keyed:
        if mm==m and aa==a:
            short=norm(t["desc"])[:10]
            if len(short)>=5 and short in nm: return True
    return False

to_add=[]; skipped=0
for t in allt:
    if already_exists(t): skipped+=1; continue
    t["cat"]=cat(t["desc"]); to_add.append(t)
print("skipped(exist):",skipped,"to add:",len(to_add))
import json
json.dump([{"d":t["d"].isoformat(),"desc":t["desc"],"v":t["v"],"cur":t["cur"],"neg":t["neg"],"cat":t["cat"]} for t in to_add],open("toadd.json","w",encoding="utf-8"),ensure_ascii=False)
for t in sorted(to_add,key=lambda x:x["d"]):
    print(t["d"].isoformat(),t["desc"][:45],t["v"],t["cur"],"NEG" if t["neg"] else "",t["cat"])
