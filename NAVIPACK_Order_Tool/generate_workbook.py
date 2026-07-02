import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

# ---------------- MASTER PRODUCT DATA ----------------
# desc, colour, units/box, boxes/row, rows/pallet, box_vol(m3), weight_gr,
# price L, M, N (per 1000 pcs), family, assumed_arrangement(bool)
P = [
 ("PREMIUM FORK",          "TRANSPARENT / BLACK",       2000, 7, 10, 0.024, 2.8,  6.25, 7.00, 7.50,  "PREMIUM",  False),
 ("PREMIUM SPOON",         "TRANSPARENT / BLACK",       2000, 7,  8, 0.028, 2.8,  6.25, 7.00, 7.50,  "PREMIUM",  False),
 ("PREMIUM KNIFE",         "TRANSPARENT / BLACK",       2000, 7, 11, 0.018, 2.8,  6.25, 7.00, 7.50,  "PREMIUM",  False),
 ("PREMIUM DESSERT SPOON", "TRANSPARENT / BLACK",       2000, 7, 10, 0.019, 2.5,  6.00, 6.70, 7.25,  "PREMIUM",  False),
 ("DIAMOND FORK",          "TRANSPARENT / WHITE / BLACK",2000,12, 10, 0.024, 3.8,  8.25, 9.25,10.00,  "DIAMOND",  False),
 ("DIAMOND SPOON",         "TRANSPARENT / WHITE / BLACK",2000,12,  8, 0.031, 3.8,  8.25, 9.25,10.00,  "DIAMOND",  False),
 ("DIAMOND KNIFE",         "TRANSPARENT / WHITE / BLACK",2000,12, 12, 0.020, 3.8,  8.25, 9.25,10.00,  "DIAMOND",  False),
 ("SMART FORK",            "TRANSPARENT / BLACK",       2000,12, 10, 0.025, 3.7,  8.75, 9.00, 9.75,  "SMART",    True),
 ("WAVY ICE CREAM SPOON",  "COLORED (RED/BLUE/GREEN/ORANGE)",2000,9, 7, 0.025, 2.4, 7.00, 7.85, 8.50, "ICE CREAM", False),
 ("120 CC CHAMPAGNE GLASS","TRANSPARENT (FOOT COLORED)", 300, 6,  6, 0.053,12.5, 80.0, 90.0,100.0,  "GLASS",    False),
 ("170 CC MINI WINE GLASS","TRANSPARENT (FOOT COLORED)", 300, 6,  6, 0.049,12.5,111.0,105.0,100.0,  "GLASS",    True),
 ("185 CC WINE GLASS",     "TRANSPARENT (FOOT COLORED)", 300, 6,  6, 0.060,14.5,130.0,120.0,110.0,  "GLASS",    True),
 ("PIZZA TRIPOD",          "WHITE",                     1000, 9, 10, 0.027, 1.85, 6.50, 7.30, 8.50,  "PIZZA",    False),
]

# ---------------- STYLE HELPERS ----------------
BRAND   = "1F3864"   # dark navy
BRAND2  = "2E5496"
HEADER  = "305496"
INPUTBG = "FFF2CC"   # light yellow (input)
INPUTHD = "BF8F00"   # darker gold header for input col
OUTBG   = "E2EFDA"   # light green (outputs)
REFBG   = "F2F2F2"   # light grey (reference)
ASSUMED = "FCE4D6"   # light orange (assumed / confirm)
WHITE   = "FFFFFF"

thin = Side(style="thin", color="BFBFBF")
med  = Side(style="medium", color="808080")
border_all = Border(left=thin,right=thin,top=thin,bottom=thin)
border_box = Border(left=med,right=med,top=med,bottom=med)

def style_header(c, fill=HEADER, color=WHITE, size=10):
    c.font = Font(bold=True, color=color, size=size, name="Calibri")
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = border_all

def money(c): c.number_format = '#,##0.00'
def num0(c):  c.number_format = '#,##0'
def vol(c):   c.number_format = '0.000'
def wt(c):    c.number_format = '#,##0.0'
def pal(c):   c.number_format = '0.00'

wb = openpyxl.Workbook()

# ============================================================
# SHEET 1 : ORDER FORM
# ============================================================
ws = wb.active
ws.title = "SIPARIS - ORDER"

# ---- Title block ----
ws.merge_cells("A1:D3")
t = ws["A1"]
t.value = "NAVIPACK LLC\nSIPARIS & HACIM HESAPLAMA / ORDER & VOLUME CALCULATOR"
t.font = Font(bold=True, color=WHITE, size=13, name="Calibri")
t.fill = PatternFill("solid", fgColor=BRAND)
t.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

ws.merge_cells("E1:J1"); ws["E1"] = "MUSTERI / CUSTOMER:"; ws["E1"].font=Font(bold=True)
ws.merge_cells("K1:R1"); ws["K1"] = ""; ws["K1"].fill=PatternFill("solid",fgColor=INPUTBG); ws["K1"].border=border_all
ws.merge_cells("E2:J2"); ws["E2"] = "TARIH / DATE:"; ws["E2"].font=Font(bold=True)
ws.merge_cells("K2:R2"); ws["K2"] = "=TODAY()"; ws["K2"].number_format="dd.mm.yyyy"; ws["K2"].fill=PatternFill("solid",fgColor=INPUTBG); ws["K2"].border=border_all
ws.merge_cells("E3:J3"); ws["E3"] = "FIYAT KADEMESI (1=L / 2=M / 3=N) / PRICE TIER:"; ws["E3"].font=Font(bold=True)
tier = ws["K3"]; tier.value = 3
tier.fill=PatternFill("solid",fgColor=INPUTBG); tier.border=border_box
tier.alignment=Alignment(horizontal="center"); tier.font=Font(bold=True, size=12)
ws.merge_cells("L3:R3")
ws["L3"]="<-- Buradan 1, 2 veya 3 secin. Secilen kademe J sutununu ve tutari belirler."
ws["L3"].font=Font(italic=True, size=9, color="808080")
dv_tier = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False)
ws.add_data_validation(dv_tier); dv_tier.add(tier)

# ---- Column headers (row 5) ----
headers = [
 ("A","NO",5),
 ("B","URUN / PRODUCT",26),
 ("C","RENK / COLOUR",22),
 ("D","ADET / KOLI\n(units/box)",10),
 ("E","KOLI / SIRA\n(box/row)",9),
 ("F","SIRA / PALET\n(row/pallet)",9),
 ("G","KOLI / PALET\n(box/pallet)",9),
 ("H","KOLI HACMI\nm3 (vol/box)",10),
 ("I","KOLI NET kg\n(net/box)",9),
 ("J","BIRIM FIYAT\n/1000 (secili)",11),
 ("K","SIPARIS\n(KOLI)  ***",12),
 ("L","SIPARIS\n(ADET)",11),
 ("M","TAM SIRA\n(full rows)",9),
 ("N","PALET\n(esdeger)",9),
 ("O","NET AGIRLIK\nkg",11),
 ("P","HACIM\nm3",9),
 ("Q","TUTAR\nUSD",13),
 ("R","Fiyat-1 (L)",9),
 ("S","Fiyat-2 (M)",9),
 ("T","Fiyat-3 (N)",9),
 ("U","AILE / FAMILY",13),
]
HR = 5
for col,txt,w in headers:
    ws.column_dimensions[col].width = w
    c = ws[f"{col}{HR}"]; c.value = txt
    if col == "K":
        style_header(c, fill=INPUTHD)
    elif col in ("L","M","N","O","P","Q"):
        style_header(c, fill="548235")
    elif col in ("R","S","T","U"):
        style_header(c, fill="A6A6A6", size=9)
    else:
        style_header(c)
ws.row_dimensions[HR].height = 30

# ---- Data rows ----
first = HR+1
for i,row in enumerate(P):
    r = first+i
    (desc,colour,upb,bpr,rpp,bvol,wgr,pl,pm,pn,fam,assumed) = row
    ws[f"A{r}"] = i+1
    ws[f"B{r}"] = desc
    ws[f"C{r}"] = colour
    ws[f"D{r}"] = upb
    ws[f"E{r}"] = bpr
    ws[f"F{r}"] = rpp
    ws[f"G{r}"] = f"=E{r}*F{r}"
    ws[f"H{r}"] = bvol
    ws[f"I{r}"] = f"=D{r}*{wgr}/1000"     # net kg per box = units*gr/1000
    ws[f"R{r}"] = pl
    ws[f"S{r}"] = pm
    ws[f"T{r}"] = pn
    ws[f"J{r}"] = f"=CHOOSE($K$3,R{r},S{r},T{r})"
    ws[f"U{r}"] = fam
    # input
    ws[f"K{r}"] = 0
    # outputs
    ws[f"L{r}"] = f"=K{r}*D{r}"
    ws[f"M{r}"] = f"=IF(E{r}=0,0,K{r}/E{r})"
    ws[f"N{r}"] = f"=IF(G{r}=0,0,K{r}/G{r})"
    ws[f"O{r}"] = f"=K{r}*I{r}"
    ws[f"P{r}"] = f"=K{r}*H{r}"
    ws[f"Q{r}"] = f"=L{r}/1000*J{r}"

    # number formats
    num0(ws[f"D{r}"]); vol(ws[f"H{r}"]); wt(ws[f"I{r}"])
    money(ws[f"J{r}"]); num0(ws[f"K{r}"]); num0(ws[f"L{r}"])
    ws[f"M{r}"].number_format='0'; pal(ws[f"N{r}"]); wt(ws[f"O{r}"]); vol(ws[f"P{r}"])
    money(ws[f"Q{r}"]); money(ws[f"R{r}"]); money(ws[f"S{r}"]); money(ws[f"T{r}"])

    # fills / borders
    for col in "ABCDEFGHIJU":
        cc=ws[f"{col}{r}"]; cc.border=border_all
        if col in "RST": cc.fill=PatternFill("solid",fgColor=REFBG)
    for col in "RST":
        ws[f"{col}{r}"].fill=PatternFill("solid",fgColor=REFBG); ws[f"{col}{r}"].border=border_all
    ws[f"U{r}"].fill=PatternFill("solid",fgColor=REFBG)
    kc=ws[f"K{r}"]; kc.fill=PatternFill("solid",fgColor=INPUTBG); kc.border=border_box
    kc.font=Font(bold=True); kc.alignment=Alignment(horizontal="center")
    for col in "LMNOPQ":
        ws[f"{col}{r}"].fill=PatternFill("solid",fgColor=OUTBG); ws[f"{col}{r}"].border=border_all
    ws[f"B{r}"].font=Font(bold=True)
    ws[f"A{r}"].alignment=Alignment(horizontal="center")
    ws[f"C{r}"].alignment=Alignment(horizontal="left", wrap_text=True)
    ws[f"C{r}"].font=Font(size=9)
    for col in "DEFGHI": ws[f"{col}{r}"].alignment=Alignment(horizontal="center")

    # highlight assumed arrangements + comment
    if assumed:
        for col in ("E","F","G"):
            ws[f"{col}{r}"].fill=PatternFill("solid",fgColor=ASSUMED)
        cm = Comment("DIKKAT: Bu urun eldeki volume tablolarinda yok. Palet dizilimi "
                     "(koli/sira, sira/palet) benzer urunlere gore TAHMIN edilmistir. "
                     "Lutfen fabrikadan teyit ederek E ve F hucrelerini guncelleyin.\n\n"
                     "NOTE: arrangement is ASSUMED - please confirm with the factory.", "NAVIPACK")
        cm.width=300; cm.height=140
        ws[f"E{r}"].comment = cm

last = first+len(P)-1

# ---- Totals row ----
tr = last+1
ws[f"A{tr}"]=""; 
ws.merge_cells(f"A{tr}:J{tr}")
ws[f"A{tr}"]="TOPLAM / TOTAL"
ws[f"A{tr}"].font=Font(bold=True, color=WHITE, size=11)
ws[f"A{tr}"].fill=PatternFill("solid",fgColor=BRAND)
ws[f"A{tr}"].alignment=Alignment(horizontal="right", vertical="center")
ws[f"K{tr}"]=f"=SUM(K{first}:K{last})"
ws[f"L{tr}"]=f"=SUM(L{first}:L{last})"
ws[f"N{tr}"]=f"=SUM(N{first}:N{last})"
ws[f"O{tr}"]=f"=SUM(O{first}:O{last})"
ws[f"P{tr}"]=f"=SUM(P{first}:P{last})"
ws[f"Q{tr}"]=f"=SUM(Q{first}:Q{last})"
num0(ws[f"K{tr}"]); num0(ws[f"L{tr}"]); pal(ws[f"N{tr}"]); wt(ws[f"O{tr}"]); vol(ws[f"P{tr}"]); money(ws[f"Q{tr}"])
for col in "KLMNOPQ":
    c=ws[f"{col}{tr}"]; c.font=Font(bold=True); c.fill=PatternFill("solid",fgColor="D9E1F2"); c.border=border_box
ws.row_dimensions[tr].height=20

# note about pallets under totals
nr = tr+1
ws.merge_cells(f"A{nr}:Q{nr}")
ws[f"A{nr}"]=("*** SIPARIS (KOLI) sutununa yalnizca ilgili urunun KOLI/SIRA degerinin KATLARI girilebilir "
             "(1 palet sirasi = 1 tam sira). Ornek: Diamond Fork icin 12,24,36,48... | "
             "PALET (esdeger) = yarim paletler dahil toplam palet hacmi. Gercek fiziksel palet sayisi icin 'PALET OZETI' sayfasina bakin.")
ws[f"A{nr}"].font=Font(italic=True, size=9, color="C00000")
ws[f"A{nr}"].alignment=Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[nr].height=42

# ---- Data validation: order must be multiple of boxes/row ----
dv = DataValidation(
    type="custom",
    formula1=f"=AND(K{first}>=0,MOD(K{first},E{first})=0)",
    allow_blank=True, showInputMessage=True, showErrorMessage=True)
dv.promptTitle = "SIPARIS ADEDI / ORDER (BOXES)"
dv.prompt = ("Bu urunu KOLI cinsinden girin. Deger, KOLI/SIRA sutununun (E) KATLARI olmalidir.\n"
             "Cunku her palet sirasi tek urunle DOLU olmali; boylece kalan yarim paletler duzgun birlestirilir.\n"
             "Ornek Diamond Fork (12/sira): 12, 24, 36, 48, 60 ...\n\n"
             "Enter the order in BOXES. Must be a MULTIPLE of BOXES/ROW (column E).")
dv.errorTitle = "GECERSIZ ADET"
dv.error = ("Girdiginiz sayi KOLI/SIRA degerinin (E sutunu) tam kati degil!\n"
            "Lutfen tam sira olacak sekilde girin (or. 12'nin katlari: 12,24,36,48...).")
ws.add_data_validation(dv)
dv.sqref = f"K{first}:K{last}"

# header comment on K
kh = ws[f"K{HR}"]
kc = Comment("Siparisi KOLI (koli) cinsinden girin. Her hucre yalnizca KOLI/SIRA'nin katlarini kabul eder "
             "- boylece her palet sirasi tek urunle dolar ve yarim paletler duzgun birlesir.", "NAVIPACK")
kc.width=280; kc.height=110
kh.comment = kc

ws.freeze_panes = "A6"
ws.sheet_view.showGridLines = False

# ============================================================
# SHEET 2 : PALLET SUMMARY (consolidation by family)
# ============================================================
ps = wb.create_sheet("PALET OZETI - PALLETS")
ps.sheet_view.showGridLines = False
ps.merge_cells("A1:F2")
ps["A1"]="PALET OZETI / PALLET CONSOLIDATION SUMMARY"
ps["A1"].font=Font(bold=True,color=WHITE,size=13)
ps["A1"].fill=PatternFill("solid",fgColor=BRAND)
ps["A1"].alignment=Alignment(horizontal="left",vertical="center")

ps.merge_cells("A3:F4")
ps["A3"]=("Yarim paletler ONCELIKLE ayni urun ailesi icinde birlestirilir (Diamond'lar Diamond'larla, "
          "Premium'lar Premium'larla, kasiklar kendi icinde...). Asagidaki 'GERCEK PALET' sutunu her aile icin "
          "yukari yuvarlanmis fiziksel palet sayisidir. / Half-pallets are combined within the same product family first.")
ps["A3"].font=Font(italic=True,size=9,color="404040")
ps["A3"].alignment=Alignment(wrap_text=True,vertical="center")

ph = 6
phead = ["AILE / FAMILY","SIPARIS KOLI\n(boxes)","PALET ESDEGER\n(pallet equiv.)","TAM PALET\n(full)","GERCEK PALET\n(actual, rounded up)","NET AGIRLIK kg"]
widths=[24,14,15,12,16,15]
for j,(txt,w) in enumerate(zip(phead,widths)):
    col=get_column_letter(j+1)
    ps.column_dimensions[col].width=w
    c=ps[f"{col}{ph}"]; c.value=txt; style_header(c, fill=HEADER)
ps.row_dimensions[ph].height=30

families = ["DIAMOND","PREMIUM","SMART","ICE CREAM","GLASS","PIZZA"]
famlabel = {"DIAMOND":"DIAMOND CUTLERY","PREMIUM":"PREMIUM CUTLERY","SMART":"SMART FORK",
            "ICE CREAM":"ICE CREAM SPOON","GLASS":"GLASSES / CUPS","PIZZA":"PIZZA TRIPOD"}
oref = "'SIPARIS - ORDER'"
prow0 = ph+1
for k,fam in enumerate(families):
    r=prow0+k
    ps[f"A{r}"]=famlabel[fam]
    ps[f"B{r}"]=f"=SUMIF({oref}!$U${first}:$U${last},$A{r}_x,{oref}!$K${first}:$K${last})"
    # need family CODE match; store code in helper column H
    ps[f"H{r}"]=fam
    ps[f"B{r}"]=f"=SUMIF({oref}!$U${first}:$U${last},$H{r},{oref}!$K${first}:$K${last})"
    ps[f"C{r}"]=f"=SUMIF({oref}!$U${first}:$U${last},$H{r},{oref}!$N${first}:$N${last})"
    ps[f"D{r}"]=f"=INT(C{r})"
    ps[f"E{r}"]=f"=IF(C{r}=0,0,CEILING(C{r},1))"
    ps[f"F{r}"]=f"=SUMIF({oref}!$U${first}:$U${last},$H{r},{oref}!$O${first}:$O${last})"
    num0(ps[f"B{r}"]); pal(ps[f"C{r}"]); ps[f"D{r}"].number_format='0'; ps[f"E{r}"].number_format='0'; wt(ps[f"F{r}"])
    ps[f"A{r}"].font=Font(bold=True)
    for col in "ABCDEF":
        cc=ps[f"{col}{r}"]; cc.border=border_all
        if col=="E": cc.fill=PatternFill("solid",fgColor=OUTBG); cc.font=Font(bold=True)
    ps[f"H{r}"].font=Font(color="FFFFFF")  # hide code visually
ps.column_dimensions["H"].hidden = True

ptot=prow0+len(families)
ps[f"A{ptot}"]="TOPLAM / GRAND TOTAL"
ps[f"A{ptot}"].font=Font(bold=True,color=WHITE)
ps[f"A{ptot}"].fill=PatternFill("solid",fgColor=BRAND)
ps[f"A{ptot}"].alignment=Alignment(horizontal="right")
ps[f"B{ptot}"]=f"=SUM(B{prow0}:B{ptot-1})"
ps[f"C{ptot}"]=f"=SUM(C{prow0}:C{ptot-1})"
ps[f"D{ptot}"]=f"=SUM(D{prow0}:D{ptot-1})"
ps[f"E{ptot}"]=f"=SUM(E{prow0}:E{ptot-1})"
ps[f"F{ptot}"]=f"=SUM(F{prow0}:F{ptot-1})"
num0(ps[f"B{ptot}"]); pal(ps[f"C{ptot}"]); ps[f"D{ptot}"].number_format='0'; ps[f"E{ptot}"].number_format='0'; wt(ps[f"F{ptot}"])
for col in "BCDEF":
    c=ps[f"{col}{ptot}"]; c.font=Font(bold=True); c.fill=PatternFill("solid",fgColor="D9E1F2"); c.border=border_box

note_r=ptot+2
ps.merge_cells(f"A{note_r}:F{note_r+2}")
ps[f"A{note_r}"]=("NOT: 'GERCEK PALET', her ailenin toplam palet-esdegerinin yukari yuvarlanmasiyla bulunur "
                  "(yarim paletler ayni aile icinde birlesir). Farkli aileler ancak koli yukseklikleri uyumluysa "
                  "ayni palette birlestirilebilir - bu durumda toplam daha da dusebilir. Nihai istifleme fabrikada "
                  "fiziksel olarak teyit edilmelidir. Her palet sirasi TEK urunle dolmalidir.")
ps[f"A{note_r}"].font=Font(italic=True,size=9,color="404040")
ps[f"A{note_r}"].alignment=Alignment(wrap_text=True,vertical="top")

# ============================================================
# SHEET 3 : NOTES / INSTRUCTIONS
# ============================================================
nt = wb.create_sheet("ACIKLAMA - NOTES")
nt.sheet_view.showGridLines=False
nt.column_dimensions["A"].width=3
nt.column_dimensions["B"].width=110
nt.merge_cells("B1:B1")
lines = [
 ("NAVIPACK - SIPARIS & HACIM HESAPLAMA TABLOSU / ORDER & VOLUME CALCULATOR", "title"),
 ("", ""),
 ("NASIL KULLANILIR / HOW TO USE", "h"),
 ("1) 'SIPARIS - ORDER' sayfasinda SADECE sari 'SIPARIS (KOLI)' sutununa (K) veri girin.", "n"),
 ("   Diger tum sutunlar otomatik hesaplanir. Ust taraftan Musteri, Tarih ve Fiyat Kademesi secin.", "n"),
 ("2) Siparis KOLI (koli) cinsinden girilir. Her urun icin girilen deger, o urunun", "n"),
 ("   KOLI/SIRA degerinin KATI olmalidir (1 palet sirasi = 1 tam sira).", "n"),
 ("   Ornek: Diamond Fork 1 sirada 12 koli -> yalnizca 12, 24, 36, 48, 60 ... girilebilir.", "n"),
 ("   (Yanlis kat girilirse Excel uyari verir.)", "n"),
 ("", ""),
 ("NEDEN KATLAR? / WHY MULTIPLES OF A ROW?", "h"),
 ("Musteri her zaman paletli (100x120 cm buyuk palet) yukleme yapar. Kalan yarim paletler,", "n"),
 ("diger yarim paletlerle birlestirilir. Bir palet sirasina FARKLI urunler dizilirse, koli", "n"),
 ("yukseklikleri farkli olacagindan palet dengesiz yukselir. Bu yuzden her sira TEK urunle", "n"),
 ("DOLU olmalidir; yani siparis daima tam sira (KOLI/SIRA'nin kati) olmalidir.", "n"),
 ("", ""),
 ("ADET -> KOLI DONUSUMU / UNITS -> BOXES", "h"),
 ("Siparisler cogu zaman ADET olarak konusulur. 1 kolideki adet 'ADET/KOLI' sutunundadir.", "n"),
 ("Ornek: 600.000 adet Diamond Fork / 2.000 adet-koli = 300 koli. 1 palet = 120 koli ->", "n"),
 ("300 koli = 2 tam palet + yarim palet (60 koli). Tabloya KOLI olarak 300 girilir (12'nin kati).", "n"),
 ("", ""),
 ("YARIM PALETLERIN BIRLESTIRILMESI / HALF-PALLET CONSOLIDATION", "h"),
 ("Oncelik ayni urun ailesidir: Diamond'lar Diamond'larla (siyah+seffaf), Premium'lar", "n"),
 ("Premium'larla, kasiklar kasiklarla birlestirilir. 'PALET OZETI' sayfasi bu birlestirmeyi", "n"),
 ("aile bazinda yapar ve her aile icin GERCEK (yukari yuvarlanmis) palet sayisini gosterir.", "n"),
 ("", ""),
 ("HESAPLAMALAR / FORMULAS", "h"),
 ("SIPARIS (ADET)   = Siparis Koli x Adet/Koli", "n"),
 ("TAM SIRA         = Siparis Koli / (Koli/Sira)", "n"),
 ("PALET (esdeger)  = Siparis Koli / (Koli/Palet)   [yarim paletler dahil hacim]", "n"),
 ("NET AGIRLIK kg   = Siparis Koli x Koli Net kg     (Koli Net = Adet/Koli x urun agirligi gr /1000)", "n"),
 ("HACIM m3         = Siparis Koli x Koli Hacmi m3", "n"),
 ("TUTAR USD        = (Siparis Adet /1000) x Birim Fiyat (secili kademe)", "n"),
 ("", ""),
 ("RENK KODLARI / COLOUR LEGEND", "h"),
 ("Sari  = veri girisi (siz doldurun) | Yesil = otomatik sonuc | Gri = referans (fiyat/aile)", "n"),
 ("Turuncu (Koli/Sira, Sira/Palet) = palet dizilimi eldeki volume tablolarinda YOK, TAHMINdir - teyit edin.", "n"),
 ("", ""),
 ("KAYNAKLAR / SOURCES", "h"),
 ("Urun listesi & fiyatlar: NAVIPACK Price Offer 10.06.2026.", "n"),
 ("Palet dizilimleri: NAVIPACK Volume Calculation tablolari (11.08.2025 & 01.09.2025).", "n"),
]
r=1
for txt,kind in lines:
    cell=nt[f"B{r}"]
    cell.value=txt
    if kind=="title":
        cell.font=Font(bold=True,color=WHITE,size=13); cell.fill=PatternFill("solid",fgColor=BRAND)
        nt.row_dimensions[r].height=24
        cell.alignment=Alignment(vertical="center")
    elif kind=="h":
        cell.font=Font(bold=True,color=BRAND,size=11)
        cell.fill=PatternFill("solid",fgColor="D9E1F2")
    else:
        cell.font=Font(size=10)
        cell.alignment=Alignment(wrap_text=False)
    r+=1

# order the sheets: Notes first? keep order form first
wb.active = wb.sheetnames.index("SIPARIS - ORDER")

out = "NAVIPACK_Siparis_Hacim_Hesaplama.xlsx"
wb.save(out)
print("SAVED", out)
