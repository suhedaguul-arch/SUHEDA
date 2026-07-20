import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

# ---------------- MASTER DATA (NAVIPACK Price Offer 20.06.2026 + Volume 07.07.2026) ----------------
# desc, colour, pcs/pack(H), packs/box(I), box_vol(K), box/layer(L), layer/pallet(M), weight_gr(G), price/1000(O), family
P = [
 ("PREMIUM FORK","TRANSPARENT",50,80,0.046,7,10,2.8,7.5,"PREMIUM"),
 ("PREMIUM FORK","BLACK",50,80,0.046,7,10,2.8,7.5,"PREMIUM"),
 ("PREMIUM SPOON","TRANSPARENT",50,80,0.056,7,8,2.8,7.5,"PREMIUM"),
 ("PREMIUM SPOON","BLACK",50,80,0.056,7,8,2.8,7.5,"PREMIUM"),
 ("PREMIUM KNIFE","TRANSPARENT",50,80,0.035,7,11,2.8,7.5,"PREMIUM"),
 ("PREMIUM KNIFE","BLACK",50,80,0.035,7,11,2.8,7.5,"PREMIUM"),
 ("PREMIUM DESSERT SPOON","TRANSPARENT",50,80,0.036,7,9,2.5,7.25,"PREMIUM"),
 ("PREMIUM DESSERT SPOON","BLACK",50,80,0.036,7,9,2.5,7.25,"PREMIUM"),
 ("DIAMOND FORK","TRANSPARENT",50,40,0.024,12,10,3.8,10.0,"DIAMOND"),
 ("DIAMOND FORK","BLACK",50,40,0.024,12,10,3.8,10.0,"DIAMOND"),
 ("DIAMOND SPOON","TRANSPARENT",50,40,0.031,12,8,3.8,10.0,"DIAMOND"),
 ("DIAMOND SPOON","BLACK",50,40,0.031,12,8,3.8,10.0,"DIAMOND"),
 ("DIAMOND KNIFE","TRANSPARENT",50,40,0.020,12,12,3.8,10.0,"DIAMOND"),
 ("DIAMOND KNIFE","BLACK",50,40,0.020,12,12,3.8,10.0,"DIAMOND"),
 ("SMART FORK","TRANSPARENT / BLACK",50,40,0.025,12,10,3.7,9.75,"SMART"),
 ("WAVY ICE CREAM SPOON","COLORED (RED/BLUE/GREEN/ORANGE)",50,40,0.025,9,9,2.4,8.5,"ICE CREAM"),
 ("120 CC CHAMPAGNE GLASS","TRANSPARENT (FOOT COLORED)",6,50,0.053,6,6,12.5,100.0,"GLASS"),
 ("170 CC MINI WINE GLASS","TRANSPARENT (FOOT COLORED)",6,50,0.049,6,10,12.5,100.0,"GLASS"),
 ("185 CC WINE GLASS","TRANSPARENT (FOOT COLORED)",6,50,0.067,6,7,14.5,110.0,"GLASS"),
 ("PIZZA TRIPOD","WHITE",100,10,0.027,9,11,1.85,8.5,"PIZZA"),
]
TRUCK_DEFAULT = 90   # NAVIPACK standard TIR (m3); double-trailer = 120

# ---------------- STYLES ----------------
BRAND="1F3864"; HEADER="305496"; INPUTBG="FFF2CC"; INPUTHD="BF8F00"
OUTBG="E2EFDA"; OUTHD="548235"; REFBG="F2F2F2"; REFHD="A6A6A6"
KPIBG="DDEBF7"; WHITE="FFFFFF"
thin=Side(style="thin",color="BFBFBF"); med=Side(style="medium",color="808080")
b_all=Border(left=thin,right=thin,top=thin,bottom=thin)
b_box=Border(left=med,right=med,top=med,bottom=med)
def H(c,fill=HEADER,color=WHITE,size=10):
    c.font=Font(bold=True,color=color,size=size); c.fill=PatternFill("solid",fgColor=fill)
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=b_all
def money(c):c.number_format='#,##0.00'
def num0(c):c.number_format='#,##0'
def vol(c):c.number_format='0.000'
def wt(c):c.number_format='#,##0.0'
def pal(c):c.number_format='0.00'
def pct(c):c.number_format='0.0%'

wb=openpyxl.Workbook()

# ============================================================ SHEET 1 : ORDER
ws=wb.active; ws.title="SIPARIS - ORDER"; ws.sheet_view.showGridLines=False

# ---- Title ----
ws.merge_cells("A1:E4")
t=ws["A1"]; t.value="NAVIPACK LLC\nSIPARIS & HACIM HESAPLAMA\nORDER & VOLUME CALCULATOR"
t.font=Font(bold=True,color=WHITE,size=13); t.fill=PatternFill("solid",fgColor=BRAND)
t.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True)

def lbl(cell,txt):
    ws[cell]=txt; ws[cell].font=Font(bold=True,size=9); ws[cell].alignment=Alignment(horizontal="right",vertical="center")
def inp(cell,val,fmt=None):
    c=ws[cell]; c.value=val; c.fill=PatternFill("solid",fgColor=INPUTBG); c.border=b_box
    c.alignment=Alignment(horizontal="center",vertical="center"); c.font=Font(bold=True)
    if fmt:fmt(c); return c

lbl("F1","MUSTERI / CUSTOMER:");        ws.merge_cells("G1:I1"); inp("G1","")
lbl("F2","TARIH / DATE:");              inp("G2","=TODAY()"); ws["G2"].number_format="dd.mm.yyyy"
lbl("F3","TIR KAPASITESI (m3):");       inp("G3",TRUCK_DEFAULT)
_tc=Comment("NAVIPACK standart TIR = 90 m3\nCift dorse (double trailer) = 120 m3\nDolum HACME gore hesaplanir.","NAVIPACK")
_tc.width=220;_tc.height=80; ws["G3"].comment=_tc

# ---- KPI band (row 1-4, cols K..S) ----
kpis=[("K1","TOPLAM HACIM (m3)","=R{tr}","vol"),
      ("N1","TIR DOLULUK","=IF($G$3=0,0,R{tr}/$G$3)","pct"),
      ("K3","GEREKEN TIR ADEDI","=IF($G$3=0,0,CEILING(R{tr}/$G$3,1))","int"),
      ("N3","TOPLAM TUTAR (USD)","=S{tr}","money")]
# placeholders; fill after we know totals row

HR=6
cols=[("A","NO",5),("B","URUN / PRODUCT",24),("C","RENK / COLOUR",20),
 ("D","ADET/KOLI\n(pcs/box)",10),("E","KOLI HACMI\nm3",9),("F","KOLI NET\nkg",8),
 ("G","KOLI/SIRA\n(box/layer)",10),("H","SIRA/PALET\n(layer/pal)",10),("I","KOLI/PALET\n(box/pal)",9),
 ("J","FIYAT/1000\nUSD",10),
 ("K","SIPARIS\n(KOLI) ***",12),
 ("L","SIPARIS\n(ADET)",11),("M","PALET\n(esdeger)",9),("N","TAM SIRA\n(rows)",8),
 ("O","NET AGIRLIK\nkg",11),("P","HACIM\nm3",9),("Q","TUTAR\nUSD",12),
 ("R","AILE / FAMILY",13)]
for col,txt,w in cols:
    ws.column_dimensions[col].width=w
    c=ws[f"{col}{HR}"]; c.value=txt
    if col=="K": H(c,fill=INPUTHD)
    elif col in "LMNOPQ": H(c,fill=OUTHD)
    elif col=="R": H(c,fill=REFHD,size=9)
    else: H(c)
ws.row_dimensions[HR].height=30

first=HR+1
for i,row in enumerate(P):
    r=first+i
    desc,colour,pp,pb,bvol,bpl,lpp,wgr,price,fam=row
    ws[f"A{r}"]=i+1
    ws[f"B{r}"]=desc; ws[f"C{r}"]=colour
    ws[f"D{r}"]=pp*pb                 # pcs/box
    ws[f"E{r}"]=bvol
    ws[f"F{r}"]=f"=D{r}*{wgr}/1000"   # net kg/box
    ws[f"G{r}"]=bpl; ws[f"H{r}"]=lpp
    ws[f"I{r}"]=f"=G{r}*H{r}"         # box/pallet
    ws[f"J{r}"]=price
    ws[f"K{r}"]=0                     # INPUT boxes
    ws[f"L{r}"]=f"=K{r}*D{r}"
    ws[f"M{r}"]=f"=IF(I{r}=0,0,K{r}/I{r})"
    ws[f"N{r}"]=f"=IF(G{r}=0,0,K{r}/G{r})"
    ws[f"O{r}"]=f"=K{r}*F{r}"
    ws[f"P{r}"]=f"=K{r}*E{r}"
    ws[f"Q{r}"]=f"=L{r}/1000*J{r}"
    ws[f"R{r}"]=fam
    num0(ws[f"D{r}"]); vol(ws[f"E{r}"]); wt(ws[f"F{r}"]); money(ws[f"J{r}"])
    num0(ws[f"K{r}"]); num0(ws[f"L{r}"]); pal(ws[f"M{r}"]); ws[f"N{r}"].number_format='0'
    wt(ws[f"O{r}"]); vol(ws[f"P{r}"]); money(ws[f"Q{r}"])
    # styling
    for col in "ABCDEFGHIJ": ws[f"{col}{r}"].border=b_all
    ws[f"R{r}"].border=b_all; ws[f"R{r}"].fill=PatternFill("solid",fgColor=REFBG)
    kc=ws[f"K{r}"]; kc.fill=PatternFill("solid",fgColor=INPUTBG); kc.border=b_box; kc.font=Font(bold=True); kc.alignment=Alignment(horizontal="center")
    for col in "LMNOPQ": ws[f"{col}{r}"].fill=PatternFill("solid",fgColor=OUTBG); ws[f"{col}{r}"].border=b_all
    ws[f"B{r}"].font=Font(bold=True); ws[f"A{r}"].alignment=Alignment(horizontal="center")
    ws[f"C{r}"].font=Font(size=9); ws[f"C{r}"].alignment=Alignment(wrap_text=True,vertical="center")
    for col in "DEFGHIJ": ws[f"{col}{r}"].alignment=Alignment(horizontal="center")
    if i%2==1:
        for col in "ABCDEFGHIJ":
            if ws[f"{col}{r}"].fill.fgColor.rgb in (None,"00000000"):
                ws[f"{col}{r}"].fill=PatternFill("solid",fgColor="FAFAFA")

last=first+len(P)-1
tr=last+1

# ---- Totals ----
ws.merge_cells(f"A{tr}:J{tr}")
a=ws[f"A{tr}"]; a.value="TOPLAM / GRAND TOTAL"; a.font=Font(bold=True,color=WHITE,size=11)
a.fill=PatternFill("solid",fgColor=BRAND); a.alignment=Alignment(horizontal="right",vertical="center")
ws[f"K{tr}"]=f"=SUM(K{first}:K{last})"; ws[f"L{tr}"]=f"=SUM(L{first}:L{last})"
ws[f"M{tr}"]=f"=SUM(M{first}:M{last})"; ws[f"O{tr}"]=f"=SUM(O{first}:O{last})"
ws[f"P{tr}"]=f"=SUM(P{first}:P{last})"; ws[f"Q{tr}"]=f"=SUM(Q{first}:Q{last})"
num0(ws[f"K{tr}"]); num0(ws[f"L{tr}"]); pal(ws[f"M{tr}"]); wt(ws[f"O{tr}"]); vol(ws[f"P{tr}"]); money(ws[f"Q{tr}"])
for col in "KLMNOPQ":
    c=ws[f"{col}{tr}"]; c.font=Font(bold=True); c.fill=PatternFill("solid",fgColor="D9E1F2"); c.border=b_box
ws.row_dimensions[tr].height=20

# ---- Fill KPI band now that tr is known ----
def kpi(anchor,label,formula,kind):
    lc=ws[anchor]; 
    # label cell spans 2, value next
    r=int(anchor[1:]); col=anchor[0]
    ws.merge_cells(f"{col}{r}:{chr(ord(col)+1)}{r}")
    lc.value=label; lc.font=Font(bold=True,size=9,color=BRAND); lc.fill=PatternFill("solid",fgColor=KPIBG)
    lc.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); lc.border=b_all
    vcol=chr(ord(col)+2); vc=ws[f"{vcol}{r}"]
    ws.merge_cells(f"{vcol}{r}:{chr(ord(col)+3)}{r}")
    vc.value=formula.format(tr=tr); vc.font=Font(bold=True,size=12,color="C00000")
    vc.alignment=Alignment(horizontal="center",vertical="center"); vc.fill=PatternFill("solid",fgColor=WHITE); vc.border=b_box
    if kind=="vol":vol(vc)
    elif kind=="pct":pct(vc)
    elif kind=="int":vc.number_format='0'
    elif kind=="money":money(vc)
kpi("K1","TOPLAM\nHACIM m3","=P{tr}","vol")
kpi("K3","GEREKEN\nTIR ADEDI","=IF($G$3=0,0,CEILING(P{tr}/$G$3,1))","int")
kpi("O1","TIR\nDOLULUK %","=IF($G$3=0,0,P{tr}/$G$3)","pct")
kpi("O3","TOPLAM\nTUTAR USD","=Q{tr}","money")
ws.row_dimensions[1].height=26; ws.row_dimensions[3].height=26

# ---- note ----
nr=tr+1
ws.merge_cells(f"A{nr}:R{nr}")
ws[f"A{nr}"]=("*** SIPARIS (KOLI) sutununa yalnizca KOLI/SIRA (G sutunu) degerinin KATLARI girilebilir "
 "(1 palet sirasi = 1 tam sira; ornek Diamond Fork 12/sira -> 12,24,36,48...). Yanlis katta Excel uyari verir. | "
 "PALET (esdeger) yarim paletleri de kapsar; gercek fiziksel palet ve TIR dolulugu icin 'PALET & TIR OZETI' sayfasina bakin. | "
 "NAVIPACK standart TIR = 90 m3 (dolum hacme gore hesaplanir).")
ws[f"A{nr}"].font=Font(italic=True,size=9,color="C00000"); ws[f"A{nr}"].alignment=Alignment(wrap_text=True,vertical="center")
ws.row_dimensions[nr].height=48

# ---- validation ----
dv=DataValidation(type="custom",formula1=f"=AND(K{first}>=0,MOD(K{first},G{first})=0)",
    allow_blank=True,showInputMessage=True,showErrorMessage=True)
dv.promptTitle="SIPARIS (KOLI)"; dv.prompt=("KOLI cinsinden girin. KOLI/SIRA (G) sutununun KATI olmalidir; "
 "boylece her palet sirasi tek urunle dolar ve yarim paletler duzgun birlesir.\nOrnek Diamond Fork (12/sira): 12,24,36,48,60...")
dv.errorTitle="GECERSIZ ADET"; dv.error=("Girdiginiz sayi KOLI/SIRA (G sutunu) degerinin tam kati degil. "
 "Lutfen tam sira olacak sekilde girin (or. 12'nin katlari).")
ws.add_data_validation(dv); dv.sqref=f"K{first}:K{last}"
kc=Comment("Siparisi KOLI cinsinden girin. Her hucre sadece KOLI/SIRA katlarini kabul eder.","NAVIPACK"); kc.width=260;kc.height=90
ws[f"K{HR}"].comment=kc

ws.freeze_panes=f"A{first}"

# ============================================================ SHEET 2 : PALLET & TRUCK
ps=wb.create_sheet("PALET & TIR OZETI"); ps.sheet_view.showGridLines=False
ps.merge_cells("A1:G2"); ps["A1"]="PALET & TIR OZETI / PALLET & TRUCK CONSOLIDATION"
ps["A1"].font=Font(bold=True,color=WHITE,size=13); ps["A1"].fill=PatternFill("solid",fgColor=BRAND)
ps["A1"].alignment=Alignment(horizontal="left",vertical="center")
ps.merge_cells("A3:G4")
ps["A3"]=("Yarim paletler ONCELIKLE ayni urun ailesi icinde birlestirilir (Diamond'lar Diamond'larla, Premium'lar "
 "Premium'larla, kasiklar kasiklarla). Cok az sayidaki kalanlar 'palet ustu dokme' olarak mevcut paletlerin "
 "ustune yuklenebilir. 'GERCEK PALET' = her ailenin palet-esdegerinin yukari yuvarlanmisi.")
ps["A3"].font=Font(italic=True,size=9,color="404040"); ps["A3"].alignment=Alignment(wrap_text=True,vertical="center")

oref="'SIPARIS - ORDER'"
ph=6
phdr=["AILE / FAMILY","SIPARIS KOLI","PALET ESDEGER","TAM PALET","GERCEK PALET\n(yuvarlanmis)","NET AGIRLIK kg","HACIM m3"]
pw=[22,13,14,11,15,14,12]
for j,(txt,w) in enumerate(zip(phdr,pw)):
    col=get_column_letter(j+1); ps.column_dimensions[col].width=w
    H(ps[f"{col}{ph}"],fill=HEADER)
ps.row_dimensions[ph].height=30
fams=["DIAMOND","PREMIUM","SMART","ICE CREAM","GLASS","PIZZA"]
flabel={"DIAMOND":"DIAMOND CUTLERY","PREMIUM":"PREMIUM CUTLERY","SMART":"SMART FORK","ICE CREAM":"ICE CREAM SPOON","GLASS":"GLASSES / CUPS","PIZZA":"PIZZA TRIPOD"}
p0=ph+1
for k,fam in enumerate(fams):
    r=p0+k
    ps[f"A{r}"]=flabel[fam]; ps[f"A{r}"].font=Font(bold=True)
    ps[f"I{r}"]=fam; ps.column_dimensions["I"].hidden=True
    ps[f"B{r}"]=f"=SUMIF({oref}!$R${first}:$R${last},$I{r},{oref}!$K${first}:$K${last})"
    ps[f"C{r}"]=f"=SUMIF({oref}!$R${first}:$R${last},$I{r},{oref}!$M${first}:$M${last})"
    ps[f"D{r}"]=f"=INT(C{r})"
    ps[f"E{r}"]=f"=IF(C{r}=0,0,CEILING(C{r},1))"
    ps[f"F{r}"]=f"=SUMIF({oref}!$R${first}:$R${last},$I{r},{oref}!$O${first}:$O${last})"
    ps[f"G{r}"]=f"=SUMIF({oref}!$R${first}:$R${last},$I{r},{oref}!$P${first}:$P${last})"
    num0(ps[f"B{r}"]); pal(ps[f"C{r}"]); ps[f"D{r}"].number_format='0'; ps[f"E{r}"].number_format='0'; wt(ps[f"F{r}"]); vol(ps[f"G{r}"])
    for col in "ABCDEFG":
        cc=ps[f"{col}{r}"]; cc.border=b_all
        if col=="E": cc.fill=PatternFill("solid",fgColor=OUTBG); cc.font=Font(bold=True)
pt=p0+len(fams)
ps[f"A{pt}"]="TOPLAM / GRAND TOTAL"; ps[f"A{pt}"].font=Font(bold=True,color=WHITE)
ps[f"A{pt}"].fill=PatternFill("solid",fgColor=BRAND); ps[f"A{pt}"].alignment=Alignment(horizontal="right")
for col in "BCDEFG":
    ps[f"{col}{pt}"]=f"=SUM({col}{p0}:{col}{pt-1})"
num0(ps[f"B{pt}"]); pal(ps[f"C{pt}"]); ps[f"D{pt}"].number_format='0'; ps[f"E{pt}"].number_format='0'; wt(ps[f"F{pt}"]); vol(ps[f"G{pt}"])
for col in "BCDEFG":
    c=ps[f"{col}{pt}"]; c.font=Font(bold=True); c.fill=PatternFill("solid",fgColor="D9E1F2"); c.border=b_box

# ---- TRUCK block ----
tb=pt+2
ps.merge_cells(f"A{tb}:G{tb}"); ps[f"A{tb}"]="TIR / TRUCK YUKLEME (hacme gore / by volume)"
ps[f"A{tb}"].font=Font(bold=True,color=WHITE,size=11); ps[f"A{tb}"].fill=PatternFill("solid",fgColor="833C00")
ps[f"A{tb}"].alignment=Alignment(vertical="center"); ps.row_dimensions[tb].height=20
def trow(r,label,formula,kind,hl=False):
    ps.merge_cells(f"A{r}:D{r}"); ps[f"A{r}"]=label; ps[f"A{r}"].font=Font(bold=True,size=10)
    ps[f"A{r}"].alignment=Alignment(horizontal="right",vertical="center")
    ps.merge_cells(f"E{r}:G{r}"); c=ps[f"E{r}"]; c.value=formula; c.alignment=Alignment(horizontal="center",vertical="center")
    c.font=Font(bold=True,size=12,color="C00000" if hl else "000000"); c.border=b_box
    if hl: c.fill=PatternFill("solid",fgColor="FCE4D6")
    if kind=="vol":vol(c)
    elif kind=="int":c.number_format='0'
    elif kind=="pct":pct(c)
    elif kind=="m":money(c)
    for col in "ABCD": ps[f"{col}{r}"].border=b_all
trow(tb+1,"TIR KAPASITESI (m3):",f"='SIPARIS - ORDER'!$G$3","vol")
trow(tb+2,"TOPLAM SIPARIS HACMI (m3):",f"=G{pt}","vol")
trow(tb+3,"TIR DOLULUK ORANI:",f"=IF(E{tb+1}=0,0,G{pt}/E{tb+1})","pct")
trow(tb+4,"GEREKEN TIR ADEDI:",f"=IF(E{tb+1}=0,0,CEILING(G{pt}/E{tb+1},1))","int",hl=True)
trow(tb+5,"SON TIR'DA BOS HACIM (m3):",f"=IF(E{tb+1}=0,0,E{tb+4}*E{tb+1}-G{pt})","vol")

note=tb+7
ps.merge_cells(f"A{note}:G{note+3}")
ps[f"A{note}"]=("NOT: TIR dolulugu HACME gore hesaplanir (NAVIPACK standart TIR = 90 m3; cift dorse = 120 m3 - "
 "ust taraftaki kapasiteyi degistirebilirsiniz). Farkli aileler yalnizca koli yukseklikleri uyumluysa ayni palette "
 "birlestirilebilir; bu durumda palet sayisi daha da dusebilir. Nihai istifleme fabrikada fiziksel olarak teyit edilir. "
 "Her palet sirasi TEK urunle dolmalidir.")
ps[f"A{note}"].font=Font(italic=True,size=9,color="404040"); ps[f"A{note}"].alignment=Alignment(wrap_text=True,vertical="top")

# ============================================================ SHEET 3 : NOTES
nt=wb.create_sheet("ACIKLAMA - NOTES"); nt.sheet_view.showGridLines=False
nt.column_dimensions["A"].width=3; nt.column_dimensions["B"].width=112
lines=[
 ("NAVIPACK - SIPARIS & HACIM HESAPLAMA / ORDER & VOLUME CALCULATOR (v2)","title"),
 ("","" ),
 ("NASIL KULLANILIR / HOW TO USE","h"),
 ("1) 'SIPARIS - ORDER' sayfasinda SADECE sari 'SIPARIS (KOLI)' sutununu (K) doldurun.","n"),
 ("   Ustten Musteri, Tarih ve TIR Kapasitesini (varsayilan 90 m3) girin.","n"),
 ("2) Ust bantta anlik KPI'lar: Toplam Hacim, TIR Doluluk %, Gereken TIR adedi, Toplam Tutar.","n"),
 ("3) 'PALET & TIR OZETI' sayfasi paletleri aile bazinda birlestirir ve TIR dolulugunu gosterir.","n"),
 ("","" ),
 ("NEDEN KOLI/SIRA'NIN KATLARI? / WHY MULTIPLES OF BOX-PER-LAYER?","h"),
 ("Musteri her zaman paletli (100x120 cm) yukleme yapar. Bir palet sirasina FARKLI urun dizilirse koli","n"),
 ("yukseklikleri farkli olur ve palet dengesiz yukselir. Bu yuzden her sira TEK urunle DOLU olmalidir;","n"),
 ("siparis daima tam sira (KOLI/SIRA'nin kati) girilir. Boylece kalan yarim paletler duzgun birlesir.","n"),
 ("Ornek: Diamond Fork 1 sirada 12 koli -> 12, 24, 36, 48 ... (45 yerine 48).","n"),
 ("","" ),
 ("ADET -> KOLI / UNITS -> BOXES","h"),
 ("Siparis cogu zaman ADET konusulur. 1 kolideki adet 'ADET/KOLI' sutunundadir (or. Diamond 2.000).","n"),
 ("600.000 adet Diamond Fork / 2.000 = 300 koli. 1 palet = 120 koli -> 300 koli = 2 tam palet + yarim palet.","n"),
 ("","" ),
 ("YARIM PALET BIRLESTIRME & PALET USTU DOKME","h"),
 ("Oncelik ayni aile: Diamond'lar Diamond'larla, Premium'lar Premium'larla, kasiklar kasiklarla.","n"),
 ("Cok az sayidaki kalan koliler (or. birkac koli Wavy/Shot Glass) 'palet ustu dokme' olarak mevcut","n"),
 ("paletlerin ustune yerlestirilebilir. 'PALET & TIR OZETI' bu birlestirmeyi ozetler.","n"),
 ("","" ),
 ("TIR / TRUCK","h"),
 ("NAVIPACK standart TIR = 90 m3 (dolum HACME gore). Cift dorse TIR = 120 m3.","n"),
 ("Gereken TIR = Toplam Hacim / Kapasite (yukari yuvarlanir). FERPROM gibi tamamen dokme yukleyen","n"),
 ("firmalarda palet yoktur; sadece toplam hacim TIR'a sigacak sekilde hesaplanir.","n"),
 ("","" ),
 ("HESAPLAMALAR / FORMULAS","h"),
 ("SIPARIS (ADET)  = Siparis Koli x Adet/Koli","n"),
 ("PALET (esdeger) = Siparis Koli / (Koli/Palet)     [Koli/Palet = Koli/Sira x Sira/Palet]","n"),
 ("TAM SIRA        = Siparis Koli / (Koli/Sira)","n"),
 ("NET AGIRLIK kg  = Siparis Koli x (Adet/Koli x urun agirligi gr / 1000)","n"),
 ("HACIM m3        = Siparis Koli x Koli Hacmi m3","n"),
 ("TUTAR USD       = (Siparis Adet / 1000) x Birim Fiyat","n"),
 ("","" ),
 ("RENK KODLARI / LEGEND","h"),
 ("Sari = veri girisi (siz doldurun) | Yesil = otomatik sonuc | Gri = referans (aile) | Mavi bant = KPI.","n"),
 ("","" ),
 ("KAYNAKLAR / SOURCES","h"),
 ("Urun listesi, fiyatlar ve palet dizilimleri: NAVIPACK Price Offer 20.06.2026.","n"),
 ("Palet dogrulama & palet ustu dokme: NAVIPACK Volume Calculation 07.07.2026.","n"),
 ("TIR kapasiteleri: NAVIPACK standart 90 m3, FERPROM cift dorse 120 m3.","n"),
]
r=1
for txt,kind in lines:
    c=nt[f"B{r}"]; c.value=txt
    if kind=="title": c.font=Font(bold=True,color=WHITE,size=13); c.fill=PatternFill("solid",fgColor=BRAND); nt.row_dimensions[r].height=24; c.alignment=Alignment(vertical="center")
    elif kind=="h": c.font=Font(bold=True,color=BRAND,size=11); c.fill=PatternFill("solid",fgColor="D9E1F2")
    else: c.font=Font(size=10)
    r+=1

wb.active=wb.sheetnames.index("SIPARIS - ORDER")
out="NAVIPACK_Siparis_Hacim_Hesaplama.xlsx"
wb.save(out); print("SAVED",out,"rows",first,"-",last,"total row",tr)
