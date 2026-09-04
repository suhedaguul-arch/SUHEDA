' ============================================================
'  NAVIPACK - ORDER OF BOX otomatik yuvarlama
'  ORDER OF BOX (R sutunu) girilen sayiyi, ayni satirdaki
'  BOX QTY PER LAYER (L sutunu) degerinin katina YUKARI yuvarlar.
'  Ornek: L=7 iken 1..6 -> 7 ; 8..13 -> 14 ; 0 -> 0
'
'  KURULUM (Mac Excel):
'   1) Excel > Tercihler > Serit ve Arac Cubugu > "Gelistirici" isaretle > Kaydet
'   2) Gelistirici sekmesi > Visual Basic  (kisayol: Option+F11)
'   3) Sol agacta bu kitabin "Sayfa1" (worksheet) ogesine CIFT TIKLA
'   4) Asagidaki kodun TAMAMINI acilan pencereye yapistir
'   5) Kapat. Dosyayi "Excel Makro Etkin Calisma Kitabi (.xlsm)" olarak kaydet
'   6) Dosyayi acarken "Makrolari Etkinlestir" de
'
'  NOT: Makro kullanacaksan, ayni hucrelere veri dogrulama (reddetme)
'       KOYMA - dogrulama 5'i reddedip makronun yuvarlamasini engeller.
' ============================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rng As Range, c As Range
    Dim layer As Variant, v As Variant

    ' ORDER OF BOX = R sutunu (18) ; BOX QTY PER LAYER = L sutunu (12)
    Set rng = Intersect(Target, Me.Columns(18))
    If rng Is Nothing Then Exit Sub

    Application.EnableEvents = False
    On Error Resume Next
    For Each c In rng.Cells
        layer = Me.Cells(c.Row, 12).Value
        v = c.Value
        If IsNumeric(layer) And IsNumeric(v) Then
            If layer > 0 And v <> "" Then
                If v <= 0 Then
                    c.Value = 0
                Else
                    c.Value = Application.WorksheetFunction.Ceiling(v, layer)
                End If
            End If
        End If
    Next c
    Application.EnableEvents = True
End Sub
