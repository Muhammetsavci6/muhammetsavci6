Attribute VB_Name = "ZCO002_FS10N"
Option Explicit
'--------------------------------------------------------------------------
' SAP ZCO002 FS10N RAPORU - v2 (DUZELTILMIS)
'
' v1'e gore degisiklikler:
'  - Hedef sayfa sabit: "ZCO002" (yoksa olusturulur). Aktif sayfa silinmez.
'  - Hata akisi: Resume Kapanis + hataOldu bayragi. Hata sonrasi
'    "Islem Tamamlandi" mesaji artik cikmaz; hata mesajinda adim yazar.
'  - S_BWKEY-HIGH bos birakilir (mukerrer plant verisi gelmez).
'  - Indirme basarisiz olursa SAP ekrani sifirlanir, sonraki plant'e gecilir.
'  - SAP dosyayi Excel'de kendisi actiysa once o kopya kapatilir.
'  - Tek satirlik veride cokme giderildi (AralikDizi).
'  - Lookup anahtari: Val() yerine bastaki sifirlari silen AnahtarYap()
'    (harfli kodlar artik "0" anahtarinda cakismaz).
'  - Anaveri acilamazsa kullanici uyarilir.
'  - Pivotta "EB" yoksa uyari verilir, pivot filtresiz olusturulur.
'  - FS10N: sabit Sleep yerine SAP Busy bekleme + tekrar deneme,
'    bakiye SapSayi() ile gercek sayiya cevrilir (1.234,56- / -1.234,56 /
'    1,234.56 bicimlerinin hepsi taninir).
'  - DosyaHazirMi: CPU'yu %100 yakan bos dongu ve gece yarisi hatasi giderildi.
'  - Hesaplama modu calisma oncesindeki degerine geri doner.
'--------------------------------------------------------------------------

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const HEDEF_SAYFA As String = "ZCO002"
Private Const ANAVERI_YOLU As String = "\\10.10.1.22\Maliyet_Butce_Raporlama\99_TKG Otomotiv Dokumantasyon\01_TKG Otomotiv Anaveri.xlsx"
Private Const LAYOUT_GRID As String = "wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell"
Private Const BKLAS_TABLO As String = "usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I"
Private Const FS10N_GRID As String = "wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell"

Sub ZCO002_FS10N_Final_DirektAktarim()
    Dim SapGuiAuto As Object, SAPApp As Object, connection As Object, session As Object
    Dim wsGirdi As Worksheet, wsHedef As Worksheet, wsPivot As Worksheet
    Dim wbTemp As Workbook, wsTemp As Worksheet
    Dim wbAnaveri As Workbook, wsAna As Worksheet, wbAcik As Workbook
    Dim i As Long, j As Long, k As Long, deneme As Long, nokta As Long
    Dim sonSatir As Long, tempSonSatir As Long, sonSutun As Long
    Dim hedefSonSutun As Long, anaSonSatir As Long, bayrakSutun As Long
    Dim downloadPath As String, dosyaAdi As String, tamYol As String, aktifPlant As String
    Dim raporKlasoru As String, kayitDosyaAdi As String, kayitTamYol As String
    Dim anaKey As String, dKey As String, adim As String
    Dim donem As Long, bakiyeRaw As Variant
    Dim hesaplar As Variant, hedefler As Variant, veriArray As Variant
    Dim okA As Variant, okB As Variant, isaret() As String
    Dim anaKeyArray As Variant, anaVeriArray As Variant
    Dim hedefD As Variant, hedefT As Variant, hedefE() As Variant
    Dim dict As Object
    Dim pc As PivotCache, pt As PivotTable, pTable As PivotTable
    Dim ilkPlant As Boolean, silVar As Boolean, hataOldu As Boolean, okundu As Boolean
    Dim eskiHesaplama As XlCalculation
    Dim tempDosyalar As New Collection, yol As Variant

    eskiHesaplama = Application.Calculation
    On Error GoTo HataYakala

    '--- 1. AYARLAR VE HAZIRLIK ---
    adim = "Hazirlik"
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    Set wsGirdi = ThisWorkbook.Sheets("Girdiler")
    Set wsPivot = ThisWorkbook.Sheets("Pivot")
    Set wsHedef = ThisWorkbook.Sheets(HEDEF_SAYFA)
    On Error GoTo HataYakala

    If wsGirdi Is Nothing Or wsPivot Is Nothing Then
        MsgBox "'Girdiler' veya 'Pivot' sayfasi eksik!", vbCritical
        hataOldu = True: GoTo Kapanis
    End If

    If wsHedef Is Nothing Then
        Set wsHedef = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsHedef.Name = HEDEF_SAYFA
    End If

    downloadPath = Environ("USERPROFILE") & "\Downloads\"
    wsHedef.Cells.Clear

    '--- 2. SAP BAGLANTISI ---
    adim = "SAP baglantisi"
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set connection = SAPApp.Children(0)
    Set session = connection.Children(0)
    On Error GoTo HataYakala

    If session Is Nothing Then
        MsgBox "SAP Oturumu Baglanamadi!", vbCritical
        hataOldu = True: GoTo Kapanis
    End If

    '--- 3. ZCO002 DONGUSU ---
    ilkPlant = True

    For i = 3 To 7
        aktifPlant = Trim(wsGirdi.Cells(i, "B").Text)
        If aktifPlant = "" Then GoTo SiradakiPlant

        adim = "ZCO002 - Plant " & aktifPlant
        Application.StatusBar = "Plant " & aktifPlant & " - SAP'den veri cekiliyor..."

        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZCO002"
        session.findById("wnd[0]").sendVKey 0

        With session.findById("wnd[0]")
            .findById("usr/ctxtP_BUKRS").Text = wsGirdi.Range("B1").Text
            .findById("usr/ctxtS_BWKEY-LOW").Text = aktifPlant
            .findById("usr/ctxtS_BWKEY-HIGH").Text = ""          ' sadece bu plant
            .findById("usr/btn%_S_BKLAS_%_APP_%-VALU_PUSH").Press
        End With

        With session.findById("wnd[1]")
            .findById(BKLAS_TABLO & "[1,0]").Text = wsGirdi.Range("B10").Text
            .findById(BKLAS_TABLO & "[1,1]").Text = wsGirdi.Range("B11").Text
            .findById(BKLAS_TABLO & "[1,2]").Text = wsGirdi.Range("B12").Text
            .findById("tbar[0]/btn[8]").Press
        End With

        With session.findById("wnd[0]")
            .findById("usr/txtP_CURTP").Text = wsGirdi.Range("B14").Text
            .findById("usr/txtP_BDATJ").Text = wsGirdi.Range("B16").Text
            .findById("usr/txtP_POPER").Text = wsGirdi.Range("B17").Text
            .findById("tbar[1]/btn[8]").Press
            .findById("tbar[1]/btn[33]").Press
        End With

        session.findById(LAYOUT_GRID).setCurrentCell 2, "TEXT"
        session.findById(LAYOUT_GRID).selectedRows = "2"
        session.findById(LAYOUT_GRID).clickCurrentCell

        dosyaAdi = "ZCO002_Temp_" & aktifPlant & ".xlsx"
        tamYol = downloadPath & dosyaAdi
        tempDosyalar.Add tamYol           ' sonda SAP'nin actigi kopyalari kapatmak icin

        ' Onceki calismadan acik kalmis kopya varsa kapat, eski dosyayi sil
        DosyayiKapat dosyaAdi
        On Error Resume Next
        If Len(Dir(tamYol)) > 0 Then Kill tamYol
        On Error GoTo HataYakala

        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
        session.findById("wnd[1]/tbar[0]/btn[0]").Press
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = downloadPath
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = dosyaAdi
        session.findById("wnd[1]/tbar[0]/btn[0]").Press

        If Not DosyaHazirMi(tamYol, 15) Then
            MsgBox "Uyari: " & aktifPlant & " icin dosya indirilemedi. Diger plant'lerle devam.", vbExclamation
            SapSifirla session
            GoTo SiradakiPlant
        End If

        ' SAP dosyayi Excel'de kendisi actiysa o kopyayi kapat
        Sleep 500: DoEvents
        DosyayiKapat dosyaAdi

        '=== DOSYA AKTARIMI ===
        Set wbTemp = Nothing
        On Error Resume Next
        Set wbTemp = Workbooks.Open(tamYol, ReadOnly:=True)
        On Error GoTo HataYakala

        If wbTemp Is Nothing Then
            MsgBox "Uyari: " & aktifPlant & " dosyasi acilamadi, atlaniyor.", vbExclamation
            GoTo SiradakiPlant
        End If

        Set wsTemp = wbTemp.Sheets(1)
        tempSonSatir = wsTemp.Cells(wsTemp.Rows.Count, "B").End(xlUp).Row

        If tempSonSatir > 1 Then
            sonSutun = wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column

            With wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(tempSonSatir, sonSutun))
                .NumberFormat = "@"
                .Value = .Value
            End With

            If ilkPlant Then
                wsHedef.Range(wsHedef.Cells(1, 1), wsHedef.Cells(1, sonSutun)).EntireColumn.NumberFormat = "@"
                veriArray = AralikDizi(wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(tempSonSatir, sonSutun)))
                wsHedef.Cells(1, 1).Resize(UBound(veriArray, 1), UBound(veriArray, 2)).Value = veriArray
                ilkPlant = False
            Else
                veriArray = AralikDizi(wsTemp.Range(wsTemp.Cells(2, 1), wsTemp.Cells(tempSonSatir, sonSutun)))
                sonSatir = wsHedef.Cells(wsHedef.Rows.Count, "B").End(xlUp).Row
                wsHedef.Cells(sonSatir + 1, 1).Resize(UBound(veriArray, 1), UBound(veriArray, 2)).Value = veriArray
            End If
        End If

        wbTemp.Close False
        Set wbTemp = Nothing
        Set wsTemp = Nothing
        DoEvents

        On Error Resume Next
        Kill tamYol
        On Error GoTo HataYakala

SiradakiPlant:
    Next i

    '--- 4. EXCEL TEMIZLIK (AUTOFILTER ILE TOPLU SILME) ---
    adim = "Temizlik"
    Application.StatusBar = "Veriler temizleniyor..."
    wsHedef.Activate
    On Error Resume Next
    wsHedef.Outline.ShowLevels RowLevels:=8
    wsHedef.Cells.ClearOutline
    If wsHedef.AutoFilterMode Then wsHedef.AutoFilterMode = False
    On Error GoTo HataYakala

    sonSatir = SonSatirBul(wsHedef)

    If sonSatir >= 2 Then
        bayrakSutun = wsHedef.Cells(1, wsHedef.Columns.Count).End(xlToLeft).Column + 1

        okA = AralikDizi(wsHedef.Range(wsHedef.Cells(2, 1), wsHedef.Cells(sonSatir, 1)))
        okB = AralikDizi(wsHedef.Range(wsHedef.Cells(2, 2), wsHedef.Cells(sonSatir, 2)))
        ReDim isaret(1 To sonSatir - 1, 1 To 1)

        silVar = False
        For i = 1 To sonSatir - 1
            If (Metin(okA(i, 1)) = "" And Metin(okB(i, 1)) = "") Or _
               InStr(1, Metin(okA(i, 1)), "Toplam") > 0 Then
                isaret(i, 1) = "SIL"
                silVar = True
            Else
                isaret(i, 1) = "KAL"
            End If
        Next i

        If silVar Then
            wsHedef.Cells(1, bayrakSutun).Value = "BAYRAK"
            wsHedef.Range(wsHedef.Cells(2, bayrakSutun), wsHedef.Cells(sonSatir, bayrakSutun)).Value = isaret

            With wsHedef.Range(wsHedef.Cells(1, 1), wsHedef.Cells(sonSatir, bayrakSutun))
                .AutoFilter Field:=bayrakSutun, Criteria1:="SIL"
                On Error Resume Next
                .Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
                On Error GoTo HataYakala
            End With
            wsHedef.AutoFilterMode = False
            wsHedef.Columns(bayrakSutun).ClearContents
        End If
    End If

    '--- 5. MAL GRUBU TANIMI (ANAVERI -> DICTIONARY) ---
    adim = "Anaveri lookup"
    sonSatir = SonSatirBul(wsHedef)
    If sonSatir < 2 Then
        MsgBox "SAP'den hic veri gelmedi, rapor olusturulamadi.", vbExclamation
        hataOldu = True: GoTo Kapanis
    End If

    wsHedef.Columns("E:E").Insert Shift:=xlToRight
    wsHedef.Range("E1").Value = "Mal grubu Tanimi"

    Set dict = CreateObject("Scripting.Dictionary")

    Application.StatusBar = "Anaveri okunuyor..."
    Set wbAnaveri = Nothing
    On Error Resume Next
    Set wbAnaveri = Workbooks.Open(ANAVERI_YOLU, UpdateLinks:=0, ReadOnly:=True)
    On Error GoTo HataYakala

    If wbAnaveri Is Nothing Then
        MsgBox "Uyari: Anaveri dosyasi acilamadi, 'Mal grubu Tanimi' sutunu bos kalacak." & _
               vbCrLf & vbCrLf & ANAVERI_YOLU, vbExclamation
    Else
        Set wsAna = wbAnaveri.Sheets("4")
        anaSonSatir = wsAna.Cells(wsAna.Rows.Count, 1).End(xlUp).Row
        If anaSonSatir >= 5 Then
            anaKeyArray = AralikDizi(wsAna.Range("A5:A" & anaSonSatir))
            anaVeriArray = AralikDizi(wsAna.Range("C5:C" & anaSonSatir))
            For k = 1 To UBound(anaKeyArray, 1)
                anaKey = AnahtarYap(anaKeyArray(k, 1))
                If anaKey <> "" Then
                    If Not dict.Exists(anaKey) Then dict(anaKey) = anaVeriArray(k, 1)
                End If
            Next k
        End If
        wbAnaveri.Close False
        Set wbAnaveri = Nothing
    End If

    ThisWorkbook.Activate
    hedefD = AralikDizi(wsHedef.Range("D2:D" & sonSatir))
    hedefT = AralikDizi(wsHedef.Range("T2:T" & sonSatir))
    ReDim hedefE(1 To sonSatir - 1, 1 To 1)

    For i = 1 To sonSatir - 1
        dKey = AnahtarYap(hedefD(i, 1))
        Select Case Metin(hedefT(i, 1))
            Case "3000": hedefE(i, 1) = "Mamul"
            Case "2000": hedefE(i, 1) = "Yari Mamul"
            Case Else
                If dict.Exists(dKey) Then hedefE(i, 1) = dict(dKey) Else hedefE(i, 1) = ""
        End Select
    Next i

    wsHedef.Range("E2:E" & sonSatir).Value = hedefE

    '--- 6. DINAMIK PIVOT ---
    adim = "Pivot"
    Application.StatusBar = "Pivot olusturuluyor..."

    For Each pTable In wsPivot.PivotTables
        pTable.TableRange2.Clear
    Next pTable

    wsPivot.Columns("A:A").ClearContents
    wsPivot.Columns("F:Z").ClearContents

    hedefSonSutun = wsHedef.Cells(1, wsHedef.Columns.Count).End(xlToLeft).Column
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
             SourceData:=wsHedef.Range(wsHedef.Cells(1, 1), wsHedef.Cells(sonSatir, hedefSonSutun)))
    Set pt = pc.CreatePivotTable(TableDestination:=wsPivot.Range("A1"), TableName:="ZCO002_Ozet")

    With pt
        .PivotFields(wsHedef.Range("E1").Value).Orientation = xlRowField
        .PivotFields(wsHedef.Range("C1").Value).Orientation = xlColumnField
        With .PivotFields(wsHedef.Range("AE1").Value)
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
            .Caption = " Toplam Tutar"
        End With
        With .PivotFields(wsHedef.Range("B1").Value)
            .Orientation = xlPageField
            On Error Resume Next
            .CurrentPage = "EB"
            If Err.Number <> 0 Then
                Err.Clear
                MsgBox "Uyari: Pivot filtresinde 'EB' bulunamadi, pivot filtresiz olusturuldu.", vbExclamation
            End If
            On Error GoTo HataYakala
        End With
    End With

    '--- 7. FS10N CANLI BAKIYE ---
    adim = "FS10N"
    donem = Val(wsGirdi.Range("B17").Value)
    hesaplar = Array(wsPivot.Range("B19").Value, wsPivot.Range("B20").Value, _
                     wsPivot.Range("B21").Value, wsPivot.Range("B22").Value)
    hedefler = Array("D19", "D20", "C21", "C22")

    For j = 0 To UBound(hesaplar)
        If Metin(hesaplar(j)) <> "" Then
            adim = "FS10N - Hesap " & Metin(hesaplar(j))
            Application.StatusBar = "FS10N bakiye kontrol: " & Metin(hesaplar(j))

            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nfs10n"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/ctxtRACCT-LOW").Text = Metin(hesaplar(j))
            session.findById("wnd[0]/usr/ctxtRBUKRS-LOW").Text = wsGirdi.Range("B1").Text
            session.findById("wnd[0]/usr/txtRYEAR").Text = wsGirdi.Range("B16").Text
            session.findById("wnd[0]/tbar[1]/btn[8]").Press

            SapBekle session, 30

            ' Grid hazir olana kadar en cok ~5 sn tekrar dene
            okundu = False
            bakiyeRaw = Empty
            On Error Resume Next
            For deneme = 1 To 10
                Err.Clear
                bakiyeRaw = session.findById(FS10N_GRID).getCellValue(donem, "BALANCE_CUM")
                If Err.Number = 0 Then okundu = True: Exit For
                Sleep 500
            Next deneme
            Err.Clear
            On Error GoTo HataYakala

            If okundu Then
                wsPivot.Range(hedefler(j)).Value = SapSayi(bakiyeRaw)
            Else
                MsgBox "Uyari: " & Metin(hesaplar(j)) & " hesabi icin FS10N bakiyesi okunamadi.", vbExclamation
            End If

            session.findById("wnd[0]").sendVKey 12
        End If
    Next j

    '--- 8. XLSX KAYDET ---
    adim = "Kaydetme"
    If ThisWorkbook.Path = "" Then
        MsgBox "Bu dosya henuz diske kaydedilmemis; rapor kopyasi olusturulamadi.", vbExclamation
        hataOldu = True: GoTo Kapanis
    End If

    raporKlasoru = ThisWorkbook.Path & "\Raporlar\"
    If Len(Dir(raporKlasoru, vbDirectory)) = 0 Then MkDir raporKlasoru

    nokta = InStrRev(ThisWorkbook.Name, ".")
    If nokta > 0 Then
        kayitDosyaAdi = Left(ThisWorkbook.Name, nokta - 1) & ".xlsx"
    Else
        kayitDosyaAdi = ThisWorkbook.Name & ".xlsx"
    End If
    kayitTamYol = raporKlasoru & kayitDosyaAdi

    ' Ayni isimli rapor Excel'de aciksa SaveAs hata verir -> kullaniciya soyle
    Set wbAcik = Nothing
    On Error Resume Next
    Set wbAcik = Workbooks(kayitDosyaAdi)
    On Error GoTo HataYakala
    If Not wbAcik Is Nothing Then
        MsgBox "'" & kayitDosyaAdi & "' su an acik. Kapatip raporu tekrar kaydedin." & vbCrLf & _
               "(Veriler bu dosyada hazir, sadece kopya kaydedilmedi.)", vbExclamation
        hataOldu = True: GoTo Kapanis
    End If

    On Error Resume Next
    If Len(Dir(kayitTamYol)) > 0 Then Kill kayitTamYol
    On Error GoTo HataYakala

    ThisWorkbook.Sheets(Array(wsHedef.Name, wsPivot.Name)).Copy
    With ActiveWorkbook
        .SaveAs Filename:=kayitTamYol, FileFormat:=xlOpenXMLWorkbook
        .Close False
    End With

    GoTo Kapanis

HataYakala:
    hataOldu = True
    Application.StatusBar = False
    MsgBox "HATA!" & vbCrLf & vbCrLf & _
           "Adim: " & adim & vbCrLf & _
           "Aciklama: " & Err.Description & vbCrLf & _
           "Hata No: " & Err.Number, vbCritical, "Hata"
    Resume Kapanis

Kapanis:
    '--- 9. KAPANIS ---
    On Error Resume Next
    If Not wbTemp Is Nothing Then wbTemp.Close False
    If Not wbAnaveri Is Nothing Then wbAnaveri.Close False

    ' SAP'nin kendi actigi (cogu zaman ayri Excel penceresi) temp dosyalari kapat ve sil
    Application.StatusBar = "Gecici dosyalar kapatiliyor..."
    For Each yol In tempDosyalar
        SapExcelKapat CStr(yol)
    Next yol

    Application.Calculation = eskiHesaplama
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    wsPivot.Activate
    On Error GoTo 0

    If Not hataOldu Then
        MsgBox "Islem Tamamlandi!" & vbNewLine & "Kayit: " & kayitTamYol, vbInformation
    End If
End Sub

'==========================================================================
' YARDIMCI FONKSIYONLAR
'==========================================================================

' Dosya olustu ve boyutu sabitlendi mi? (en cok MaxBekle saniye bekler)
Private Function DosyaHazirMi(ByVal DosyaYolu As String, ByVal MaxBekle As Double) As Boolean
    Dim t As Double, b1 As Long, b2 As Long
    t = Timer
    On Error Resume Next
    Do While GecenSure(t) < MaxBekle
        If Len(Dir(DosyaYolu)) > 0 Then
            Err.Clear
            b1 = FileLen(DosyaYolu)
            Sleep 800
            b2 = FileLen(DosyaYolu)
            If Err.Number = 0 And b1 = b2 And b1 > 0 Then
                DosyaHazirMi = True
                Exit Function
            End If
        Else
            Sleep 250              ' bos donguyle CPU'yu yakma
        End If
        DoEvents
    Loop
    DosyaHazirMi = False
End Function

' Timer gece yarisi sifirlanir; gecen sureyi buna dayanikli hesapla
Private Function GecenSure(ByVal baslangic As Double) As Double
    GecenSure = Timer - baslangic
    If GecenSure < 0 Then GecenSure = GecenSure + 86400
End Function

' SAP mesgulken bekle
Private Sub SapBekle(ByVal ses As Object, Optional ByVal maxSn As Double = 30)
    Dim t As Double
    t = Timer
    On Error Resume Next
    Do While ses.Busy
        Sleep 200
        DoEvents
        If GecenSure(t) > maxSn Then Exit Do
    Loop
End Sub

' Acik popup'lari kapatip SAP'yi ana ekrana dondur
Private Sub SapSifirla(ByVal ses As Object)
    Dim n As Long
    On Error Resume Next
    For n = 1 To 3
        If ses.Children.Count > 1 Then ses.ActiveWindow.sendVKey 12
    Next n
    ses.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    ses.findById("wnd[0]").sendVKey 0
End Sub

' Bu Excel'de ayni isimle acik bir dosya varsa kaydetmeden kapat
Private Sub DosyayiKapat(ByVal ad As String)
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(ad)
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

' Dosya baska bir program/Excel penceresinde acik (kilitli) mi?
Private Function DosyaKilitliMi(ByVal yol As String) As Boolean
    Dim f As Integer
    If Len(Dir(yol)) = 0 Then Exit Function
    f = FreeFile
    On Error Resume Next
    Open yol For Binary Access Read Write Lock Read Write As #f
    DosyaKilitliMi = (Err.Number <> 0)
    Close #f
End Function

' SAP'nin actigi temp dosyayi HANGI Excel penceresinde olursa olsun kapat, sonra sil.
' Workbooks("ad") sadece bu Excel'i gorur; GetObject(tamYol) diger Excel
' orneklerinde acik dosyayi da bulur.
Private Sub SapExcelKapat(ByVal yol As String)
    Dim wb As Object, app As Object, n As Long
    On Error Resume Next
    For n = 1 To 5
        If Not DosyaKilitliMi(yol) Then Exit For
        Set wb = Nothing: Set app = Nothing
        Set wb = GetObject(yol)
        If Not wb Is Nothing Then
            Set app = wb.Application
            wb.Close SaveChanges:=False
            ' Baska Excel ornegi bos kaldiysa onu da kapat (bizimkine dokunma)
            If Not app Is Nothing Then
                If app.Hwnd <> Application.Hwnd And app.Workbooks.Count = 0 Then app.Quit
            End If
        End If
        Sleep 300
        DoEvents
    Next n
    If Not DosyaKilitliMi(yol) Then
        If Len(Dir(yol)) > 0 Then Kill yol
    End If
End Sub

' Sayfadaki herhangi bir sutunda dolu olan son satir (bos sayfa -> 0)
Private Function SonSatirBul(ByVal ws As Worksheet) As Long
    Dim r As Range
    Set r = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If r Is Nothing Then SonSatirBul = 0 Else SonSatirBul = r.Row
End Function

' Range.Value tek hucrede dizi dondurmez; her zaman 2 boyutlu dizi dondur
Private Function AralikDizi(ByVal rng As Range) As Variant
    Dim a() As Variant
    If rng.Cells.CountLarge = 1 Then
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = rng.Value
        AralikDizi = a
    Else
        AralikDizi = rng.Value
    End If
End Function

' Hata/bos degerlere dayanikli Trim(CStr())
Private Function Metin(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        Metin = ""
    Else
        Metin = Trim(CStr(v))
    End If
End Function

' Lookup anahtari: bastaki sifirlari sil ("000123" = 123), harfli kodlari koru
Private Function AnahtarYap(ByVal v As Variant) As String
    Dim s As String
    s = Metin(v)
    Do While Len(s) > 1 And Left(s, 1) = "0"
        s = Mid(s, 2)
    Loop
    AnahtarYap = s
End Function

' SAP tutar metnini sayiya cevir.
' Taninan bicimler: 1.234,56-  -1.234,56  1234,56  1,234.56  1.234  (2 ondalik varsayimi)
Private Function SapSayi(ByVal v As Variant) As Variant
    Dim s As String, ayrac As String
    Dim neg As Boolean, pNokta As Long, pVirgul As Long, p As Long, adet As Long

    s = Replace(Metin(v), " ", "")
    If s = "" Then SapSayi = Empty: Exit Function

    If Right(s, 1) = "-" Then neg = True: s = Left(s, Len(s) - 1)
    If Left(s, 1) = "-" Then neg = True: s = Mid(s, 2)

    pNokta = InStrRev(s, ".")
    pVirgul = InStrRev(s, ",")

    If pNokta > 0 And pVirgul > 0 Then
        If pVirgul > pNokta Then                 ' 1.234,56
            s = Replace(s, ".", "")
            s = Replace(s, ",", ".")
        Else                                     ' 1,234.56
            s = Replace(s, ",", "")
        End If
    ElseIf pNokta > 0 Or pVirgul > 0 Then        ' tek tur ayrac
        If pVirgul > 0 Then ayrac = "," Else ayrac = "."
        p = InStrRev(s, ayrac)
        adet = Len(s) - Len(Replace(s, ayrac, ""))
        If adet > 1 Or Len(s) - p = 3 Then
            s = Replace(s, ayrac, "")            ' binlik ayraci
        Else
            s = Replace(s, ayrac, ".")           ' ondalik ayraci
        End If
    End If

    If Len(s) = 0 Or s Like "*[!0-9.]*" Then
        SapSayi = Metin(v)                       ' sayi degilse oldugu gibi birak
    Else
        SapSayi = CDbl(Val(s)) * IIf(neg, -1, 1)
    End If
End Function
