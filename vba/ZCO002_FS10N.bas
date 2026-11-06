Attribute VB_Name = "ZCO002_FS10N"
Option Explicit
'--------------------------------------------------------------------------
' SAP ZCO002 FS10N RAPORU - v3 (KORUMALI)
'
' v3 korumalari:
'  1. Baslik kontrolu: "Malzeme turu" ve "Toplam deger" sutunlari ADIYLA
'     bulunur. Yoksa makro durur; beklenen yerde (T / AE) degilse sorar.
'     Her plant'in sutun yapisi ilk plant ile ayni olmali, degilse alinmaz.
'  2. Eski dosya korumasi: indirmeden once eski temp dosya kapatilip silinir;
'     silinemezse veya inen dosya eski tarihliyse o plant ALINMAZ.
'  3. Son mesajda ozet: alinan / alinamayan plant'ler + nedenleri,
'     SAP sistem/mandant/kullanici, tum uyarilar.
'  4. "Toplam" ve "Total" satirlari silinir (TR / EN SAP).
'  5. Bir plant'te SAP hatasi / veri yok -> makro durmaz, plant atlanir.
'  + Ayni plant iki kez yazilmissa ikinci kez alinmaz.
'  + Birden fazla SAP oturumu aciksa hangisinin kullanilacagi sorulur.
'  + Girdiler (B1, B14, B16, B17) en basta kontrol edilir.
'  + Anaveride eslesmeyen mal grubu kodlari raporlanir.
'  + Pivot B19:D22'ye tasarsa FS10N yazilmaz, uyari verilir.
'  + Dosya OneDrive/SharePoint'teyse rapor Belgeler\Raporlar'a kaydedilir.
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
Private Const BEKLENEN_KOL_TUR As Long = 20     ' T  (E sutunu eklendikten sonra)
Private Const BEKLENEN_KOL_TUTAR As Long = 31   ' AE (E sutunu eklendikten sonra)

' Turkce karakterli basliklar (.bas dosyasi kod sayfasindan bagimsiz olsun diye ChrW ile)
Private Function BASLIK_TUR() As String
    BASLIK_TUR = "Malzeme t" & ChrW(252) & "r" & ChrW(252)       ' Malzeme turu
End Function
Private Function BASLIK_TUTAR() As String
    BASLIK_TUTAR = "Toplam de" & ChrW(287) & "er"                ' Toplam deger
End Function

Sub ZCO002_FS10N_Final_DirektAktarim()
    Dim SapGuiAuto As Object, SAPApp As Object, connection As Object, session As Object
    Dim wsGirdi As Worksheet, wsHedef As Worksheet, wsPivot As Worksheet
    Dim wbTemp As Workbook, wsTemp As Worksheet
    Dim wbAnaveri As Workbook, wsAna As Worksheet, wbAcik As Workbook
    Dim i As Long, j As Long, k As Long, deneme As Long, nokta As Long
    Dim sonSatir As Long, tempSonSatir As Long, sonSutun As Long
    Dim hedefSonSutun As Long, anaSonSatir As Long, bayrakSutun As Long
    Dim kolT As Long, kolAE As Long, oturumSayisi As Long, eslesmeyen As Long
    Dim downloadPath As String, dosyaAdi As String, tamYol As String, aktifPlant As String
    Dim raporKlasoru As String, kayitDosyaAdi As String, kayitTamYol As String
    Dim anaKey As String, dKey As String, adim As String, girdiHata As String
    Dim alinan As String, eksik As String, uyarilar As String, sapBilgi As String
    Dim imza As String, ilkImza As String, mesaj As String, aStr As String
    Dim donem As Long, bakiyeRaw As Variant
    Dim hesaplar As Variant, hedefler As Variant, veriArray As Variant
    Dim okA As Variant, okB As Variant, isaret() As String
    Dim anaKeyArray As Variant, anaVeriArray As Variant
    Dim hedefD As Variant, hedefT As Variant, hedefE() As Variant
    Dim dict As Object, gorulen As Object, eslesmeyenKod As Object
    Dim pc As PivotCache, pt As PivotTable, pTable As PivotTable
    Dim ilkPlant As Boolean, silVar As Boolean, hataOldu As Boolean
    Dim okundu As Boolean, fs10nYaz As Boolean
    Dim indirmeBasi As Date
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

    '--- 1b. GIRDI KONTROLU ---
    adim = "Girdi kontrolu"
    If Trim(wsGirdi.Range("B1").Text) = "" Then girdiHata = girdiHata & "- B1 (Sirket kodu) bos" & vbCrLf
    If Trim(wsGirdi.Range("B14").Text) = "" Then girdiHata = girdiHata & "- B14 (Para tipi) bos" & vbCrLf
    If Len(Trim(wsGirdi.Range("B16").Text)) <> 4 Or Not IsNumeric(wsGirdi.Range("B16").Value) Then _
        girdiHata = girdiHata & "- B16 (Yil) 4 haneli sayi olmali" & vbCrLf
    donem = Val(wsGirdi.Range("B17").Value)
    If donem < 1 Or donem > 16 Then girdiHata = girdiHata & "- B17 (Donem) 1-16 arasi olmali" & vbCrLf
    If Application.WorksheetFunction.CountA(wsGirdi.Range("B3:B7")) = 0 Then _
        girdiHata = girdiHata & "- B3:B7 arasinda hic plant yok" & vbCrLf

    If girdiHata <> "" Then
        MsgBox "Girdiler sayfasinda eksik/hatali alan var:" & vbCrLf & vbCrLf & girdiHata, vbCritical
        hataOldu = True: GoTo Kapanis
    End If

    ' FS10N hesaplarini pivot silinmeden ONCE oku
    hesaplar = Array(wsPivot.Range("B19").Value, wsPivot.Range("B20").Value, _
                     wsPivot.Range("B21").Value, wsPivot.Range("B22").Value)
    hedefler = Array("D19", "D20", "C21", "C22")

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

    On Error Resume Next
    sapBilgi = session.Info.SystemName & " / Mandant " & session.Info.Client & " / " & session.Info.User
    oturumSayisi = 0
    For k = 0 To SAPApp.Children.Count - 1
        oturumSayisi = oturumSayisi + SAPApp.Children(CLng(k)).Children.Count
    Next k
    On Error GoTo HataYakala

    If oturumSayisi > 1 Then
        If MsgBox("Birden fazla SAP oturumu acik (" & oturumSayisi & ")." & vbCrLf & vbCrLf & _
                  "Kullanilacak oturum: " & sapBilgi & vbCrLf & vbCrLf & _
                  "Bu oturumla devam edilsin mi?", vbQuestion + vbYesNo) = vbNo Then
            hataOldu = True: GoTo Kapanis
        End If
    End If

    '--- 3. ZCO002 DONGUSU (her plant kendi hata korumasiyla) ---
    Set gorulen = CreateObject("Scripting.Dictionary")
    ilkPlant = True

    For i = 3 To 7
        On Error GoTo PlantHata
        Set wbTemp = Nothing
        aktifPlant = Trim(wsGirdi.Cells(i, "B").Text)
        If aktifPlant = "" Then GoTo SiradakiPlant

        If gorulen.Exists(aktifPlant) Then
            eksik = eksik & "  " & aktifPlant & ": Girdiler'de tekrar yazilmis, ikinci kez alinmadi" & vbCrLf
            GoTo SiradakiPlant
        End If
        gorulen(aktifPlant) = True

        adim = "ZCO002 - Plant " & aktifPlant
        Application.StatusBar = "Plant " & aktifPlant & " - SAP'den veri cekiliyor..."

        dosyaAdi = "ZCO002_Temp_" & aktifPlant & ".xlsx"
        tamYol = downloadPath & dosyaAdi
        tempDosyalar.Add tamYol

        ' Onceki calismadan kalan dosyayi (hangi Excel'de acik olursa olsun) kapat ve sil
        DosyayiKapat dosyaAdi
        SapExcelKapat tamYol
        If Len(Dir(tamYol)) > 0 Then
            eksik = eksik & "  " & aktifPlant & ": eski temp dosya silinemedi (acik olabilir), plant alinmadi" & vbCrLf
            GoTo SiradakiPlant
        End If

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

        indirmeBasi = Now
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
        session.findById("wnd[1]/tbar[0]/btn[0]").Press
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = downloadPath
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = dosyaAdi
        session.findById("wnd[1]/tbar[0]/btn[0]").Press

        If Not DosyaHazirMi(tamYol, 15) Then
            eksik = eksik & "  " & aktifPlant & ": dosya 15 sn icinde inmedi" & vbCrLf
            GoTo PlantTemizle
        End If

        ' Ek guvence: inen dosya bu calismada mi olustu?
        If FileDateTime(tamYol) < indirmeBasi - TimeSerial(0, 0, 5) Then
            eksik = eksik & "  " & aktifPlant & ": dosya eski tarihli (onceki calismadan), alinmadi" & vbCrLf
            GoTo PlantTemizle
        End If

        ' SAP dosyayi bu Excel'de kendisi actiysa o kopyayi kapat
        Sleep 500: DoEvents
        DosyayiKapat dosyaAdi

        '=== DOSYA AKTARIMI ===
        On Error Resume Next
        Set wbTemp = Workbooks.Open(tamYol, ReadOnly:=True)
        On Error GoTo PlantHata

        If wbTemp Is Nothing Then
            eksik = eksik & "  " & aktifPlant & ": indirilen dosya acilamadi" & vbCrLf
            GoTo PlantTemizle
        End If

        Set wsTemp = wbTemp.Sheets(1)
        tempSonSatir = wsTemp.Cells(wsTemp.Rows.Count, "B").End(xlUp).Row

        If tempSonSatir <= 1 Then
            eksik = eksik & "  " & aktifPlant & ": SAP'de bu donem icin veri yok" & vbCrLf
            GoTo PlantTemizle
        End If

        sonSutun = wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column

        ' Her plant ayni sutun yapisinda gelmeli (layout kaymasina karsi)
        imza = BaslikImza(wsTemp, sonSutun)
        If Not ilkPlant And imza <> ilkImza Then
            eksik = eksik & "  " & aktifPlant & ": sutun yapisi ilk plant'ten farkli (layout?), alinmadi" & vbCrLf
            GoTo PlantTemizle
        End If

        With wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(tempSonSatir, sonSutun))
            .NumberFormat = "@"
            .Value = .Value
        End With

        If ilkPlant Then
            wsHedef.Range(wsHedef.Cells(1, 1), wsHedef.Cells(1, sonSutun)).EntireColumn.NumberFormat = "@"
            veriArray = AralikDizi(wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(tempSonSatir, sonSutun)))
            wsHedef.Cells(1, 1).Resize(UBound(veriArray, 1), UBound(veriArray, 2)).Value = veriArray
            ilkImza = imza
            ilkPlant = False
        Else
            veriArray = AralikDizi(wsTemp.Range(wsTemp.Cells(2, 1), wsTemp.Cells(tempSonSatir, sonSutun)))
            sonSatir = wsHedef.Cells(wsHedef.Rows.Count, "B").End(xlUp).Row
            wsHedef.Cells(sonSatir + 1, 1).Resize(UBound(veriArray, 1), UBound(veriArray, 2)).Value = veriArray
        End If

        alinan = alinan & aktifPlant & " (" & (tempSonSatir - 1) & " satir)  "

        wbTemp.Close False
        Set wbTemp = Nothing
        Set wsTemp = Nothing
        DoEvents
        On Error Resume Next
        Kill tamYol
        GoTo SiradakiPlant

PlantHata:
        eksik = eksik & "  " & aktifPlant & ": SAP/dosya hatasi - " & Err.Description & vbCrLf
        Resume PlantTemizle

PlantTemizle:
        On Error Resume Next
        If Not wbTemp Is Nothing Then wbTemp.Close False
        Set wbTemp = Nothing
        Set wsTemp = Nothing
        SapSifirla session

SiradakiPlant:
    Next i
    On Error GoTo HataYakala

    If alinan = "" Then
        MsgBox "Hicbir plant icin veri alinamadi, rapor olusturulmadi." & vbCrLf & vbCrLf & eksik, vbCritical
        hataOldu = True: GoTo Kapanis
    End If

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
            aStr = Metin(okA(i, 1))
            If (aStr = "" And Metin(okB(i, 1)) = "") Or _
               InStr(1, aStr, "Toplam") > 0 Or InStr(1, aStr, "Total") > 0 Then
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

    '--- 5. BASLIK KONTROLU + MAL GRUBU TANIMI ---
    adim = "Baslik kontrolu"
    sonSatir = SonSatirBul(wsHedef)
    If sonSatir < 2 Then
        MsgBox "Temizlikten sonra hic veri satiri kalmadi, rapor olusturulmadi.", vbCritical
        hataOldu = True: GoTo Kapanis
    End If

    wsHedef.Columns("E:E").Insert Shift:=xlToRight
    wsHedef.Range("E1").Value = "Mal grubu Tanimi"

    kolT = BaslikBul(wsHedef, BASLIK_TUR())
    kolAE = BaslikBul(wsHedef, BASLIK_TUTAR())

    If kolT = 0 Or kolAE = 0 Then
        MsgBox "SAP'den gelen basliklarda '" & BASLIK_TUR() & "' ve/veya '" & BASLIK_TUTAR() & _
               "' bulunamadi." & vbCrLf & "ZCO002'de yanlis layout secilmis olabilir." & vbCrLf & vbCrLf & _
               "Rapor olusturulmadi. Gelen veri '" & HEDEF_SAYFA & "' sayfasinda incelenebilir.", vbCritical
        hataOldu = True: GoTo Kapanis
    End If

    If kolT <> BEKLENEN_KOL_TUR Or kolAE <> BEKLENEN_KOL_TUTAR Then
        If MsgBox("Sutunlar beklenen yerde degil:" & vbCrLf & _
                  "  " & BASLIK_TUR() & ": " & SutunHarfi(kolT) & " (beklenen T)" & vbCrLf & _
                  "  " & BASLIK_TUTAR() & ": " & SutunHarfi(kolAE) & " (beklenen AE)" & vbCrLf & vbCrLf & _
                  "SAP layout'u degismis olabilir; B, C, D sutunlari da kaymis olabilir." & vbCrLf & _
                  "Yine de devam edilsin mi?", vbExclamation + vbYesNo) = vbNo Then
            hataOldu = True: GoTo Kapanis
        End If
        uyarilar = uyarilar & "- Sutunlar beklenen yerde degildi (kullanici onayiyla devam edildi)" & vbCrLf
    End If

    adim = "Anaveri lookup"
    Set dict = CreateObject("Scripting.Dictionary")
    Set eslesmeyenKod = CreateObject("Scripting.Dictionary")

    Application.StatusBar = "Anaveri okunuyor..."
    Set wbAnaveri = Nothing
    On Error Resume Next
    Set wbAnaveri = Workbooks.Open(ANAVERI_YOLU, UpdateLinks:=0, ReadOnly:=True)
    On Error GoTo HataYakala

    If wbAnaveri Is Nothing Then
        uyarilar = uyarilar & "- Anaveri dosyasi acilamadi; 'Mal grubu Tanimi' bos kaldi" & vbCrLf
    Else
        Set wsAna = Nothing
        On Error Resume Next
        Set wsAna = wbAnaveri.Sheets("4")
        On Error GoTo HataYakala

        If wsAna Is Nothing Then
            uyarilar = uyarilar & "- Anaveri dosyasinda '4' sayfasi yok; 'Mal grubu Tanimi' bos kaldi" & vbCrLf
        Else
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
        End If
        wbAnaveri.Close False
        Set wbAnaveri = Nothing
    End If

    ThisWorkbook.Activate
    hedefD = AralikDizi(wsHedef.Range("D2:D" & sonSatir))
    hedefT = AralikDizi(wsHedef.Range(wsHedef.Cells(2, kolT), wsHedef.Cells(sonSatir, kolT)))
    ReDim hedefE(1 To sonSatir - 1, 1 To 1)

    eslesmeyen = 0
    For i = 1 To sonSatir - 1
        dKey = AnahtarYap(hedefD(i, 1))
        Select Case Metin(hedefT(i, 1))
            Case "3000": hedefE(i, 1) = "Mamul"
            Case "2000": hedefE(i, 1) = "Yari Mamul"
            Case Else
                If dict.Exists(dKey) Then
                    hedefE(i, 1) = dict(dKey)
                Else
                    hedefE(i, 1) = ""
                    eslesmeyen = eslesmeyen + 1
                    If dKey = "" Then dKey = "(bos)"
                    If Not eslesmeyenKod.Exists(dKey) Then eslesmeyenKod(dKey) = True
                End If
        End Select
    Next i

    wsHedef.Range("E2:E" & sonSatir).Value = hedefE

    If eslesmeyen > 0 And dict.Count > 0 Then
        uyarilar = uyarilar & "- Anaveride bulunamayan mal grubu: " & eslesmeyen & " satir, kodlar: " & _
                   IlkNAnahtar(eslesmeyenKod, 10) & vbCrLf
    End If

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
        With .PivotFields(wsHedef.Cells(1, kolAE).Value)
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
                uyarilar = uyarilar & "- Pivot filtresinde 'EB' bulunamadi; pivot FILTRESIZ olusturuldu" & vbCrLf
            End If
            On Error GoTo HataYakala
        End With
    End With

    ' Pivot FS10N hucrelerine tastiysa oraya yazma
    fs10nYaz = (Intersect(pt.TableRange2, wsPivot.Range("B19:D22")) Is Nothing)
    If Not fs10nYaz Then
        uyarilar = uyarilar & "- Pivot B19:D22 alanina tasti; FS10N bakiyeleri YAZILMADI" & vbCrLf
    End If

    '--- 7. FS10N CANLI BAKIYE ---
    If fs10nYaz Then
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
                    wsPivot.Range(hedefler(j)).ClearContents   ' eski bakiye kalmasin
                    uyarilar = uyarilar & "- FS10N: " & Metin(hesaplar(j)) & " hesabinin bakiyesi okunamadi" & vbCrLf
                End If

                SapSifirla session
            End If
        Next j
    End If

    '--- 8. XLSX KAYDET ---
    adim = "Kaydetme"
    If ThisWorkbook.Path = "" Then
        MsgBox "Bu dosya henuz diske kaydedilmemis; rapor kopyasi olusturulamadi.", vbExclamation
        hataOldu = True: GoTo Kapanis
    End If

    If LCase(Left(ThisWorkbook.Path, 4)) = "http" Then
        raporKlasoru = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\Raporlar\"
        uyarilar = uyarilar & "- Dosya OneDrive/SharePoint'te calisiyor; rapor yerel klasore kaydedildi" & vbCrLf
    Else
        raporKlasoru = ThisWorkbook.Path & "\Raporlar\"
    End If
    If Len(Dir(raporKlasoru, vbDirectory)) = 0 Then MkDir raporKlasoru

    nokta = InStrRev(ThisWorkbook.Name, ".")
    If nokta > 0 Then
        kayitDosyaAdi = Left(ThisWorkbook.Name, nokta - 1) & ".xlsx"
    Else
        kayitDosyaAdi = ThisWorkbook.Name & ".xlsx"
    End If
    kayitTamYol = raporKlasoru & kayitDosyaAdi

    Set wbAcik = Nothing
    On Error Resume Next
    Set wbAcik = Workbooks(kayitDosyaAdi)
    On Error GoTo HataYakala
    If Not wbAcik Is Nothing Then
        uyarilar = uyarilar & "- '" & kayitDosyaAdi & "' acik oldugu icin rapor kopyasi KAYDEDILMEDI" & vbCrLf
        kayitTamYol = "(kaydedilmedi)"
    Else
        On Error Resume Next
        If Len(Dir(kayitTamYol)) > 0 Then Kill kayitTamYol
        On Error GoTo HataYakala

        ThisWorkbook.Sheets(Array(wsHedef.Name, wsPivot.Name)).Copy
        With ActiveWorkbook
            .SaveAs Filename:=kayitTamYol, FileFormat:=xlOpenXMLWorkbook
            .Close False
        End With
    End If

    GoTo Kapanis

HataYakala:
    hataOldu = True
    Application.StatusBar = False
    mesaj = "HATA!" & vbCrLf & vbCrLf & _
            "Adim: " & adim & vbCrLf & _
            "Aciklama: " & Err.Description & vbCrLf & _
            "Hata No: " & Err.Number
    If eksik <> "" Then mesaj = mesaj & vbCrLf & vbCrLf & "Alinamayan plant'ler:" & vbCrLf & eksik
    MsgBox mesaj, vbCritical, "Hata"
    Resume Kapanis

Kapanis:
    '--- 9. KAPANIS ---
    On Error Resume Next
    If Not wbTemp Is Nothing Then wbTemp.Close False
    If Not wbAnaveri Is Nothing Then wbAnaveri.Close False

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
        mesaj = "Islem Tamamlandi!" & vbCrLf & vbCrLf & _
                "SAP: " & sapBilgi & vbCrLf & _
                "Kayit: " & kayitTamYol & vbCrLf & vbCrLf & _
                "Alinan plant'ler: " & alinan
        If eksik <> "" Then mesaj = mesaj & vbCrLf & vbCrLf & "!!! ALINAMAYAN PLANT'LER:" & vbCrLf & eksik
        If uyarilar <> "" Then mesaj = mesaj & vbCrLf & "UYARILAR:" & vbCrLf & uyarilar
        If eksik <> "" Or uyarilar <> "" Then
            MsgBox mesaj, vbExclamation, "Tamamlandi - kontrol edin"
        Else
            MsgBox mesaj, vbInformation, "Tamamlandi"
        End If
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
            Sleep 250
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

' Baslik karsilastirmasi icin: kucuk harf, bosluk/satir sonu temizligi
Private Function Normallestir(ByVal s As String) As String
    s = Replace(Replace(Replace(s, ChrW(160), " "), vbLf, " "), vbCr, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Normallestir = LCase(Trim(s))
End Function

' 1. satirda basligi ada gore bul (yoksa 0)
Private Function BaslikBul(ByVal ws As Worksheet, ByVal ad As String) As Long
    Dim c As Long, son As Long
    son = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To son
        If Normallestir(ws.Cells(1, c).Text) = Normallestir(ad) Then
            BaslikBul = c
            Exit Function
        End If
    Next c
End Function

' Bir sayfanin baslik satirinin "parmak izi" (plant'ler arasi karsilastirma)
Private Function BaslikImza(ByVal ws As Worksheet, ByVal sonSutun As Long) As String
    Dim c As Long, s As String
    For c = 1 To sonSutun
        s = s & "|" & Normallestir(ws.Cells(1, c).Text)
    Next c
    BaslikImza = s
End Function

' Sutun numarasi -> harf (31 -> AE)
Private Function SutunHarfi(ByVal kol As Long) As String
    SutunHarfi = Split(Cells(1, kol).Address(True, False), "$")(0)
End Function

' Dictionary'nin ilk N anahtarini virgullu metin yap
Private Function IlkNAnahtar(ByVal d As Object, ByVal n As Long) As String
    Dim k As Variant, sayac As Long, s As String
    For Each k In d.Keys
        sayac = sayac + 1
        If sayac > n Then
            s = s & " ... (+" & (d.Count - n) & ")"
            Exit For
        End If
        If s <> "" Then s = s & ", "
        s = s & k
    Next k
    IlkNAnahtar = s
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
