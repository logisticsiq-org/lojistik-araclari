'==============================================================================
' TCMBKur.bas  |  TCMB günlük döviz kurlarını Excel formülü olarak getirir
'
' Amaç:
'   =TCMBKur(A2; "USD")                    1 USD için TL döviz satış kuru
'   =TCMBKur(A2; "EUR"; "ForexBuying")     1 EUR için TL döviz alış kuru
'   =TCMBBultenTarihi(A2)                  Kurun alındığı bültenin tarihi
'
' Kurulum (yalnızca Windows için Excel, MSXML kullanır):
'   1. Alt+F11 ile Visual Basic düzenleyicisini açın.
'   2. Insert > Module ile boş bir modül ekleyin.
'   3. Bu dosyanın içeriğini kopyalayıp modüle yapıştırın.
'      (File > Import yerine yapıştırmayı kullanın. Türkçe harfler böylece bozulmaz.)
'   4. Çalışma kitabını "Excel Macro-Enabled Workbook (.xlsm)" olarak kaydedin.
'   Şirket BT ilkesi makroları engelliyor olabilir. Bu durumda Python
'   betiğini (tcmb_kur.py) kullanın.
'
' Kur türleri (üçüncü bağımsız değişken):
'   ForexSelling (varsayılan), ForexBuying, BanknoteSelling, BanknoteBuying
'   Takma adlar: doviz-satis, doviz-alis, efektif-satis, efektif-alis
'
' Davranış:
'   - Seçilen günün bülteni yoksa (hafta sonu, resmi tatil) en çok 10 gün geriye gidilir.
'   - JPY gibi 100 birimlik kurlar 100'e bölünür, sonuç 1 birimin TL karşılığıdır.
'   - Hata olursa hücreye "Hata: ..." metni yazılır.
'   - İnternet bağlantısı gerekir. Aynı adres bir Excel oturumunda bir kez indirilir.
'
' Kaynak: https://www.tcmb.gov.tr/kurlar/YYYYMM/DDMMYYYY.xml
' Lisans: MIT
' LogisticsIQ hazırladı: https://logisticsiq.org
'==============================================================================
Option Explicit

Private Const GERI_GUN As Long = 10
Private Const KULLANICI_ADI As String = "lojistik-araclari (LogisticsIQ)"

' Adres -> XML metni. Bülten yoksa (404) boş metin saklanır.
Private mOnbellek As Object

' Tarihi 2 haneli gün ve ay ile kurar. Format$ kullanılmaz, çünkü Excel dil ayarına
' göre farklı davranabilir.
Private Function BultenUrl(ByVal Tarih As Date) As String
    Dim Yil As String, Ay As String, Gun As String
    Yil = CStr(Year(Tarih))
    Ay = Right$("0" & CStr(Month(Tarih)), 2)
    Gun = Right$("0" & CStr(Day(Tarih)), 2)
    BultenUrl = "https://www.tcmb.gov.tr/kurlar/" & Yil & Ay & "/" & Gun & Ay & Yil & ".xml"
End Function

Private Function TarihMetni(ByVal Tarih As Date) As String
    TarihMetni = Right$("0" & CStr(Day(Tarih)), 2) & "." & Right$("0" & CStr(Month(Tarih)), 2) & "." & CStr(Year(Tarih))
End Function

' Dönüş: 200 (Metin dolu), 404 (bülten yok) ya da başka bir HTTP kodu.
Private Function Indir(ByVal Url As String, ByRef Metin As String) As Long
    Dim Http As Object
    If mOnbellek Is Nothing Then Set mOnbellek = CreateObject("Scripting.Dictionary")
    If mOnbellek.Exists(Url) Then
        Metin = mOnbellek(Url)
        If Len(Metin) > 0 Then Indir = 200 Else Indir = 404
        Exit Function
    End If
    Set Http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Http.setTimeouts 15000, 15000, 15000, 15000
    Http.Open "GET", Url, False
    Http.setRequestHeader "User-Agent", KULLANICI_ADI
    Http.send
    Select Case Http.Status
        Case 200
            Metin = Http.responseText
            mOnbellek(Url) = Metin
            Indir = 200
        Case 404
            Metin = ""
            ' Bugünün bülteni gün içinde yayımlanır. Bu yüzden bugünün 404 yanıtı saklanmaz.
            If InStr(Url, "/" & Right$("0" & CStr(Day(Date)), 2) & Right$("0" & CStr(Month(Date)), 2) & CStr(Year(Date)) & ".xml") = 0 Then mOnbellek(Url) = ""
            Indir = 404
        Case Else
            Metin = ""
            Indir = Http.Status
    End Select
End Function

' İstenen günden başlayıp geriye giderek bülten bulur.
' Dönüş boş metinse başarılıdır. Aksi halde hata metnidir.
Private Function BultenBul(ByVal Tarih As Date, ByRef Belge As Object, ByRef BultenTarihi As Date) As String
    Dim i As Long, Kod As Long, Metin As String
    For i = 0 To GERI_GUN
        BultenTarihi = Tarih - i
        Kod = Indir(BultenUrl(BultenTarihi), Metin)
        If Kod = 200 Then
            Set Belge = CreateObject("MSXML2.DOMDocument.6.0")
            Belge.async = False
            Belge.setProperty "SelectionLanguage", "XPath"
            If Len(Metin) > 0 Then
                If AscW(Left$(Metin, 1)) = &HFEFF Then Metin = Mid$(Metin, 2)
            End If
            If Not Belge.LoadXML(Metin) Then
                BultenBul = "Hata: " & TarihMetni(BultenTarihi) & " bülteni XML olarak okunamadı"
            End If
            Exit Function
        ElseIf Kod <> 404 Then
            BultenBul = "Hata: TCMB sunucusu HTTP " & CStr(Kod) & " döndürdü"
            Exit Function
        End If
    Next i
    BultenBul = "Hata: " & TarihMetni(Tarih) & " ve önceki " & CStr(GERI_GUN) & " gün için bülten bulunamadı"
End Function

' Takma adı XML alan adına çevirir. Tanınmazsa boş metin döndürür.
Private Function TurAdi(ByVal Tur As String) As String
    Select Case LCase$(Replace(Replace(Trim$(Tur), "_", "-"), " ", ""))
        Case "forexselling", "doviz-satis": TurAdi = "ForexSelling"
        Case "forexbuying", "doviz-alis": TurAdi = "ForexBuying"
        Case "banknoteselling", "efektif-satis": TurAdi = "BanknoteSelling"
        Case "banknotebuying", "efektif-alis": TurAdi = "BanknoteBuying"
        Case Else: TurAdi = ""
    End Select
End Function

' Üç harfli ASCII döviz kodu mu? XPath ifadesine yalnızca doğrulanmış kod girer.
Private Function DovizKoduGecerli(ByVal Kod As String) As Boolean
    Dim i As Long, K As Long
    If Len(Kod) <> 3 Then Exit Function
    For i = 1 To 3
        K = AscW(Mid$(Kod, i, 1))
        If K < 65 Or K > 90 Then Exit Function
    Next i
    DovizKoduGecerli = True
End Function

Public Function TCMBKur(ByVal Tarih As Date, ByVal Doviz As String, Optional ByVal Tur As String = "ForexSelling") As Variant
    Dim Belge As Object, BultenTarihi As Date, Hata As String
    Dim Alan As String, Kod As String, DugumKur As Object, DugumBirim As Object
    Dim Birim As Double, KurMetni As String

    On Error GoTo Sorun
    Alan = TurAdi(Tur)
    If Len(Alan) = 0 Then
        TCMBKur = "Hata: kur türü anlaşılamadı (" & Tur & ")"
        Exit Function
    End If
    Kod = UCase$(Trim$(Doviz))
    If Not DovizKoduGecerli(Kod) Then
        TCMBKur = "Hata: döviz kodu 3 harfli olmalı (örnek: USD)"
        Exit Function
    End If

    If Tarih < DateSerial(1996, 1, 1) Or Tarih > Date Then
        TCMBKur = "Hata: geçerli bir tarih girin (bugün ya da daha önce)"
        Exit Function
    End If

    Hata = BultenBul(Tarih, Belge, BultenTarihi)
    If Len(Hata) > 0 Then
        TCMBKur = Hata
        Exit Function
    End If

    Set DugumKur = Belge.selectSingleNode("//Currency[@CurrencyCode='" & Kod & "']/" & Alan)
    If DugumKur Is Nothing Then
        TCMBKur = "Hata: " & Kod & " kodu " & TarihMetni(BultenTarihi) & " bülteninde yok"
        Exit Function
    End If
    KurMetni = Trim$(DugumKur.Text)
    If Len(KurMetni) = 0 Then
        TCMBKur = "Hata: " & Kod & " için " & Alan & " alanı " & TarihMetni(BultenTarihi) & " bülteninde boş"
        Exit Function
    End If

    Set DugumBirim = Belge.selectSingleNode("//Currency[@CurrencyCode='" & Kod & "']/Unit")
    If DugumBirim Is Nothing Then
        Birim = 1
    Else
        Birim = Val(DugumBirim.Text)
    End If
    If Birim <= 0 Then
        TCMBKur = "Hata: " & Kod & " için birim okunamadı"
        Exit Function
    End If

    ' Val() her zaman nokta ondalık ayırıcısını bekler, Windows bölge ayarından etkilenmez.
    TCMBKur = Val(KurMetni) / Birim
    Exit Function
Sorun:
    TCMBKur = "Hata: " & Err.Description
End Function

' Kurun alındığı bültenin tarihini döndürür (hücreyi tarih biçimine çevirin).
Public Function TCMBBultenTarihi(ByVal Tarih As Date) As Variant
    Dim Belge As Object, BultenTarihi As Date, Hata As String
    On Error GoTo Sorun
    If Tarih < DateSerial(1996, 1, 1) Or Tarih > Date Then
        TCMBBultenTarihi = "Hata: geçerli bir tarih girin (bugün ya da daha önce)"
        Exit Function
    End If
    Hata = BultenBul(Tarih, Belge, BultenTarihi)
    If Len(Hata) > 0 Then
        TCMBBultenTarihi = Hata
    Else
        TCMBBultenTarihi = BultenTarihi
    End If
    Exit Function
Sorun:
    TCMBBultenTarihi = "Hata: " & Err.Description
End Function
