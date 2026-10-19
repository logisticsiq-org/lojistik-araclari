'==============================================================================
' KontrolHanesi.bas  |  Konteyner (ISO 6346) ve AWB kontrol hanesi işlevleri
'
' Amaç:
'   =KonteynerKontrolHanesi("CSQU305438")   3 (10 karakterden kontrol hanesi hesaplar)
'   =KonteynerGecerli(A2)                   DOĞRU ya da YANLIŞ
'   =AWBGecerli(A2)                         DOĞRU ya da YANLIŞ
'
' Kurulum:
'   1. Alt+F11 ile Visual Basic düzenleyicisini açın.
'   2. Insert > Module ile boş bir modül ekleyin.
'   3. Bu dosyanın içeriğini kopyalayıp modüle yapıştırın.
'      (File > Import yerine yapıştırmayı kullanın. Türkçe harfler böylece bozulmaz.)
'   4. Çalışma kitabını "Excel Macro-Enabled Workbook (.xlsm)" olarak kaydedin.
'   Şirket BT ilkesi makroları engelliyor olabilir. Bu durumda makro içermeyen
'   konteyner-awb-kontrol.xlsx dosyasını kullanın.
'
' Notlar:
'   - Yalnızca saf VBA kullanılır, MSXML gibi bir bileşen gerekmez.
'   - Büyük ya da küçük harf, boşluk, tire, nokta ve eğik çizgi fark etmez.
'   - KonteynerKontrolHanesi 10 karakterli (kontrol hanesiz) ya da 11 karakterli numarayı
'     kabul eder. Biçim hatalıysa #DEĞER! hatası (VBA'da CVErr(2015)) döndürür.
'   - Geçerli bir kontrol hanesi yalnızca numaranın biçimce doğru olduğunu gösterir.
'     Konteynerin var olduğunu ya da o sahibe ait olduğunu göstermez.
'   - AWB kontrolü yalnızca 8 haneli seri numarasının ilk 7 hanesini kapsar. Havayolu
'     öneki (ilk 3 hane) kontrol hanesine dahil değildir.
'
' Lisans: MIT
' LogisticsIQ hazırladı: https://logisticsiq.org
'==============================================================================
Option Explicit

' Büyük harfe çevirir (yalnızca a-z) ve boşluk, tire, nokta, eğik çizgiyi siler.
' Dönüşüm bölge ayarından bağımsızdır (UCase kullanılmaz).
Private Function Normalle(ByVal No As String) As String
    Dim i As Long, K As Long, Sonuc As String
    For i = 1 To Len(No)
        K = AscW(Mid$(No, i, 1))
        Select Case K
            Case 32, 45, 46, 47
                ' boşluk, tire, nokta, eğik çizgi: atlanır
            Case 97 To 122
                Sonuc = Sonuc & ChrW(K - 32)
            Case Else
                Sonuc = Sonuc & Mid$(No, i, 1)
        End Select
    Next i
    Normalle = Sonuc
End Function

' ISO 6346 harf değeri: A=10, B=12, C=13 ... Z=38 (11'in katları atlanır).
' A-Z dışındaki karakter için 0 döndürür.
Private Function HarfDegeri(ByVal Harf As String) As Long
    Dim K As Long, J As Long, V As Long
    K = AscW(Harf)
    If K < 65 Or K > 90 Then Exit Function
    V = 10
    For J = 65 To K
        If V Mod 11 = 0 Then V = V + 1
        If J = K Then
            HarfDegeri = V
            Exit Function
        End If
        V = V + 1
    Next J
End Function

Private Function RakamMi(ByVal Karakter As String) As Boolean
    Dim K As Long
    K = AscW(Karakter)
    RakamMi = (K >= 48 And K <= 57)
End Function

' Normalleştirilmiş numaranın ilk 10 karakterini doğrular ve kontrol hanesini hesaplar.
' Biçim doğruysa True döner ve Hane dolar.
Private Function KonteynerHane(ByVal No As String, ByRef Hane As Long) As Boolean
    Dim i As Long, K As Long, Deger As Long, Toplam As Long, Agirlik As Long
    If Len(No) < 10 Then Exit Function
    For i = 1 To 3
        K = AscW(Mid$(No, i, 1))
        If K < 65 Or K > 90 Then Exit Function
    Next i
    Select Case Mid$(No, 4, 1)
        Case "U", "J", "Z"
            ' geçerli kategori harfi
        Case Else
            Exit Function
    End Select
    For i = 5 To 10
        If Not RakamMi(Mid$(No, i, 1)) Then Exit Function
    Next i
    Agirlik = 1
    For i = 1 To 10
        If i <= 4 Then
            Deger = HarfDegeri(Mid$(No, i, 1))
        Else
            Deger = AscW(Mid$(No, i, 1)) - 48
        End If
        Toplam = Toplam + Deger * Agirlik
        Agirlik = Agirlik * 2
    Next i
    Hane = (Toplam Mod 11) Mod 10
    KonteynerHane = True
End Function

' 10 karakterli (kontrol hanesiz) ya da 11 karakterli numaradan kontrol hanesini (0-9) hesaplar.
' Biçim hatalıysa #DEĞER! hatası döndürür.
Public Function KonteynerKontrolHanesi(ByVal No As String) As Variant
    Dim Temiz As String, Hane As Long
    Temiz = Normalle(No)
    Select Case Len(Temiz)
        Case 10
            ' kontrol hanesiz numara
        Case 11
            If Not RakamMi(Mid$(Temiz, 11, 1)) Then
                KonteynerKontrolHanesi = CVErr(2015)
                Exit Function
            End If
        Case Else
            KonteynerKontrolHanesi = CVErr(2015)
            Exit Function
    End Select
    If KonteynerHane(Temiz, Hane) Then
        KonteynerKontrolHanesi = Hane
    Else
        KonteynerKontrolHanesi = CVErr(2015)
    End If
End Function

' 11 karakterli konteyner numarası biçimce doğru ve kontrol hanesi tutuyorsa True döndürür.
Public Function KonteynerGecerli(ByVal No As String) As Boolean
    Dim Temiz As String, Hane As Long
    Temiz = Normalle(No)
    If Len(Temiz) <> 11 Then Exit Function
    If Not RakamMi(Mid$(Temiz, 11, 1)) Then Exit Function
    If Not KonteynerHane(Temiz, Hane) Then Exit Function
    KonteynerGecerli = ((AscW(Mid$(Temiz, 11, 1)) - 48) = Hane)
End Function

' 11 rakamlı AWB numarası (3 haneli önek + 8 haneli seri) geçerliyse True döndürür.
' Seri numarasının 8. hanesi, ilk 7 hanenin 7'ye bölümünden kalana eşit olmalıdır.
Public Function AWBGecerli(ByVal No As String) As Boolean
    Dim Temiz As String, i As Long, Seri As Long
    Temiz = Normalle(No)
    If Len(Temiz) <> 11 Then Exit Function
    For i = 1 To 11
        If Not RakamMi(Mid$(Temiz, i, 1)) Then Exit Function
    Next i
    Seri = CLng(Mid$(Temiz, 4, 7))
    AWBGecerli = ((AscW(Mid$(Temiz, 11, 1)) - 48) = (Seri Mod 7))
End Function
