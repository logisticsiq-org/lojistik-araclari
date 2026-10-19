# Konteyner ve AWB Kontrolü

Bu araç, konteyner numaralarının (ISO 6346) ve hava konşimento numaralarının (AWB) kontrol hanesini denetler. Bir Excel dosyası, bir VBA modülü ve bir Python betiği olarak kullanılabilir. Üçü de aynı kuralları uygular.

Geçerli bir kontrol hanesi yalnızca numaranın biçimce doğru olduğunu gösterir. Konteynerin gerçekten var olduğunu ya da o sahibe ait olduğunu göstermez. AWB için de aynı şey geçerlidir.

## Kurallar

**Konteyner numarası** 11 karakterdir: 3 harfli sahip kodu, kategori harfi (U, J ya da Z), 6 rakamlı seri numarası ve 1 kontrol hanesi. İlk 10 karakterin her birinin bir değeri vardır. Rakamlar kendi değerini taşır. Harfler A=10, B=12, C=13 diye gider ve 11'in katları atlanır (L=23, V=34, Z=38). Her değer, karakterin sırasına göre 2'nin kuvvetiyle çarpılır (ilk karakter 1, ikinci karakter 2, üçüncü karakter 4 ve böyle devam eder). Toplamın 11'e bölümünden kalan kontrol hanesidir. Kalan 10 çıkarsa kontrol hanesi 0 olur. Geçerli bir örnek `CSQU3054383` numarasıdır.

**AWB numarası** 11 rakamdır: 3 haneli havayolu öneki ve 8 haneli seri numarası. Seri numarasının 8. hanesi, ilk 7 hanenin 7'ye bölümünden kalana eşit olmalıdır. Geçerli bir örnek `020-12345675` numarasıdır.

Her iki türde de büyük ya da küçük harf, boşluk, tire, nokta ve eğik çizgi fark etmez. Bunlar silinir.

## Nasıl kullanılır

### Excel dosyası

1. `konteyner-awb-kontrol.xlsx` dosyasını indirin ve açın. Excel dosyayı Korumalı Görünüm'de açarsa "Düzenlemeyi Etkinleştir" düğmesine tıklayın.
2. Konteyner numaraları için "Konteyner" sayfasını açın. Numaraları A5 hücresinden başlayarak sarı hücrelere yapıştırın.
3. B sütunu temizlenmiş numarayı, C sütunu hesaplanan kontrol hanesini, D sütunu sonucu gösterir. Sonuç "Geçerli", "Kontrol hanesi hatalı (doğrusu: X)" ya da "Biçim hatalı" olur. Satır geçerliyse yeşile, değilse kırmızıya döner.
4. AWB numaraları için "AWB" sayfasını aynı şekilde kullanın.
5. Her sayfa 1000 satırlıktır. Daha fazlası için 1004. satırdaki formülleri aşağı doğru kopyalayın. Yardımcı hesaplar gizli sütunlardadır (Konteyner sayfasında F'den R'ye, AWB sayfasında F'den H'ye). Kopyalarken bunları da seçin.

Dosyada makro yoktur ve dosya koruma şifresi kullanmaz.

### Excel ile VBA (formül olarak)

1. Excel'de Alt+F11 tuşlarına basın.
2. Insert > Module menüsüyle boş bir modül ekleyin.
3. `KontrolHanesi.bas` dosyasının içeriğini kopyalayıp modüle yapıştırın. File > Import kullanmayın, yapıştırın.
4. Çalışma kitabını "Excel Macro-Enabled Workbook (.xlsm)" olarak kaydedin.
5. Hücrelere şu formülleri yazabilirsiniz:

   ```
   =KonteynerGecerli(A2)
   =AWBGecerli(A2)
   =KonteynerKontrolHanesi("CSQU305438")
   ```

   İlk ikisi DOĞRU ya da YANLIŞ döndürür. Üçüncüsü 10 karakterli bir numaradan kontrol hanesini hesaplar. Biçim hatalıysa #DEĞER! hatası verir.

Şirketinizin BT ilkesi makroları tamamen engelliyor olabilir. Bu durumda Excel dosyasını kullanın.

### Python ile

1. Python 3.9 ya da daha yeni bir sürüm gerekir. Ek paket kurmanız gerekmez.
2. `kontrol_hanesi.py` dosyasını indirin.
3. Tek tek numaraları denetlemek için:

   ```
   python3 kontrol_hanesi.py CSQU3054383 CSQU3054384 020-12345675 CSQX3054383
   ```

   Çıktı şöyledir:

   ```
   CSQU3054383    konteyner  Geçerli
   CSQU3054384    konteyner  Kontrol hanesi hatalı (doğrusu: 3)
   02012345675    awb        Geçerli
   CSQX3054383    konteyner  Biçim hatalı
   ```

   Harfle başlayan numara konteyner, rakamla başlayan numara AWB sayılır. `--tur konteyner` ya da `--tur awb` ile türü kendiniz belirleyebilirsiniz. Tüm numaralar geçerliyse çıkış kodu 0, değilse 1 olur.

4. Bir CSV dosyasındaki numaraları denetlemek için:

   ```
   python3 kontrol_hanesi.py --csv liste.csv --sutun 2 --cikti sonuc.csv --baslik
   ```

   `--sutun` numaranın bulunduğu sütundur ve 1'den başlar. `--baslik` ilk satır başlık ise kullanılır. Çıktı dosyasına `tur`, `normalle_no`, `durum` ve `beklenen_hane` sütunları eklenir. `durum` değeri `gecerli`, `hane_hatali` ya da `bicim_hatali` olur.

Betiği başka bir Python programından da kullanabilirsiniz. `konteyner_kontrol("CSQU3054383")` komutu `("gecerli", 3)` döndürür. `awb_kontrol("020-12345675")` komutu `("gecerli", 5)` döndürür.

## Sınırlar

- Kontrol hanesi numaranın var olduğunu, bir sahibe kayıtlı olduğunu ya da AWB'nin gerçekten düzenlendiğini göstermez. Sahip kodu ve havayolu öneki kayıtlarla karşılaştırılmaz.
- Kontrol hanesi her yazım hatasını yakalamaz. AWB'de ilk 3 hane kontrol hanesine dahil değildir, yani önekteki bir hata fark edilmez. Seri numarasının ilk 7 hanesinde farkı tam 7 olan bir rakam değişikliği de fark edilmez (örneğin 1 yerine 8). Konteynerde bazı harf değişiklikleri fark edilmez (örneğin A yerine K, çünkü değerleri 11 fark eder).
- Kategori harfi U, J ya da Z değilse numara biçim hatalı sayılır.
- Boşluk, tire, nokta ve eğik çizgi dışındaki karakterler (satır sonu ya da sert boşluk gibi) numarayı biçim hatalı yapar. Noktasız küçük "ı" harfi de geçersiz sayılır.
- Excel dosyasında AWB numarası sayı olarak yapıştırılırsa baştaki sıfır kaybolabilir ve numara biçim hatalı görünür. Giriş sütunu metin biçimindedir ama başka bir dosyadan kopyaladığınız hücrenin biçimi bunun üstüne yazabilir.
- Excel dosyası sayfa başına 1000 numara için hazırlanmıştır.
- Kurallar ISO 6346'daki kontrol hanesi yöntemine ve AWB numaralarındaki mod 7 yöntemine dayanır. Standardın ücretli tam metnini görmedik. 19 Ekim 2026'da harf değerlerini ve hesap yöntemini Wikipedia'daki [ISO 6346](https://en.wikipedia.org/wiki/ISO_6346) maddesiyle, kategori harflerini konteyner sahip kodlarını kaydeden [BIC'in sayfasıyla](https://www.bic-code.org/check-digit-calculator/), AWB yöntemini Wikipedia'daki [Air waybill](https://en.wikipedia.org/wiki/Air_waybill) maddesiyle karşılaştırdık. Üç araç ayrıca birbirini tutuyor.

## Nasıl test edildi

Testler 19 Ekim 2026'da macOS'ta Python 3.9.6 ve LibreOffice ile çalıştırıldı. Komut:

```
python3 -m unittest discover -s konteyner-awb-kontrol/tests -v
```

33 test geçti, hiçbiri atlanmadı. LibreOffice kurulu değilse 5 test atlanır.

- **Python (24 test).** Bilinen örnekler (`CSQU3054383` ve `020-12345675`), konteyner örneğindeki 63 tek rakam değişikliğinin hepsinin yakalanması, harf değeri tablosu, mod 11 sonucu 10 çıkınca hanenin 0 olması, normalleştirme, yanlış kategori harfi, yanlış uzunluk, AWB'deki farkı 7 olan değişikliklerin yakalanmadığı, önekin kontrole dahil olmadığı ve CSV çıktısı denendi.
- **Excel dosyası (9 test).** Dosyanın bir kopyasına 20 konteyner ve 16 AWB örneği (geçerli, yanlış hane, biçim hatalı, boş, boşluklu ve tireli yazımlar) yazıldı. LibreOffice formülleri yeniden hesapladı ve B, C, D sütunlarındaki her sonuç Python sonucuyla karşılaştırıldı. Hepsi eşit çıktı. Ayrıca yasak işlevlerin (LET, XLOOKUP, INDIRECT, OFFSET ve benzerleri) kullanılmadığı, giriş hücrelerinin sarı olduğu ve sayfaların korumasız olduğu denetlendi.
- **Denenmeyenler.** Excel dosyası gerçek Excel'de açılıp denenmedi, hesaplamayı LibreOffice yaptı. Dosyada yalnızca Excel 2010'da bulunan işlevler kullanıldı ama bu Excel'de çalıştırılarak doğrulanmadı. VBA modülü (`KontrolHanesi.bas`) hiçbir Excel'de çalıştırılmadı, çünkü test makinesi bir Mac'ti ve bu makinede VBA çalıştıracak bir Excel yoktu. Kodun mantığı elle izlenerek gözden geçirildi. Windows'ta Python ve Python 3.9 dışındaki sürümler denenmedi.

## English

This tool checks the check digit of container numbers (ISO 6346) and air waybill (AWB) numbers. It comes as a formula-only Excel workbook, a VBA module and a Python script. Paste numbers into the yellow cells of the workbook and each row shows the cleaned number, the expected check digit and a result in Turkish. A valid check digit only proves the number is well formed. It does not prove that the container exists or belongs to that owner. The AWB check does not cover the 3-digit airline prefix and cannot see digit changes of exactly 7 in the serial number. The Python tests and a LibreOffice recalculation of the workbook (36 sample numbers against the Python result) passed on macOS. The workbook was not opened in Excel and the VBA module was not run in any Excel.

MIT lisansı.

LogisticsIQ hazırladı: https://logisticsiq.org
