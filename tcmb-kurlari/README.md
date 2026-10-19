# TCMB Kurları

Bu araç, Türkiye Cumhuriyet Merkez Bankası'nın (TCMB) günlük gösterge niteliğindeki döviz kurlarını tarihe göre getirir. Python betiği tablo ve CSV üretir, VBA modülü ise Windows'taki Excel'de formül olarak çalışır.

Araç yalnızca TCMB bülteninde yazan kuru okur. Gümrükte, vergi hesabında ya da sözleşmede hangi kurun kullanılacağını söylemez. Bu konuda ilgili kurumun duyurusuna ve sözleşmenizin metnine bakın.

## Nasıl kullanılır

Bu araç için hazır bir Excel dosyası (.xlsx) yoktur. Kurları internetten çekmek bir makro ya da betik gerektirir. İki yol vardır.

### Python ile (Windows, Mac, Linux)

1. Bilgisayarınızda Python 3.9 ya da daha yeni bir sürüm olmalıdır. Ek paket kurmanız gerekmez.
2. `tcmb_kur.py` dosyasını indirin.
3. Terminalde ya da komut isteminde dosyanın klasörüne geçin ve şu komutu çalıştırın:

   ```
   python3 tcmb_kur.py 2026-10-18 USD JPY --tur efektif-satis
   ```

   Çıktı şöyledir:

   ```
   İstenen tarih  Bülten tarihi  Bülten no  Döviz  Tür              Kur (TL)
   -------------  -------------  ---------  -----  ---------------  --------
   2026-10-18     2026-10-16     2026/196   USD    BanknoteSelling  49.5204
   2026-10-18     2026-10-16     2026/196   JPY    BanknoteSelling  0.314965
   ```

4. Bir tarih aralığını CSV dosyasına yazmak için:

   ```
   python3 tcmb_kur.py --bas 2026-10-01 --bit 2026-10-16 USD EUR --csv kurlar.csv
   ```

   Dosyada her istenen gün ve her döviz için bir satır bulunur. Sütunlar şunlardır: `istenen_tarih`, `bulten_tarihi`, `bulten_no`, `doviz`, `tur`, `kur`.

5. Tarihi `2026-10-16` ya da `16.10.2026` biçiminde yazabilirsiniz. Tüm seçenekler için `python3 tcmb_kur.py --help` komutunu çalıştırın.

Betiği başka bir Python programından da kullanabilirsiniz. `kur_getir(tarih, "USD")` komutu kuru, kullanılan bülten tarihini ve bülten numarasını döndürür.

### Excel ile (VBA, yalnızca Windows)

1. Excel'de Alt+F11 tuşlarına basın.
2. Insert > Module menüsüyle boş bir modül ekleyin.
3. `TCMBKur.bas` dosyasının içeriğini kopyalayıp modüle yapıştırın. File > Import kullanmayın, yapıştırın.
4. Çalışma kitabını "Excel Macro-Enabled Workbook (.xlsm)" olarak kaydedin.
5. Hücreye şu formülleri yazın. Türkçe Excel'de bağımsız değişkenler noktalı virgülle, İngilizce Excel'de virgülle ayrılır.

   ```
   =TCMBKur(A2;"USD")
   =TCMBKur(A2;"EUR";"ForexBuying")
   =TCMBBultenTarihi(A2)
   ```

   `A2` hücresinde bir tarih olmalıdır. `TCMBBultenTarihi` bir tarih serisi döndürür, hücreyi tarih biçimine çevirin.

İndirdiğiniz dosyayı Excel Korumalı Görünüm'de açarsa "Düzenlemeyi Etkinleştir" düğmesine tıklayın. Şirketinizin BT ilkesi makroları tamamen engelliyor olabilir. Bu durumda Python yolunu kullanın.

## Kur türleri

| Betikte ve VBA'da | XML alanı | Anlamı |
|---|---|---|
| `doviz-alis` | `ForexBuying` | Döviz alış |
| `doviz-satis` (varsayılan) | `ForexSelling` | Döviz satış |
| `efektif-alis` | `BanknoteBuying` | Efektif alış |
| `efektif-satis` | `BanknoteSelling` | Efektif satış |

Sonuç her zaman 1 birim dövizin TL karşılığıdır. TCMB bazı dövizleri 100 birim için yayımlar (örneğin JPY). Araç bu kurları 100'e böler.

## Sınırlar

- İstenen günün bülteni yoksa (hafta sonu, resmî tatil ya da bülten henüz yayımlanmadıysa) araç en çok 10 gün geriye gider. Kullanılan bültenin tarihi her zaman çıktıda görünür ve istenen tarihten farklı olabilir.
- Bazı dövizlerde bazı alanlar boştur. Örneğin 16 Ekim 2026 bülteninde XDR (SDR) için döviz satış alanı boştur. Araç bu durumda hata mesajı verir, tahmin yürütmez.
- Bu araç internet bağlantısı gerektirir. TCMB adres yapısını değiştirirse çalışmaz.
- VBA modülü MSXML kullanır, bu yüzden Mac'teki Excel'de çalışmaz. Şirket proxy'si ya da güvenlik duvarı da isteği engelleyebilir.
- Excel'de çok sayıda farklı tarih için formül yazarsanız her yeni tarih ayrı bir internet isteği yapar. Dosya açılırken ya da yeniden hesaplanırken Excel bir süre yanıt vermeyebilir.
- CSV dosyasındaki kur sütununda ondalık ayırıcı noktadır. Bölge ayarı Türkçe olan Excel'de dosyayı çift tıklayarak açarsanız sayılar yanlış görünebilir. Bu durumu Excel'de denemedik. Veri > Metin/CSV'den yolunu ve ondalık ayırıcı olarak noktayı seçmenizi öneririz.
- VBA kaynak dosyasında Türkçe karakterler vardır. Türkçe olmayan bir Windows bölge ayarında hata mesajlarındaki bazı harfler soru işareti gibi görünebilir. Hesaplama bundan etkilenmez.
- Resmî bir işlemde kullanmadan önce değeri TCMB'nin kendi sayfasıyla karşılaştırın.

## Kaynak

TCMB günlük bülten adresi: `https://www.tcmb.gov.tr/kurlar/YYYYMM/DDMMYYYY.xml` (örnek: `https://www.tcmb.gov.tr/kurlar/202610/16102026.xml`). En son bülten: `https://www.tcmb.gov.tr/kurlar/today.xml`. Adres biçimi 19 Ekim 2026'da kontrol edildi.

## Nasıl test edildi

Testler 19 Ekim 2026'da macOS'ta Python 3.9.6 ile çalıştırıldı. Komut:

```
python3 -m unittest discover -s tcmb-kurlari/tests -v
```

- Çevrimdışı 18 test geçti. Bunlar TCMB'nin 16 Ekim 2026 tarihli gerçek bülteninin kayıtlı bir kopyasını (`tests/fixtures/tcmb-16102026.xml`) ve sahte bir ağ yanıtını kullanır. USD ve EUR değerleri, JPY'nin 100'e bölünmesi, kur türü takma adları, bilinmeyen döviz ve boş alan hataları, hafta sonu için geriye gitme ve CSV çıktısı denendi.
- Canlı test `LIVE=1 python3 -m unittest discover -s tcmb-kurlari/tests -v` komutuyla çalışır ve TCMB sunucusuna bağlanır. Bu testle 16.10.2026 bülteni (no 2026/196) alındı ve USD döviz satış kuru 49.4463 olarak doğrulandı. 18 Ekim 2026 Pazar günü için istek 16 Ekim bültenine döndü.
- VBA modülü (`TCMBKur.bas`) çalıştırılmadı. Test makinesi bir Mac olduğu için Windows Excel ve MSXML yoktu. Kod yalnızca okunarak gözden geçirildi. Gerçek kullanımda sorun çıkarsa lütfen bildirin.
- Windows'ta Python, Linux'ta Python ve Python 3.9 dışındaki sürümler denenmedi.

## English

This tool fetches the Central Bank of the Republic of Türkiye (TCMB) daily indicative exchange rates for a given date. The Python script prints a table or writes a CSV for a date range, and the VBA module provides `TCMBKur` and `TCMBBultenTarihi` as worksheet functions in Windows Excel. If no bulletin exists for the requested day (weekend, holiday), it steps back up to 10 days and reports which bulletin it used. Rates quoted per 100 units, such as JPY, are divided to give the TL value of one unit. The tool only reads the published rate. It does not say which rate customs or tax rules require. The Python script and its tests were run on macOS with Python 3.9.6, including one live fetch of the 16 October 2026 bulletin. The VBA module could not be run on the Mac used for development and has not been tested in Excel.

MIT lisansı.

LogisticsIQ hazırladı: https://logisticsiq.org
