# İş günü hesaplayıcı (Türkiye)

Bu araç bir tarihe iş günü ekler ve iki tarih arasındaki iş günlerini sayar. Hesapta 2026, 2027 ve 2028 yıllarının resmî tatilleri dikkate alınır. Aynı hesabı Excel dosyasıyla da Python betiğiyle de yapabilirsiniz.

## Dosyalar

- `is-gunu-turkiye.xlsx`: Formüllerle çalışan Excel dosyası. Makro içermez.
- `is_gunu.py`: Aynı hesabı yapan Python modülü ve komut satırı aracı.
- `resmi-tatiller-2026-2028.csv`: Tatil listesi. İki araç da bu veriyi kullanır.
- `olustur_xlsx.py`: Excel dosyasını CSV'den yeniden üretir.
- `tests/`: Birim testleri ve Excel formüllerini Python ile karşılaştıran testler.

## Excel dosyası nasıl kullanılır

1. İndirdiğiniz dosya Korumalı Görünüm'de açılırsa Düzenlemeyi Etkinleştir düğmesine tıklayın.
2. Hesapla sayfasında Hafta sonu seçeneğini kontrol edin. Cumartesi günü çalışıyorsanız Yalnız Pazar tatil seçeneğini seçin.
3. Birinci blokta başlangıç tarihini ve eklenecek iş günü sayısını yazın. Geriye gitmek için eksi işaretli sayı yazın. Başlangıç günü sayılmaz.
4. İkinci blokta iki tarih yazın. İki tarih de sayıma dahildir. Sonuç olarak iş günü sayısı, takvim günü sayısı, tam gün tatil sayısı ve yarım gün tatil sayısı görünür. Aralıktaki tatillerin listesi altta yer alır.
5. Uyarı satırında bir mesaj varsa sonuç gösterilmez. Mesajı okuyup girdiyi düzeltin.

Sarı hücreler girdidir. Kalın yazılı hücreler sonuçtur. Tarihler gg.aa.yyyy biçiminde görünür.

## Python ile nasıl kullanılır

Python 3.9 veya daha yenisi yeterlidir. Ek paket gerekmez.

1. Bu klasöre geçin: `cd is-gunu-turkiye`
2. Bir tarihe iş günü ekleyin: `python3 is_gunu.py ekle 2026-05-25 3`
3. İki tarih arasındaki iş günlerini sayın: `python3 is_gunu.py say 2026-05-01 2026-05-31`
4. Cumartesi çalışılıyorsa `--cumartesi-calisilir` seçeneğini ekleyin: `python3 is_gunu.py say 2026-05-01 2026-05-31 --cumartesi-calisilir`
5. Bir aralıktaki tatilleri listeleyin: `python3 is_gunu.py tatiller 2026-03-01 2026-03-31`

Tarihleri YYYY-AA-GG veya GG.AA.YYYY biçiminde yazabilirsiniz. Modülü kendi kodunuzda da kullanabilirsiniz: `is_gunu_ekle`, `is_gunu_say` ve `aradaki_tatiller` işlevleri vardır.

## Hesap kuralları

- Pazar günü her zaman tatildir. Cumartesi varsayılan olarak tatildir, seçenekle iş günü yapılabilir.
- Tam gün resmî tatiller iş günü sayılmaz.
- Bayram arifeleri ve 28 Ekim öğleden sonrası yarım gün tatildir. Sabah çalışıldığı için bu günler iş günü sayılır. Yarım günler ayrıca listelenir ve sayılır.
- İşyeriniz bu yarım günlerde hiç çalışmıyorsa Excel'de Tatiller sayfasındaki Süre hücresini Tam gün yapın. Python için CSV dosyasında `yarim` değerini `tam` olarak değiştirin.
- Aynı güne iki tatil denk gelirse (örneğin 19 Mayıs 2027) gün bir kez sayılır. Excel listesinde bu iki tatil ayrı satırlarda görünür.
- Birinci blok Excel'deki WORKDAY.INTL işlevi gibi çalışır. Eklenecek sayı 0 ise başlangıç tarihi aynen döner.
- İkinci blok Excel'deki NETWORKDAYS.INTL işlevi gibi iki uç tarihi de sayar.

## Sınırlar

- Kapsam 1 Ocak 2026 ile 31 Aralık 2028 arasıdır. Başlangıç tarihi, son tarih veya sonuç bu aralığın dışına çıkarsa araç sonuç göstermez ve açık bir uyarı verir.
- Listede yalnızca 2429 sayılı kanuna dayanan resmî tatiller vardır. İdari izinler, köprü günleri ve işyerinize özel tatiller yoktur.
- Dinî bayramların tarihleri her yıl değişir ve Diyanet her yılın listesini ayrıca yayımlar. Bu yüzden listenin her yıl güncellenmesi gerekir.
- Sonuçları sözleşme, gümrük ve yasal süre hesaplarında kullanmadan önce kendi takviminizle kontrol edin.

## Her yıl nasıl güncellenir

1. Diyanet yeni yılın listesini yayımladığında tarihleri `resmi-tatiller-2026-2028.csv` dosyasının sonuna ekleyin. Sütunlar tarih (YYYY-AA-GG), ad, süre (`tam` veya `yarim`) ve kaynaktır.
2. `is_gunu.py` içindeki `KAPSAM_BITIS` değerini yeni yılın son gününe ilerletin.
3. Excel dosyasını yenilemek için `python3 olustur_xlsx.py` komutunu çalıştırın. Bunun için `openpyxl` gerekir. Betik kapsamı CSV'deki ilk ve son yıldan alır.
4. Yalnızca Excel dosyasını güncellemek isterseniz Tatiller sayfasının sağ tarafındaki yönergeleri izleyin. Yeni satırları tablonun altına ekleyip kapsamın son gününü yazmanız yeterlidir.

## Kaynak

Diyanet İşleri Başkanlığı resmî tatil listeleri, 2429 sayılı kanuna göre hazırlanmıştır:

- 2026: https://vakithesaplama.diyanet.gov.tr/icerik.php?icerik=158
- 2027: https://vakithesaplama.diyanet.gov.tr/icerik.php?icerik=159
- 2028: https://vakithesaplama.diyanet.gov.tr/icerik.php?icerik=193

## Nasıl test edildi

Python modülü için elle takvimden kontrol edilmiş birim testleri yazıldı. Bunlar Kurban Bayramı üzerinden iş günü eklemeyi, Ramazan Bayramı 2026 boyunca sayımı, Cumartesi çalışılan modu, negatif sayıyı, kapsam dışı tarihleri ve 19 Mayıs 2027 mükerrer kaydını kapsar.

Excel formülleri LibreOffice ile yeniden hesaplatıldı ve Python sonuçlarıyla karşılaştırıldı. On bir farklı girdi birleşimi, kapsam dışına çıkan üç durum ve geçersiz girdili bir durum denendi. Karşılaştırılan değerler sonuç tarihi, gün adı, iş günü sayısı, takvim günü sayısı, tatil sayıları ve tatil listesidir.

Dosya Microsoft Excel'de açılarak denenmedi. Kullanılan işlevler (WORKDAY.INTL, NETWORKDAYS.INTL, COUNTIFS, IFERROR, INDEX, MATCH) Excel 2010 ve sonrasında bulunur.

Testleri çalıştırmak için:

```
cd is-gunu-turkiye
python3 -m unittest discover -s tests -v
```

LibreOffice kurulu değilse Excel karşılaştırma testleri atlanır.

## English

This tool adds working days to a date and counts working days between two dates in Türkiye, using the official public holidays for 2026 to 2028 from the Diyanet lists. Half-day holidays (eves of religious holidays and the afternoon of 28 October) count as working days and are listed separately. Dates outside 2026 to 2028 return a clear warning instead of a result. The holiday list must be updated every year. The Excel workbook was recalculated in LibreOffice and compared with the Python code, but it was not opened in Microsoft Excel. If the downloaded file opens in Protected View, click "Düzenlemeyi Etkinleştir" (Enable Editing). The code is released under the MIT licence.

LogisticsIQ hazırladı: https://logisticsiq.org
