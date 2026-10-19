# Demuraj, detention ve ardiye maliyeti (tahmin)

Bu araç, konteynerin tahliye, terminal çıkışı ve boş iade tarihlerinden demuraj, detention ve ardiye maliyetini tahmin eder. Siz taşıyıcınızın ve terminalinizin tarifesini kademe kademe girersiniz, araç günleri ve tutarları hesaplar.

**Sonuç bir tahmindir.** Tarifeler taşıyıcıya, limana ve sözleşmeye göre değişir. Geçerli olan, sizin sözleşmenizdeki ve terminal tarifenizdeki koşullardır. Bu aracı fatura kontrolünde yalnızca ilk kaba kontrol olarak kullanın.

## Tanımlar

- Demuraj: Dolu konteynerin tahliyeden terminal çıkışına kadar terminalde kaldığı günler.
- Detention: Konteynerin terminalden çıkışından boş iadeye kadar terminal dışında geçen günler.
- Birleşik süre: Bazı taşıyıcılar tahliyeden boş iadeye kadar tek bir serbest süre uygular. Bu durumda tek tarife tablosu kullanılır.
- Ardiye: Terminalin kendi tarifesiyle, demurajla aynı dönem için kestiği depolama bedeli.

Tüm sayımlar takvim günüdür. Hafta sonu ve tatiller sayıma dahildir.

## Dosyalar

- `demuraj-detention.xlsx`: Formüllerle çalışan Excel dosyası. Makro içermez. İçindeki tarihler ve tarifeler örnektir.
- `demuraj.py`: Aynı hesabı yapan Python modülü ve komut satırı aracı.
- `ornek.json`: Python aracı için örnek girdi. Excel dosyasındaki örnekle aynıdır.
- `olustur_xlsx.py`: Excel dosyasını yeniden üretir.
- `tests/`: Birim testleri ve Excel formüllerini Python ile karşılaştıran testler.

## Excel dosyası nasıl kullanılır

1. İndirdiğiniz dosya Korumalı Görünüm'de açılırsa Düzenlemeyi Etkinleştir düğmesine tıklayın.
2. Hesapla sayfasında ayarları girin: konteyner sayısı, para birimi, hesap şekli (Ayrı veya Birleşik) ve ilk günün sayılıp sayılmayacağı.
3. Tahliye, çıkış ve boş iade tarihlerini yazın. Sıra yanlışsa veya bir tarih eksikse Uyarı satırı bunu söyler ve sonuç gösterilmez.
4. Tarife tablolarını sözleşmenizdeki gibi doldurun. Her satır bir gün aralığı ve o aralıktaki günlük ücrettir. Serbest süre, ücreti 0 olan bir satırdır. Bitiş gününü boş bırakırsanız satır son güne kadar sürer. Her tabloda en fazla 5 satır vardır.
5. Ardiye tablosu isteğe bağlıdır. Boş bırakırsanız ardiye hesaplanmaz.
6. İsterseniz TRY kurunu elle yazın. Güncel kur için bu depodaki `../tcmb-kurlari/` aracına bakabilirsiniz. Kur girmezseniz TRY toplamı boş kalır.
7. Sonuç bölümünde süreleri, tablo bazında tutarları, genel toplamı ve bir özet cümlesini görürsünüz.

Sarı hücreler girdidir. Kalın yazılı hücreler sonuçtur. Satır notu sütunu, hatalı bir tarife satırını gösterir.

## Python ile nasıl kullanılır

Python 3.9 veya daha yenisi yeterlidir. Ek paket gerekmez.

1. Bu klasöre geçin: `cd demuraj-detention`
2. `ornek.json` dosyasını kopyalayıp kendi bilgilerinizle değiştirin.
3. Çalıştırın: `python3 demuraj.py ornek.json`

JSON dosyasında şu alanlar bulunur:

- `konteyner_sayisi`, `konteyner_tipi`, `para_birimi` (`USD`, `EUR` veya `TRY`)
- `mod`: `ayri` veya `birlesik`
- `ilk_gun_sayilir`: `true` ise başlangıç günü 1. gündür, `false` ise gün sayısı bitiş eksi başlangıçtır
- `kur_try`: İsteğe bağlı, yoksa `null`
- `tahliye`, `cikis`, `bos_iade`: `YYYY-AA-GG` biçiminde tarihler
- `tarifeler`: `demuraj`, `detention`, `birlesik` ve `ardiye` listeleri. Her satır `[başlangıç günü, bitiş günü, günlük ücret]` biçimindedir. Bitiş günü `null` ise satır sonuna kadar sürer.

Birleşik modda `cikis` yalnızca ardiye tablosu doluysa gereklidir.

## Hesap kuralı

Bir tablo satırına düşen gün sayısı şöyle hesaplanır:

```
en büyük değer (0, en küçük değer (toplam gün, bitiş günü) eksi başlangıç günü artı 1)
```

Tutar, satıra düşen gün sayısı x günlük ücret x konteyner sayısıdır. Ayrı modda demuraj tahliyeden çıkışa, detention çıkıştan boş iadeye kadar hesaplanır. Birleşik modda tahliyeden boş iadeye kadar tek tablo kullanılır. Ardiye her iki modda tahliyeden çıkışa kadardır.

Örnek: 2 konteyner, tahliye 1 Eylül 2026, çıkış 12 Eylül, boş iade 20 Eylül, ilk gün sayılıyor. Demuraj süresi 12 gündür. Tarifede 1 ile 7. gün ücretsiz, 8 ile 14. gün 40 USD olduğundan 5 ücretli gün x 40 x 2 = 400 USD çıkar. Detention 9 gün, 280 USD. Ardiye 12 gün, 268 USD. Genel toplam 948 USD.

## Sınırlar

- Dosyadaki örnek tarifeler uydurmadır. Gerçek bir taşıyıcıya veya terminale ait değildir.
- Tüm sayımlar takvim günüdür. Hafta sonunu veya tatili saymayan tarifeler için bu araç uygun değildir.
- Vergi, KDV, işlem ücreti, serbest süre uzatması, azami tutar ve konteyner tipine göre farklar hesaba katılmaz. Bunları tarifeye satır olarak yansıtamıyorsanız sonucu elle düzeltmeniz gerekir.
- Tarife satırlarının çakışmamasını siz sağlarsınız. Çakışan satırlar iki kez ücretlendirilir. Aralarında boşluk bırakılan günler ücretlendirilmez.
- Kur kutusu elle doldurulur. Araç kur çekmez.
- Taşıyıcınız ilk günü farklı sayıyorsa İlk gün sayılır mı ayarını tarifenizdeki kurala göre seçin.
- İlk gün sayılıyorsa çıkış günü hem demurajın son günü hem detention'ın ilk günü olur. Taşıyıcınız bu günü yalnızca bir tarafta sayıyorsa sonucu buna göre düzeltin.

## Kaynak

Bu araçta hazır tarife yoktur. Tarifeyi taşıyıcınızın ve terminalinizin yayımladığı belgeden veya sözleşmenizden alın. Tanımlar genel kullanıma dayanır. Bir taşıyıcı bunları farklı tanımlıyorsa onun tanımı geçerlidir.

## Nasıl test edildi

Python modülü için elle hesaplanmış birim testleri yazıldı. Bunlar serbest süre içinde iadeyi, kademe sınırlarını (7, 8, 14, 15 ve 30. gün), açık uçlu son kademeyi, ilk gün Evet ve Hayır durumlarını, birleşik modu, boş ardiye tablosunu, tarih sırası hatalarını, konteyner sayısı çarpanını ve TRY dönüşümünü kapsar.

Excel formülleri LibreOffice ile yeniden hesaplatıldı ve Python sonuçlarıyla karşılaştırıldı. On üç doğru girdi senaryosunda süreler, her tablo satırındaki gün sayısı ve tutar, ara toplamlar, genel toplam, TRY toplamı ve özet cümlesi karşılaştırıldı. Dokuz hatalı girdi senaryosunda Uyarı hücresinin metninin Python hata metniyle aynı olduğu ve sonuç hücrelerinin boş kaldığı doğrulandı.

Dosya Microsoft Excel'de açılarak denenmedi. Kullanılan işlevler (IF, MIN, MAX, SUM, SUMPRODUCT, COUNTA, FIXED, TRIM) Excel 2010 ve sonrasında bulunur.

Testleri çalıştırmak için:

```
cd demuraj-detention
python3 -m unittest discover -s tests -v
```

LibreOffice kurulu değilse Excel karşılaştırma testleri atlanır.

## English

This tool estimates demurrage, detention and port storage costs for a container from the discharge, gate-out and empty return dates. You enter the tier tables from your own carrier and terminal tariff, with free time as a row at rate 0, and the tool counts calendar days and multiplies by the daily rate and the container count. It supports separate demurrage and detention periods, or one combined free time. The result is an estimate only, and your contract and tariff decide the real amount. The example tariffs are invented. The Excel workbook was recalculated in LibreOffice and compared with the Python code, but it was not opened in Microsoft Excel. If the downloaded file opens in Protected View, click "Düzenlemeyi Etkinleştir" (Enable Editing). The code is released under the MIT licence.

LogisticsIQ hazırladı: https://logisticsiq.org
