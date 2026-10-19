# Lojistik araçları

Lojistikte her gün elle yapılan küçük hesaplar için ücretsiz Excel dosyaları, makrolar ve Python betikleri. Araçları [LogisticsIQ](https://logisticsiq.org) hazırladı ve MIT lisansıyla herkesin kullanımına açtı.

## Araçlar

| Klasör | Ne yapar |
|---|---|
| [tcmb-kurlari](tcmb-kurlari/) | TCMB'nin gösterge niteliğindeki günlük döviz kurlarını istediğiniz tarih için getirir. Hafta sonu ya da tatil gününde bir önceki bülteni kullanır. |
| [konteyner-awb-kontrol](konteyner-awb-kontrol/) | Konteyner numarasının (ISO 6346) ve hava konşimentosu (AWB) numarasının kontrol hanesini denetler. Yazım hatalarını rezervasyondan ya da takipten önce yakalar. |
| [is-gunu-turkiye](is-gunu-turkiye/) | Türkiye'deki resmî tatilleri hesaba katarak bir tarihe iş günü ekler ve iki tarih arasındaki iş günlerini sayar. |
| [demuraj-detention](demuraj-detention/) | Demuraj, detention ve ardiye maliyetini, taşıyıcının tarifesini olduğu gibi girerek tahmin eder. |

## Nasıl kullanılır

Her klasörde ayrı bir README var. Kullanım adımları, aracın sınırları ve hangi testlerin yapıldığı orada yazıyor.

- **Excel dosyaları** (.xlsx) yalnızca formül içerir, makro içermez. İnternetten indirilen dosya Excel'de Korumalı Görünüm'de açılır. Dosyaya güveniyorsanız "Düzenlemeyi Etkinleştir" düğmesine basın.
- **Makrolar** dosya olarak değil, okunabilir kod (.bas) olarak duruyor. Kodu okuyup kendi çalışma kitabınıza yapıştırırsınız. Windows'taki Excel, internetten indirilen makrolu dosyaları varsayılan olarak engellediği için bu yolu seçtik. Şirketinizin BT politikası makroları tümüyle kapatmış olabilir.
- **Python betikleri** yalnızca Python'un standart kütüphanesini kullanır. Python 3.9 ve sonrası yeterlidir.

Tüm dosyaları birlikte indirmek için bu sayfadaki yeşil **Code** düğmesine basıp **Download ZIP** seçeneğini kullanın.

## Bir hata mı buldunuz?

Bir hesap sizin tarifenizle ya da beklediğiniz sonuçla uyuşmuyorsa veya işinize yarayacak başka bir araç aklınıza geliyorsa [bir issue açın](https://github.com/logisticsiq-org/lojistik-araclari/issues). Bu araçlar hesaplama yardımcısıdır. Son sözü her zaman sözleşmeniz, tarifeniz ve ilgili mevzuat söyler.

## English

Free Excel workbooks, VBA macros and Python scripts for everyday logistics calculations in Türkiye: TCMB daily exchange rates, container and air waybill check digits, working days with Turkish public holidays, and demurrage, detention and storage cost estimates. Each folder has its own README with usage steps, limits and what was tested. The workbooks contain formulas only. Macros are published as readable .bas code that you paste into your own workbook, because Excel on Windows blocks macros in files downloaded from the internet. Released under the MIT licence by [LogisticsIQ](https://logisticsiq.org).

---

LogisticsIQ hazırladı: https://logisticsiq.org
