#!/usr/bin/env python3
"""TCMB günlük gösterge niteliğindeki döviz kurlarını getirir.

Kaynak: https://www.tcmb.gov.tr/kurlar/YYYYMM/DDMMYYYY.xml
Yalnızca Python standart kütüphanesini kullanır (Python 3.9 ve üzeri).

Örnekler:
    python3 tcmb_kur.py 2026-10-16 USD EUR
    python3 tcmb_kur.py 2026-10-16 JPY --tur doviz-alis
    python3 tcmb_kur.py --bas 2026-10-01 --bit 2026-10-16 USD EUR --csv kurlar.csv

Lisans: MIT. LogisticsIQ hazırladı: https://logisticsiq.org
"""
import argparse
import csv
import datetime
import sys
import time
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation

KULLANICI_ADI = "lojistik-araclari (LogisticsIQ)"
ZAMAN_ASIMI = 15  # saniye
EN_COK_BAYT = 2000000  # Bir bülten yaklaşık 10 KB'dır, bu üst sınır yalnızca güvenlik içindir.
AG_BEKLEME = 0.2  # Gerçek ağ istekleri arasında saniye cinsinden kısa bekleme.
EN_COK_GUN = 366  # --bas ve --bit arasında izin verilen en çok gün sayısı.

TUR_ADLARI = ("ForexBuying", "ForexSelling", "BanknoteBuying", "BanknoteSelling")
TUR_TAKMA_ADLARI = {
    "doviz-alis": "ForexBuying",
    "doviz-satis": "ForexSelling",
    "efektif-alis": "BanknoteBuying",
    "efektif-satis": "BanknoteSelling",
}

_ONBELLEK = {}  # Bir çalıştırma boyunca url -> bayt (ya da bülten yoksa None).


class TCMBHata(Exception):
    """Kullanıcıya Türkçe gösterilecek hata."""


class BultenBulunamadi(TCMBHata):
    """Sunucu 404 döndürdü: hafta sonu, resmi tatil ya da bülten henüz yok."""


def bulten_url(tarih):
    """Verilen günün bülten adresini döndürür."""
    return "https://www.tcmb.gov.tr/kurlar/%04d%02d/%02d%02d%04d.xml" % (
        tarih.year, tarih.month, tarih.day, tarih.month, tarih.year)


def tur_coz(tur):
    """Takma adı ya da XML alan adını XML alan adına çevirir."""
    metin = str(tur).strip()
    for ad in TUR_ADLARI:
        if metin.lower() == ad.lower():
            return ad
    anahtar = metin.lower().replace("_", "-")
    if anahtar in TUR_TAKMA_ADLARI:
        return TUR_TAKMA_ADLARI[anahtar]
    gecerli = ", ".join(sorted(TUR_TAKMA_ADLARI) + list(TUR_ADLARI))
    raise TCMBHata("Kur türü anlaşılamadı: '%s'. Geçerli türler: %s." % (tur, gecerli))


def _ondalik(metin):
    if metin is None or not metin.strip():
        return None
    try:
        return Decimal(metin.strip())
    except InvalidOperation:
        raise TCMBHata("Bültende sayı olarak okunamayan bir değer var: '%s'." % metin.strip())


def _tarih_oku(metin):
    try:
        gun, ay, yil = metin.strip().split(".")
        return datetime.date(int(yil), int(ay), int(gun))
    except (AttributeError, ValueError):
        raise TCMBHata("Bülten tarihi okunamadı: '%s'." % metin)


def bulten_coz(xml_bayt):
    """Bülten XML'ini çözer.

    Döndürür: (bülten tarihi, bülten numarası, kurlar).
    kurlar: kod -> {"birim": int, "isim": str, "alanlar": {alan adı: Decimal ya da None}}
    Boş alanlar None olarak gelir.
    """
    try:
        kok = ET.fromstring(xml_bayt)
    except ET.ParseError as hata:
        raise TCMBHata("Bülten XML olarak okunamadı: %s" % hata)
    if kok.tag != "Tarih_Date":
        raise TCMBHata("Beklenen bülten biçimi bulunamadı (kök öğe: %s)." % kok.tag)
    tarih = _tarih_oku(kok.get("Tarih"))
    bulten_no = (kok.get("Bulten_No") or "").strip()
    kurlar = {}
    for oge in kok.findall("Currency"):
        kod = (oge.get("CurrencyCode") or oge.get("Kod") or "").strip().upper()
        if not kod:
            continue
        birim_metin = (oge.findtext("Unit") or "").strip()
        try:
            birim = int(birim_metin)
        except ValueError:
            birim = 0
        kurlar[kod] = {
            "birim": birim,
            "isim": (oge.findtext("Isim") or "").strip(),
            "alanlar": {ad: _ondalik(oge.findtext(ad)) for ad in TUR_ADLARI},
        }
    if not kurlar:
        raise TCMBHata("Bültende hiç döviz satırı yok.")
    return tarih, bulten_no, kurlar


def varsayilan_opener(url):
    """Adresi indirip baytları döndürür. 404 için BultenBulunamadi fırlatır."""
    time.sleep(AG_BEKLEME)
    istek = urllib.request.Request(url, headers={"User-Agent": KULLANICI_ADI})
    try:
        with urllib.request.urlopen(istek, timeout=ZAMAN_ASIMI) as yanit:
            return yanit.read(EN_COK_BAYT)
    except urllib.error.HTTPError as hata:
        if hata.code == 404:
            raise BultenBulunamadi("Bülten yok: %s" % url)
        raise TCMBHata("TCMB sunucusu HTTP %d döndürdü: %s" % (hata.code, url))
    except (urllib.error.URLError, OSError) as hata:
        raise TCMBHata("TCMB sunucusuna ulaşılamadı (%s): %s" % (hata, url))


def _bayt_getir(url, opener, onbellek):
    if url in onbellek:
        return onbellek[url]
    try:
        veri = opener(url)
    except BultenBulunamadi:
        veri = None
    except urllib.error.HTTPError as hata:
        if hata.code != 404:
            raise TCMBHata("TCMB sunucusu HTTP %d döndürdü: %s" % (hata.code, url))
        veri = None
    onbellek[url] = veri
    return veri


def _gun(tarih):
    return "%02d.%02d.%04d" % (tarih.day, tarih.month, tarih.year)


def kur_getir(tarih, doviz, tur="ForexSelling", geri_gun=10, opener=None, onbellek=None):
    """TL cinsinden 1 birim dövizin kurunu getirir.

    Döndürür: (kur: Decimal, kullanılan bülten tarihi: date, bülten numarası: str).
    Verilen günün bülteni yoksa (404) her seferinde bir gün geriye gidilir.
    opener: adresi alıp bayt döndüren işlev. Bülten yoksa BultenBulunamadi fırlatmalıdır.
    """
    if isinstance(tarih, datetime.datetime):
        tarih = tarih.date()
    alan = tur_coz(tur)
    kod = str(doviz).strip().upper()
    if not kod:
        raise TCMBHata("Döviz kodu boş olamaz.")
    if opener is None:
        opener = varsayilan_opener
        if onbellek is None:
            onbellek = _ONBELLEK
    elif onbellek is None:
        onbellek = {}

    for geri in range(int(geri_gun) + 1):
        gun = tarih - datetime.timedelta(days=geri)
        veri = _bayt_getir(bulten_url(gun), opener, onbellek)
        if veri is None:
            continue
        bulten_tarihi, bulten_no, kurlar = bulten_coz(veri)
        if kod not in kurlar:
            raise TCMBHata("%s kodu %s tarihli bültende yok. Bültendeki kodlar: %s." % (
                kod, _gun(bulten_tarihi), ", ".join(sorted(kurlar))))
        satir = kurlar[kod]
        deger = satir["alanlar"][alan]
        if deger is None:
            raise TCMBHata("%s için %s alanı %s tarihli bültende boş." % (kod, alan, _gun(bulten_tarihi)))
        if satir["birim"] <= 0:
            raise TCMBHata("%s için birim (Unit) alanı %s tarihli bültende okunamadı." % (kod, _gun(bulten_tarihi)))
        return deger / Decimal(satir["birim"]), bulten_tarihi, bulten_no

    ilk = _gun(tarih)
    son = _gun(tarih - datetime.timedelta(days=int(geri_gun)))
    raise TCMBHata("%s ile %s arasında (%d gün geriye kadar) TCMB bülteni bulunamadı." % (
        son, ilk, int(geri_gun)))


# ---------------------------------------------------------------- komut satırı

def _tarih_arguman(metin):
    for bicim in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.datetime.strptime(metin, bicim).date()
        except ValueError:
            pass
    raise argparse.ArgumentTypeError("Tarih 2026-10-16 ya da 16.10.2026 biçiminde olmalı: '%s'" % metin)


def _gunler(bas, bit):
    gun = bas
    while gun <= bit:
        yield gun
        gun += datetime.timedelta(days=1)


def satirlari_uret(gunler, dovizler, tur, geri_gun, opener=None, onbellek=None):
    """Her istenen gün ve döviz için bir sözlük üretir (CSV sütunlarıyla aynı adlar)."""
    if opener is None and onbellek is None:
        onbellek = _ONBELLEK
    elif onbellek is None:
        onbellek = {}
    alan = tur_coz(tur)
    for gun in gunler:
        for doviz in dovizler:
            kur, bulten_tarihi, bulten_no = kur_getir(
                gun, doviz, alan, geri_gun=geri_gun, opener=opener, onbellek=onbellek)
            yield {
                "istenen_tarih": gun.isoformat(),
                "bulten_tarihi": bulten_tarihi.isoformat(),
                "bulten_no": bulten_no,
                "doviz": str(doviz).strip().upper(),
                "tur": alan,
                "kur": format(kur, "f"),
            }


SUTUNLAR = ["istenen_tarih", "bulten_tarihi", "bulten_no", "doviz", "tur", "kur"]


def csv_yaz(satirlar, dosya_yolu, ayirac=","):
    sayac = 0
    with open(dosya_yolu, "w", encoding="utf-8", newline="") as f:
        yazici = csv.DictWriter(f, fieldnames=SUTUNLAR, delimiter=ayirac)
        yazici.writeheader()
        for satir in satirlar:
            yazici.writerow(satir)
            sayac += 1
    return sayac


def tablo_yaz(satirlar, cikti=None):
    cikti = cikti or sys.stdout
    basliklar = ["İstenen tarih", "Bülten tarihi", "Bülten no", "Döviz", "Tür", "Kur (TL)"]
    veri = [[s["istenen_tarih"], s["bulten_tarihi"], s["bulten_no"], s["doviz"], s["tur"], s["kur"]]
            for s in satirlar]
    genislik = [max(len(basliklar[i]), *(len(v[i]) for v in veri)) if veri else len(basliklar[i])
                for i in range(len(basliklar))]
    def bicimle(hucreler):
        return "  ".join(h.ljust(genislik[i]) for i, h in enumerate(hucreler)).rstrip()
    print(bicimle(basliklar), file=cikti)
    print(bicimle(["-" * g for g in genislik]), file=cikti)
    for v in veri:
        print(bicimle(v), file=cikti)


def arguman_ayristirici():
    ap = argparse.ArgumentParser(
        prog="tcmb_kur.py",
        description="TCMB günlük gösterge niteliğindeki döviz kurlarını getirir. "
                    "Kur, 1 birim döviz için TL karşılığıdır (JPY gibi 100 birimlik kurlar 100'e bölünür). "
                    "İstenen günün bülteni yoksa (hafta sonu, resmi tatil) önceki günlere gidilir.",
        epilog="Örnekler:\n"
               "  python3 tcmb_kur.py 2026-10-16 USD EUR\n"
               "  python3 tcmb_kur.py 2026-10-16 USD --tur efektif-satis\n"
               "  python3 tcmb_kur.py --bas 2026-10-01 --bit 2026-10-16 USD EUR --csv kurlar.csv",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("argumanlar", nargs="+", metavar="TARIH_VE_DOVIZ",
                    help="Tek gün için: tarih ve ardından bir ya da daha çok döviz kodu "
                         "(2026-10-16 USD EUR). --bas ve --bit kullanılırsa yalnızca döviz kodları.")
    ap.add_argument("--tur", default="doviz-satis",
                    help="doviz-alis, doviz-satis, efektif-alis ya da efektif-satis "
                         "(ForexBuying vb. de olur). Varsayılan: doviz-satis.")
    ap.add_argument("--bas", type=_tarih_arguman, help="Aralığın ilk günü (2026-10-01).")
    ap.add_argument("--bit", type=_tarih_arguman, help="Aralığın son günü (2026-10-16).")
    ap.add_argument("--csv", metavar="DOSYA", help="Sonucu bu CSV dosyasına yazar (UTF-8).")
    ap.add_argument("--ayirac", default=",", help="CSV sütun ayırıcısı. Varsayılan: virgül.")
    ap.add_argument("--geri-gun", type=int, default=10,
                    help="Bülten yoksa en çok kaç gün geriye gidileceği. Varsayılan: 10.")
    return ap


def main(argv=None, opener=None):
    ap = arguman_ayristirici()
    args = ap.parse_args(argv)
    aralik = args.bas is not None or args.bit is not None
    try:
        if aralik:
            if args.bas is None or args.bit is None:
                ap.error("Aralık için --bas ve --bit birlikte verilmelidir.")
            if args.bit < args.bas:
                ap.error("--bit tarihi --bas tarihinden önce olamaz.")
            if (args.bit - args.bas).days + 1 > EN_COK_GUN:
                ap.error("Aralık en çok %d gün olabilir." % EN_COK_GUN)
            gunler = list(_gunler(args.bas, args.bit))
            dovizler = args.argumanlar
        else:
            if len(args.argumanlar) < 2:
                ap.error("Tarih ve en az bir döviz kodu verin. Örnek: tcmb_kur.py 2026-10-16 USD")
            gunler = [_tarih_arguman(args.argumanlar[0])]
            dovizler = args.argumanlar[1:]
        satirlar = list(satirlari_uret(gunler, dovizler, args.tur, args.geri_gun, opener=opener))
    except TCMBHata as hata:
        print("Hata: %s" % hata, file=sys.stderr)
        return 1
    except argparse.ArgumentTypeError as hata:
        ap.error(str(hata))
    if args.csv:
        sayac = csv_yaz(satirlar, args.csv, args.ayirac)
        print("%d satır yazıldı: %s" % (sayac, args.csv))
    else:
        tablo_yaz(satirlar)
    return 0


if __name__ == "__main__":
    sys.exit(main())
