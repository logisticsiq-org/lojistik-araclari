"""Türkiye için iş günü hesabı.

Resmî tatil listesi `resmi-tatiller-2026-2028.csv` dosyasından okunur.
Kapsam 2026-01-01 ile 2028-12-31 arasıdır. Bu aralığın dışındaki tarihler için
hesap yapılmaz, açık bir hata verilir.

Kurallar
- Cumartesi ve Pazar hafta sonudur. `cumartesi_tatil=False` ise yalnız Pazar tatildir.
- Tam gün resmî tatiller iş günü sayılmaz.
- Yarım gün tatiller (bayram arifeleri ve 28 Ekim öğleden sonrası) iş günü sayılır,
  çünkü sabah çalışılır. Bu günler ayrıca listelenir.
- `is_gunu_ekle` Excel'deki WORKDAY.INTL gibi çalışır: başlangıç günü sayılmaz,
  sayı negatif olabilir, sayı 0 ise başlangıç tarihi aynen döner.
- `is_gunu_say` Excel'deki NETWORKDAYS.INTL gibi iki ucu da sayar.

Komut satırı:
    python3 is_gunu.py ekle 2026-05-25 3
    python3 is_gunu.py say 2026-05-01 2026-05-31
    python3 is_gunu.py say 2026-05-01 2026-05-31 --cumartesi-calisilir
    python3 is_gunu.py tatiller 2026-03-01 2026-03-31
"""
import argparse
import csv
import os
import sys
from collections import namedtuple
from datetime import date, datetime, timedelta

SURUM_TARIHI = "2026-10-19"
KAPSAM_BASLANGIC = date(2026, 1, 1)
KAPSAM_BITIS = date(2028, 12, 31)
VARSAYILAN_CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)), "resmi-tatiller-2026-2028.csv")

GUN_ADLARI = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]

# tarih: date, ad: str, sure: "tam" veya "yarim", kaynak: str
Tatil = namedtuple("Tatil", "tarih ad sure kaynak")


class KapsamDisiHata(ValueError):
    """Tarih, tatil listesinin kapsadığı yılların dışında."""


class TarihSirasiHatasi(ValueError):
    """Son tarih, ilk tarihten önce."""


def gun_adi(tarih):
    return GUN_ADLARI[tarih.weekday()]


def tarih_yaz(tarih):
    """gg.aa.yyyy biçiminde yazar."""
    return tarih.strftime("%d.%m.%Y")


def tarih_oku(metin):
    """YYYY-MM-DD veya GG.AA.YYYY biçimindeki metni tarihe çevirir."""
    metin = metin.strip()
    for bicim in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(metin, bicim).date()
        except ValueError:
            continue
    raise ValueError(f"Tarih anlaşılamadı: '{metin}'. YYYY-AA-GG veya GG.AA.YYYY biçimini kullanın.")


def tatilleri_yukle(yol=None):
    """CSV dosyasını okur. Aynı tarihe düşen kayıtlar tek kayıtta birleştirilir.

    Örnek: 2027-05-19 hem Kurban Bayramı 4. gün hem 19 Mayıs olarak listelenir.
    Adlar " / " ile birleşir, süre "tam" ise "tam" kalır. Sonuç tarihe göre sıralıdır.
    """
    yol = yol or VARSAYILAN_CSV
    birlesik = {}
    with open(yol, encoding="utf-8", newline="") as f:
        for satir in csv.DictReader(f):
            tarih = datetime.strptime(satir["tarih"].strip(), "%Y-%m-%d").date()
            ad = satir["ad"].strip()
            sure = satir["sure"].strip()
            if sure not in ("tam", "yarim"):
                raise ValueError(f"{tarih}: süre 'tam' veya 'yarim' olmalı, bulunan '{sure}'")
            kaynak = satir["kaynak"].strip()
            if tarih in birlesik:
                onceki = birlesik[tarih]
                adlar = onceki.ad if ad in onceki.ad.split(" / ") else onceki.ad + " / " + ad
                sure = "tam" if "tam" in (onceki.sure, sure) else "yarim"
                birlesik[tarih] = Tatil(tarih, adlar, sure, onceki.kaynak)
            else:
                birlesik[tarih] = Tatil(tarih, ad, sure, kaynak)
    return [birlesik[t] for t in sorted(birlesik)]


class TatilTakvimi:
    """Tatil listesini tutar ve iş günü hesaplarını yapar."""

    def __init__(self, tatiller=None):
        self.tatiller = tatiller if tatiller is not None else tatilleri_yukle()
        self.tam_gunler = {t.tarih for t in self.tatiller if t.sure == "tam"}
        self.yarim_gunler = {t.tarih for t in self.tatiller if t.sure == "yarim"}

    def kapsam_kontrol(self, tarih, etiket="Tarih"):
        if not (KAPSAM_BASLANGIC <= tarih <= KAPSAM_BITIS):
            raise KapsamDisiHata(
                f"{etiket} ({tarih_yaz(tarih)}) tatil listesinin kapsamı dışında. "
                f"Kapsam {tarih_yaz(KAPSAM_BASLANGIC)} ile {tarih_yaz(KAPSAM_BITIS)} arasıdır. "
                "Sonuç hesaplanmadı."
            )

    def is_gunu_mu(self, tarih, cumartesi_tatil=True):
        """Hafta sonu değilse ve tam gün resmî tatil değilse True. Yarım günler iş günüdür."""
        self.kapsam_kontrol(tarih)
        if tarih.weekday() == 6:
            return False
        if tarih.weekday() == 5 and cumartesi_tatil:
            return False
        return tarih not in self.tam_gunler


def is_gunu_mu(tarih, cumartesi_tatil=True, takvim=None):
    return (takvim or varsayilan_takvim()).is_gunu_mu(tarih, cumartesi_tatil)


_VARSAYILAN = {}


def varsayilan_takvim():
    if "takvim" not in _VARSAYILAN:
        _VARSAYILAN["takvim"] = TatilTakvimi()
    return _VARSAYILAN["takvim"]


def is_gunu_ekle(baslangic, n, cumartesi_tatil=True, takvim=None):
    """Başlangıç tarihine n iş günü ekler (WORKDAY.INTL gibi).

    Başlangıç günü sayılmaz. n negatifse geriye gider. n 0 ise başlangıç tarihi döner.
    Başlangıç ya da sonuç kapsam dışındaysa KapsamDisiHata verir.
    """
    takvim = takvim or varsayilan_takvim()
    if not isinstance(n, int) or isinstance(n, bool):
        raise ValueError("İş günü sayısı tam sayı olmalı.")
    takvim.kapsam_kontrol(baslangic, "Başlangıç tarihi")
    adim = 1 if n >= 0 else -1
    kalan = abs(n)
    gun = baslangic
    while kalan > 0:
        gun = gun + timedelta(days=adim)
        if not (KAPSAM_BASLANGIC <= gun <= KAPSAM_BITIS):
            raise KapsamDisiHata(
                f"Sonuç tatil listesinin kapsamı dışına çıkıyor "
                f"({tarih_yaz(KAPSAM_BASLANGIC)} ile {tarih_yaz(KAPSAM_BITIS)}). Sonuç hesaplanmadı."
            )
        if takvim.is_gunu_mu(gun, cumartesi_tatil):
            kalan -= 1
    return gun


def _aralik_kontrol(bas, bit, takvim):
    takvim.kapsam_kontrol(bas, "İlk tarih")
    takvim.kapsam_kontrol(bit, "Son tarih")
    if bit < bas:
        raise TarihSirasiHatasi("Son tarih ilk tarihten önce olamaz.")


def is_gunu_say(bas, bit, cumartesi_tatil=True, takvim=None):
    """İki tarih arasındaki iş günü sayısı, iki uç dahil (NETWORKDAYS.INTL gibi)."""
    takvim = takvim or varsayilan_takvim()
    _aralik_kontrol(bas, bit, takvim)
    sayi = 0
    gun = bas
    while gun <= bit:
        if takvim.is_gunu_mu(gun, cumartesi_tatil):
            sayi += 1
        gun += timedelta(days=1)
    return sayi


def aradaki_tatiller(bas, bit, takvim=None):
    """Aralıktaki tüm resmî tatilleri (tam ve yarım gün) tarih sırasıyla listeler.

    Tarihler sırasız verilirse yer değiştirilir. Hafta sonuna denk gelen tatiller de listelenir.
    """
    takvim = takvim or varsayilan_takvim()
    if bit < bas:
        bas, bit = bit, bas
    takvim.kapsam_kontrol(bas, "İlk tarih")
    takvim.kapsam_kontrol(bit, "Son tarih")
    return [t for t in takvim.tatiller if bas <= t.tarih <= bit]


def takvim_gunu(bas, bit):
    """İki ucu dahil takvim günü sayısı."""
    if bit < bas:
        raise TarihSirasiHatasi("Son tarih ilk tarihten önce olamaz.")
    return (bit - bas).days + 1


def ozet(bas, bit, cumartesi_tatil=True, takvim=None):
    """Excel dosyasındaki ikinci blokla aynı alanları sözlük olarak verir."""
    takvim = takvim or varsayilan_takvim()
    _aralik_kontrol(bas, bit, takvim)
    tatiller = aradaki_tatiller(bas, bit, takvim)
    tam = [t for t in tatiller if t.sure == "tam"]
    yarim = [t for t in tatiller if t.sure == "yarim"]
    # Hafta sonu zaten tatil olduğundan yalnız iş günü olacak günlerde düşen tam gün tatiller sayılır.
    dusen = 0
    gun = bas
    while gun <= bit:
        hafta_ici = gun.weekday() < 5 or (gun.weekday() == 5 and not cumartesi_tatil)
        if hafta_ici and gun in takvim.tam_gunler:
            dusen += 1
        gun += timedelta(days=1)
    return {
        "is_gunu": is_gunu_say(bas, bit, cumartesi_tatil, takvim),
        "takvim_gunu": takvim_gunu(bas, bit),
        "tam_gun_tatil": len(tam),
        "yarim_gun_tatil": len(yarim),
        "tatil_nedeniyle_dusen_is_gunu": dusen,
    }


# ---------------------------------------------------------------- komut satırı

def _tatil_satiri(t):
    sure = "tam gün" if t.sure == "tam" else "yarım gün, iş günü sayılır"
    return f"  {tarih_yaz(t.tarih)} {gun_adi(t.tarih)}: {t.ad} ({sure})"


def _tatil_yaz(tatiller):
    if not tatiller:
        print("Aradaki resmî tatiller: yok")
        return
    print("Aradaki resmî tatiller:")
    for t in tatiller:
        print(_tatil_satiri(t))


def _arguman_ayirici():
    p = argparse.ArgumentParser(
        prog="is_gunu.py",
        description="Türkiye için iş günü hesabı (2026 ile 2028 arası). Tam gün resmî tatiller "
                    "iş günü sayılmaz, yarım gün tatiller iş günü sayılır ve ayrıca listelenir.",
    )
    alt = p.add_subparsers(dest="komut", required=True)

    ortak = argparse.ArgumentParser(add_help=False)
    ortak.add_argument("--cumartesi-calisilir", action="store_true",
                       help="Cumartesi iş günü sayılır, yalnız Pazar tatildir.")

    e = alt.add_parser("ekle", parents=[ortak], help="Bir tarihe iş günü ekler veya çıkarır.",
                       description="Başlangıç gününü saymadan n iş günü sonrasını (negatifse öncesini) bulur.")
    e.add_argument("baslangic", help="Başlangıç tarihi, YYYY-AA-GG veya GG.AA.YYYY")
    e.add_argument("n", type=int, help="Eklenecek iş günü sayısı, negatif olabilir")

    s = alt.add_parser("say", parents=[ortak], help="İki tarih arasındaki iş günlerini sayar.",
                       description="İki ucu da dahil ederek iş günü sayısını verir.")
    s.add_argument("ilk", help="İlk tarih")
    s.add_argument("son", help="Son tarih")

    t = alt.add_parser("tatiller", help="Aralıktaki resmî tatilleri listeler.")
    t.add_argument("ilk", help="İlk tarih")
    t.add_argument("son", help="Son tarih")
    return p


def ana(argv=None):
    args = _arguman_ayirici().parse_args(argv)
    try:
        if args.komut == "ekle":
            cumartesi_tatil = not args.cumartesi_calisilir
            bas = tarih_oku(args.baslangic)
            sonuc = is_gunu_ekle(bas, args.n, cumartesi_tatil)
            print(f"Sonuç: {tarih_yaz(sonuc)} {gun_adi(sonuc)}")
            araliktaki = aradaki_tatiller(min(bas, sonuc), max(bas, sonuc))
            _tatil_yaz(araliktaki)
        elif args.komut == "say":
            cumartesi_tatil = not args.cumartesi_calisilir
            bas, bit = tarih_oku(args.ilk), tarih_oku(args.son)
            o = ozet(bas, bit, cumartesi_tatil)
            print(f"İş günü sayısı (iki uç dahil): {o['is_gunu']}")
            print(f"Takvim günü (iki uç dahil): {o['takvim_gunu']}")
            print(f"Tam gün resmî tatil: {o['tam_gun_tatil']}")
            print(f"Bunlardan iş günü sayısını düşüren: {o['tatil_nedeniyle_dusen_is_gunu']}")
            print(f"Yarım gün tatil: {o['yarim_gun_tatil']} (iş günü olarak sayıldı)")
            _tatil_yaz(aradaki_tatiller(bas, bit))
        else:
            _tatil_yaz(aradaki_tatiller(tarih_oku(args.ilk), tarih_oku(args.son)))
    except (ValueError, OSError) as hata:
        print(f"Hata: {hata}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(ana())
