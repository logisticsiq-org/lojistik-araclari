#!/usr/bin/env python3
"""Konteyner numarası (ISO 6346) ve hava konşimentosu (AWB) kontrol hanesi denetimi.

Geçerli bir kontrol hanesi, numaranın biçimce doğru olduğunu gösterir. Konteynerin
gerçekten var olduğunu ya da o sahibe ait olduğunu göstermez.

Örnekler:
    python3 kontrol_hanesi.py CSQU3054383 020-12345675
    python3 kontrol_hanesi.py --csv liste.csv --sutun 1 --cikti sonuc.csv

Lisans: MIT. LogisticsIQ hazırladı: https://logisticsiq.org
"""
import argparse
import csv
import re
import sys

GECERLI = "gecerli"
HANE_HATALI = "hane_hatali"
BICIM_HATALI = "bicim_hatali"

MESAJLAR = {
    GECERLI: "Geçerli",
    BICIM_HATALI: "Biçim hatalı",
}


def _harf_degerleri():
    """ISO 6346: A=10'dan başlar, 11'in katları (11, 22, 33) atlanır."""
    degerler = {}
    deger = 10
    for kod in range(ord("A"), ord("Z") + 1):
        if deger % 11 == 0:
            deger += 1
        degerler[chr(kod)] = deger
        deger += 1
    return degerler


HARF_DEGERLERI = _harf_degerleri()

_KONTEYNER_BICIM = re.compile(r"[A-Z]{3}[UJZ][0-9]{7}")
_KONTEYNER_10 = re.compile(r"[A-Z]{3}[UJZ][0-9]{6}")
_AWB_BICIM = re.compile(r"[0-9]{11}")


def normalle(no):
    """Büyük harfe çevirir, boşluk, tire, nokta ve eğik çizgiyi siler.

    Büyük harf dönüşümü yalnızca a-z harflerine uygulanır (Türkçe i/ı sorunlarını önlemek için).
    """
    metin = "" if no is None else str(no)
    for karakter in (" ", "-", ".", "/"):
        metin = metin.replace(karakter, "")
    return "".join(k.upper() if "a" <= k <= "z" else k for k in metin)


def konteyner_kontrol_hanesi(no10):
    """İlk 10 karakter (3 harf, U/J/Z ve 6 rakam) için kontrol hanesini (0-9) hesaplar."""
    metin = normalle(no10)
    if not _KONTEYNER_10.fullmatch(metin):
        raise ValueError("Konteyner numarasının ilk 10 karakteri 3 harf, U/J/Z ve 6 rakam olmalı: %r" % no10)
    toplam = 0
    for sira, karakter in enumerate(metin):
        deger = HARF_DEGERLERI[karakter] if karakter.isalpha() else int(karakter)
        toplam += deger * (2 ** sira)
    return (toplam % 11) % 10


def konteyner_kontrol(no):
    """(durum, beklenen_hane) döndürür. Biçim hatalıysa beklenen_hane None olur."""
    metin = normalle(no)
    if not _KONTEYNER_BICIM.fullmatch(metin):
        return BICIM_HATALI, None
    beklenen = konteyner_kontrol_hanesi(metin[:10])
    if int(metin[10]) == beklenen:
        return GECERLI, beklenen
    return HANE_HATALI, beklenen


def awb_kontrol_hanesi(seri7):
    """Seri numarasının ilk 7 hanesi için kontrol hanesini (mod 7) hesaplar."""
    metin = normalle(seri7)
    if not re.fullmatch(r"[0-9]{7}", metin):
        raise ValueError("AWB seri numarasının ilk 7 hanesi rakam olmalı: %r" % seri7)
    return int(metin) % 7


def awb_kontrol(no):
    """(durum, beklenen_hane) döndürür. 11 rakam: 3 haneli önek ve 8 haneli seri numarası."""
    metin = normalle(no)
    if not _AWB_BICIM.fullmatch(metin):
        return BICIM_HATALI, None
    beklenen = awb_kontrol_hanesi(metin[3:10])
    if int(metin[10]) == beklenen:
        return GECERLI, beklenen
    return HANE_HATALI, beklenen


def sonuc_metni(durum, beklenen):
    """Excel dosyasındaki sonuç metniyle aynı metni üretir."""
    if durum == HANE_HATALI:
        return "Kontrol hanesi hatalı (doğrusu: %d)" % beklenen
    return MESAJLAR[durum]


def tur_tahmin_et(no):
    """Normalleştirilmiş numara harfle başlıyorsa konteyner, aksi halde AWB sayılır."""
    metin = normalle(no)
    return "konteyner" if metin[:1].isalpha() else "awb"


def kontrol_et(no, tur="otomatik"):
    """(tur, normalleştirilmiş no, durum, beklenen_hane) döndürür."""
    if tur == "otomatik":
        tur = tur_tahmin_et(no)
    islev = konteyner_kontrol if tur == "konteyner" else awb_kontrol
    durum, beklenen = islev(no)
    return tur, normalle(no), durum, beklenen


# ---------------------------------------------------------------- komut satırı

def csv_isle(girdi, cikti, sutun, tur, baslik, ayirac):
    """CSV'nin her satırına tur, normalle_no, durum ve beklenen_hane sütunlarını ekler."""
    okuyucu = csv.reader(girdi, delimiter=ayirac)
    yazici = csv.writer(cikti, delimiter=ayirac, lineterminator="\n")
    sayac = 0
    ilk = True
    for satir in okuyucu:
        if ilk and baslik:
            yazici.writerow(satir + ["tur", "normalle_no", "durum", "beklenen_hane"])
            ilk = False
            continue
        ilk = False
        if not satir:
            continue
        deger = satir[sutun - 1] if len(satir) >= sutun else ""
        if not deger.strip():
            yazici.writerow(satir + ["", "", "", ""])
            continue
        tur_adi, temiz, durum, beklenen = kontrol_et(deger, tur)
        yazici.writerow(satir + [tur_adi, temiz, durum, "" if beklenen is None else beklenen])
        sayac += 1
    return sayac


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="kontrol_hanesi.py",
        description="Konteyner numarası (ISO 6346) ve hava konşimentosu (AWB) kontrol hanesini denetler. "
                    "Geçerli bir kontrol hanesi yalnızca numaranın biçimce doğru olduğunu gösterir.",
        epilog="Örnekler:\n"
               "  python3 kontrol_hanesi.py CSQU3054383 020-12345675\n"
               "  python3 kontrol_hanesi.py --csv liste.csv --sutun 1 --cikti sonuc.csv --baslik",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("numaralar", nargs="*", help="Denetlenecek numaralar.")
    ap.add_argument("--tur", choices=["otomatik", "konteyner", "awb"], default="otomatik",
                    help="Numara türü. Varsayılan: otomatik (harfle başlıyorsa konteyner, rakamla başlıyorsa AWB).")
    ap.add_argument("--csv", metavar="GIRDI", help="Numaraların bulunduğu CSV dosyası.")
    ap.add_argument("--sutun", type=int, default=1, help="CSV'de numaranın sütunu (1'den başlar). Varsayılan: 1.")
    ap.add_argument("--cikti", metavar="CIKTI", help="Sonuç sütunları eklenmiş CSV dosyası.")
    ap.add_argument("--baslik", action="store_true", help="CSV'nin ilk satırı başlıktır.")
    ap.add_argument("--ayirac", default=",", help="CSV sütun ayırıcısı. Varsayılan: virgül.")
    args = ap.parse_args(argv)

    if args.csv:
        if not args.cikti:
            ap.error("--csv ile birlikte --cikti da verilmelidir.")
        if args.sutun < 1:
            ap.error("--sutun 1 ya da daha büyük olmalı.")
        with open(args.csv, encoding="utf-8-sig", newline="") as g, \
                open(args.cikti, "w", encoding="utf-8", newline="") as c:
            sayac = csv_isle(g, c, args.sutun, args.tur, args.baslik, args.ayirac)
        print("%d numara denetlendi: %s" % (sayac, args.cikti))
        return 0

    if not args.numaralar:
        ap.error("Numara verin ya da --csv kullanın.")
    hepsi_gecerli = True
    for no in args.numaralar:
        tur_adi, temiz, durum, beklenen = kontrol_et(no, args.tur)
        print("%-14s %-10s %s" % (temiz or "(boş)", tur_adi, sonuc_metni(durum, beklenen)))
        hepsi_gecerli = hepsi_gecerli and durum == GECERLI
    return 0 if hepsi_gecerli else 1


if __name__ == "__main__":
    sys.exit(main())
