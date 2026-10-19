"""Demuraj, detention ve ardiye maliyeti için tahmini hesap.

Tanımlar
- Demuraj: Dolu konteynerin tahliyeden terminal çıkışına kadar terminalde kaldığı günler.
- Detention: Konteynerin terminalden çıkışından boş iadeye kadar terminal dışında geçen günler.
- Birleşik süre: Bazı taşıyıcılar tahliyeden boş iadeye kadar tek bir serbest süre uygular.
- Ardiye: Terminalin kendi tarifesiyle, demurajla aynı dönem için kestiği depolama bedeli.

Tüm sayımlar takvim günüdür. Hafta sonu ve tatiller dahildir. Tarifeler taşıyıcıya, limana ve
sözleşmeye göre değişir. Sonuç bir tahmindir.

Tarife satırı: (başlangıç günü, bitiş günü, günlük ücret). Bitiş günü boşsa (None) satır sonuna kadar sürer.
Serbest süre, ücreti 0 olan bir satırdır. Satırda ücretli gün sayısı:
    MAX(0, MIN(toplam_gün, bitiş_ya_da_toplam) - başlangıç + 1)
Tutar = ücretli gün x günlük ücret x konteyner sayısı. Satırların çakışmaması kullanıcının sorumluluğundadır.

Komut satırı:
    python3 demuraj.py ornek.json
"""
import argparse
import json
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Dict, List, Optional, Tuple

SURUM_TARIHI = "2026-10-19"
MOD_AYRI = "ayri"
MOD_BIRLESIK = "birlesik"
# Excel dosyasındaki açılır liste metinleri
MOD_METINLERI = {
    "Ayrı (demuraj ve detention ayrı)": MOD_AYRI,
    "Birleşik (tek serbest süre)": MOD_BIRLESIK,
}
PARA_BIRIMLERI = ("USD", "EUR", "TRY")
TABLOLAR = ("demuraj", "detention", "birlesik", "ardiye")
SATIR_SAYISI = 5
UYARI_YOK = "Uyarı yok."

M_SAYI = "Konteyner sayısı 1 veya daha büyük bir tam sayı olmalı."
M_PARA = "Para birimi USD, EUR veya TRY olmalı."
M_MOD = "Hesap şekli listeden seçilmeli."
M_ILK_GUN = "İlk gün sayılır mı alanı Evet veya Hayır olmalı."
M_TAHLIYE = "Tahliye tarihini girin."
M_CIKIS = "Çıkış tarihini girin."
M_IADE = "Boş iade tarihini girin."
M_ARDIYE_CIKIS = "Ardiye tarifesi dolu olduğu için çıkış tarihini girin."
M_SIRA_CIKIS = "Çıkış tarihi tahliye tarihinden önce olamaz."
M_SIRA_IADE_TAHLIYE = "Boş iade tarihi tahliye tarihinden önce olamaz."
M_SIRA_IADE_CIKIS = "Boş iade tarihi çıkış tarihinden önce olamaz."
M_KUR = "Kur 0'dan büyük bir sayı olmalı."
M_TARIFE = "Tarife tablolarında hatalı satır var. Satır notlarına bakın."
M_COK_SATIR = "Her tarife tablosunda en fazla 5 satır olabilir."

R_BASLANGIC = "Başlangıç günü 1 veya daha büyük bir tam sayı olmalı."
R_UCRET_YOK = "Günlük ücret girilmeli."
R_UCRET_EKSI = "Günlük ücret eksi olamaz."
R_BITIS_SAYI = "Bitiş günü bir tam sayı olmalı."
R_BITIS_KUCUK = "Bitiş günü başlangıç gününden küçük olamaz."

Satir = Tuple[Optional[float], Optional[float], Optional[float]]


class GirdiHatasi(ValueError):
    """Girdi eksik, sırasız veya hatalı."""


@dataclass
class Girdi:
    konteyner_sayisi: int = 1
    konteyner_tipi: str = ""
    para_birimi: str = "USD"
    mod: str = MOD_AYRI
    ilk_gun_sayilir: bool = True
    kur_try: Optional[float] = None
    tahliye: Optional[date] = None
    cikis: Optional[date] = None
    bos_iade: Optional[date] = None
    demuraj: List[Satir] = field(default_factory=list)
    detention: List[Satir] = field(default_factory=list)
    birlesik: List[Satir] = field(default_factory=list)
    ardiye: List[Satir] = field(default_factory=list)


@dataclass
class Sonuc:
    demuraj_gun: Optional[int]
    detention_gun: Optional[int]
    birlesik_gun: Optional[int]
    ardiye_gun: Optional[int]
    satirlar: Dict[str, List[Optional[Tuple[int, float]]]]  # tablo -> satır başına (ücretli gün, tutar) ya da None
    ara_toplam: Dict[str, Optional[float]]
    toplam: float
    toplam_try: Optional[float]
    cumle: str


def _tam_sayi_mi(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool) and x == int(x)


def _sayi_mi(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def satir_notu(satir):
    """Dolu bir satır için sorunu anlatan metni, sorun yoksa boş metni verir. Boş satır için boş metin."""
    bas, bit, ucret = satir
    if bas is None:
        return ""
    if not _sayi_mi(bas) or bas < 1 or bas != int(bas):
        return R_BASLANGIC
    if not _sayi_mi(ucret):
        return R_UCRET_YOK
    if ucret < 0:
        return R_UCRET_EKSI
    if bit is not None:
        if not _tam_sayi_mi(bit):
            return R_BITIS_SAYI
        if bit < bas:
            return R_BITIS_KUCUK
    return ""


def _satirlari_duzenle(satirlar):
    """Satırları (başlangıç, bitiş, ücret) üçlülerine çevirir ve boş satırlarla SATIR_SAYISI'na tamamlar."""
    sonuc = []
    for s in (satirlar or [])[:SATIR_SAYISI]:
        s = tuple(s) + (None,) * (3 - len(s))
        sonuc.append(s[:3])
    while len(sonuc) < SATIR_SAYISI:
        sonuc.append((None, None, None))
    return sonuc


def ardiye_kullaniliyor(g):
    return any(s[0] is not None for s in _satirlari_duzenle(g.ardiye))


def ilk_hata(g):
    """Girdideki ilk sorunu Excel dosyasındaki Uyarı hücresiyle aynı sırayla ve metinle verir."""
    if not _tam_sayi_mi(g.konteyner_sayisi) or g.konteyner_sayisi < 1:
        return M_SAYI
    if g.para_birimi not in PARA_BIRIMLERI:
        return M_PARA
    if g.mod not in (MOD_AYRI, MOD_BIRLESIK):
        return M_MOD
    if not isinstance(g.ilk_gun_sayilir, bool):
        return M_ILK_GUN
    if g.tahliye is None:
        return M_TAHLIYE
    if g.mod == MOD_AYRI:
        if g.cikis is None:
            return M_CIKIS
        if g.bos_iade is None:
            return M_IADE
    else:
        if g.bos_iade is None:
            return M_IADE
        if g.cikis is None and ardiye_kullaniliyor(g):
            return M_ARDIYE_CIKIS
    if g.cikis is not None and g.cikis < g.tahliye:
        return M_SIRA_CIKIS
    if g.bos_iade < g.tahliye:
        return M_SIRA_IADE_TAHLIYE
    if g.cikis is not None and g.bos_iade < g.cikis:
        return M_SIRA_IADE_CIKIS
    if g.para_birimi != "TRY" and g.kur_try is not None:
        if not _sayi_mi(g.kur_try) or g.kur_try <= 0:
            return M_KUR
    for tablo in TABLOLAR:
        if len(getattr(g, tablo) or []) > SATIR_SAYISI:
            return M_COK_SATIR
    uygulanan = ("demuraj", "detention", "ardiye") if g.mod == MOD_AYRI else ("birlesik", "ardiye")
    for tablo in uygulanan:
        for s in _satirlari_duzenle(getattr(g, tablo)):
            if satir_notu(s):
                return M_TARIFE
    return None


def gun_say(bas, bit, ilk_gun_sayilir):
    return (bit - bas).days + (1 if ilk_gun_sayilir else 0)


def tablo_hesapla(satirlar, toplam_gun, konteyner_sayisi):
    """Her satır için (ücretli gün, tutar) ya da boş satır için None verir. Ara toplamı da döndürür."""
    sonuc = []
    ara = 0.0
    for bas, bit, ucret in _satirlari_duzenle(satirlar):
        if bas is None:
            sonuc.append(None)
            continue
        ust = toplam_gun if bit is None else min(toplam_gun, bit)
        gun = max(0, ust - bas + 1)
        tutar = gun * ucret * konteyner_sayisi
        sonuc.append((int(gun), tutar))
        ara += tutar
    return sonuc, ara


def _para_yaz(x):
    return f"{x:.2f}"


def ozet_cumle(g, s):
    tip = " ".join(g.konteyner_tipi.split())
    cumle = (f"{g.konteyner_sayisi} adet{' ' + tip if tip else ''} konteyner için tahmini toplam "
             f"{_para_yaz(s.toplam)} {g.para_birimi}. ")
    if g.mod == MOD_AYRI:
        cumle += f"Demuraj süresi {s.demuraj_gun} gün, detention süresi {s.detention_gun} gün"
    else:
        cumle += f"Birleşik süre {s.birlesik_gun} gün"
    if s.ardiye_gun is not None:
        cumle += f", ardiye süresi {s.ardiye_gun} gün"
    cumle += "."
    if g.para_birimi != "TRY" and s.toplam_try is not None:
        cumle += f" Verilen kurla yaklaşık {_para_yaz(s.toplam_try)} TRY."
    cumle += " Bu bir tahmindir, sözleşmenizdeki ve terminal tarifenizdeki koşullar geçerlidir."
    return cumle


def hesapla(g):
    """Girdiden sonucu hesaplar. Girdi hatalıysa GirdiHatasi verir."""
    hata = ilk_hata(g)
    if hata:
        raise GirdiHatasi(hata)
    n = g.konteyner_sayisi
    ilk = g.ilk_gun_sayilir
    demuraj_gun = detention_gun = birlesik_gun = ardiye_gun = None
    satirlar = {t: [None] * SATIR_SAYISI for t in TABLOLAR}
    ara = {t: None for t in TABLOLAR}

    if g.mod == MOD_AYRI:
        demuraj_gun = gun_say(g.tahliye, g.cikis, ilk)
        detention_gun = gun_say(g.cikis, g.bos_iade, ilk)
        satirlar["demuraj"], ara["demuraj"] = tablo_hesapla(g.demuraj, demuraj_gun, n)
        satirlar["detention"], ara["detention"] = tablo_hesapla(g.detention, detention_gun, n)
    else:
        birlesik_gun = gun_say(g.tahliye, g.bos_iade, ilk)
        satirlar["birlesik"], ara["birlesik"] = tablo_hesapla(g.birlesik, birlesik_gun, n)
    if ardiye_kullaniliyor(g):
        ardiye_gun = gun_say(g.tahliye, g.cikis, ilk)
        satirlar["ardiye"], ara["ardiye"] = tablo_hesapla(g.ardiye, ardiye_gun, n)

    toplam = sum(v for v in ara.values() if v is not None)
    if g.para_birimi == "TRY":
        toplam_try = toplam
    elif g.kur_try is not None:
        toplam_try = toplam * g.kur_try
    else:
        toplam_try = None
    s = Sonuc(demuraj_gun, detention_gun, birlesik_gun, ardiye_gun, satirlar, ara, toplam, toplam_try, "")
    s.cumle = ozet_cumle(g, s)
    return s


# ---------------------------------------------------------------- JSON ve komut satırı

def _tarih(x):
    if x in (None, ""):
        return None
    return datetime.strptime(x, "%Y-%m-%d").date()


def girdi_oku(sozluk):
    """JSON sözlüğünü Girdi nesnesine çevirir. Tarifeler [başlangıç, bitiş, ücret] listeleridir."""
    t = sozluk.get("tarifeler", {})
    mod = sozluk.get("mod", MOD_AYRI)
    mod = MOD_METINLERI.get(mod, mod)
    return Girdi(
        konteyner_sayisi=sozluk.get("konteyner_sayisi", 1),
        konteyner_tipi=sozluk.get("konteyner_tipi", ""),
        para_birimi=sozluk.get("para_birimi", "USD"),
        mod=mod,
        ilk_gun_sayilir=sozluk.get("ilk_gun_sayilir", True),
        kur_try=sozluk.get("kur_try"),
        tahliye=_tarih(sozluk.get("tahliye")),
        cikis=_tarih(sozluk.get("cikis")),
        bos_iade=_tarih(sozluk.get("bos_iade")),
        demuraj=_satirlari_duzenle(t.get("demuraj")),
        detention=_satirlari_duzenle(t.get("detention")),
        birlesik=_satirlari_duzenle(t.get("birlesik")),
        ardiye=_satirlari_duzenle(t.get("ardiye")),
    )


def rapor(g, s):
    satirlar = []
    basliklar = {"demuraj": "Demuraj", "detention": "Detention", "birlesik": "Birleşik süre", "ardiye": "Ardiye"}
    satirlar.append("TAHMİNİ HESAP: sözleşmenizdeki ve terminal tarifenizdeki koşullar geçerlidir.")
    satirlar.append("")
    if g.mod == MOD_AYRI:
        satirlar.append(f"Demuraj süresi (tahliye ile çıkış arası): {s.demuraj_gun} gün")
        satirlar.append(f"Detention süresi (çıkış ile boş iade arası): {s.detention_gun} gün")
    else:
        satirlar.append(f"Birleşik süre (tahliye ile boş iade arası): {s.birlesik_gun} gün")
    if s.ardiye_gun is not None:
        satirlar.append(f"Ardiye süresi (tahliye ile çıkış arası): {s.ardiye_gun} gün")
    for tablo in TABLOLAR:
        if s.ara_toplam[tablo] is None:
            continue
        satirlar.append("")
        satirlar.append(f"{basliklar[tablo]}")
        tarife = _satirlari_duzenle(getattr(g, tablo))
        for i, sonuc in enumerate(s.satirlar[tablo]):
            if sonuc is None:
                continue
            bas, bit, ucret = tarife[i]
            aralik = f"{int(bas)}. gün ve sonrası" if bit is None else f"{int(bas)} ile {int(bit)}. gün"
            satirlar.append(f"  {aralik}, günlük {ucret:g} {g.para_birimi}: "
                            f"{sonuc[0]} gün, {_para_yaz(sonuc[1])} {g.para_birimi}")
        satirlar.append(f"  Ara toplam: {_para_yaz(s.ara_toplam[tablo])} {g.para_birimi}")
    satirlar.append("")
    satirlar.append(f"Genel toplam: {_para_yaz(s.toplam)} {g.para_birimi}")
    if s.toplam_try is not None and g.para_birimi != "TRY":
        satirlar.append(f"Genel toplam (TRY): {_para_yaz(s.toplam_try)} TRY")
    satirlar.append("")
    satirlar.append(s.cumle)
    return "\n".join(satirlar)


def ana(argv=None):
    p = argparse.ArgumentParser(
        prog="demuraj.py",
        description="Demuraj, detention ve ardiye maliyetini JSON girdisinden tahmin eder. "
                    "Sonuç bir tahmindir, sözleşmenizdeki tarife geçerlidir.",
    )
    p.add_argument("dosya", help="Girdi JSON dosyası (örnek: ornek.json)")
    args = p.parse_args(argv)
    try:
        with open(args.dosya, encoding="utf-8") as f:
            sozluk = json.load(f)
        g = girdi_oku(sozluk)
        s = hesapla(g)
    except (GirdiHatasi, ValueError, OSError) as hata:
        print(f"Hata: {hata}", file=sys.stderr)
        return 2
    print(rapor(g, s))
    return 0


if __name__ == "__main__":
    sys.exit(ana())
