"""demuraj.py için birim testleri. Beklenen değerler elle hesaplanmıştır."""
import contextlib
import io
import json
import os
import sys
import tempfile
import unittest
from dataclasses import replace
from datetime import date, timedelta

KLASOR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, KLASOR)

import demuraj as d  # noqa: E402

with open(os.path.join(KLASOR, "ornek.json"), encoding="utf-8") as f:
    ORNEK_SOZLUK = json.load(f)
ORNEK = d.girdi_oku(ORNEK_SOZLUK)

DEMURAJ = [(1, 7, 0), (8, 14, 40), (15, None, 80)]
BIRLESIK = [(1, 10, 0), (11, 20, 45), (21, None, 90)]
TAHLIYE = date(2026, 9, 1)


def temel(**degisiklik):
    """Tek konteyner, yalnız demuraj tablosu dolu, ilk gün sayılmaz, USD."""
    g = d.Girdi(konteyner_sayisi=1, para_birimi="USD", mod=d.MOD_AYRI, ilk_gun_sayilir=False,
                tahliye=TAHLIYE, cikis=TAHLIYE + timedelta(days=10), bos_iade=TAHLIYE + timedelta(days=10),
                demuraj=DEMURAJ)
    return replace(g, **degisiklik)


def gunle(n, **degisiklik):
    """Tahliyeden n gün sonra çıkış ve aynı gün boş iade."""
    return temel(cikis=TAHLIYE + timedelta(days=n), bos_iade=TAHLIYE + timedelta(days=n), **degisiklik)


class OrnekTesti(unittest.TestCase):
    def test_ornek_json_sonucu(self):
        s = d.hesapla(ORNEK)
        # Demuraj 12 gün: 8-12 arası 5 gün x 40 x 2 konteyner = 400
        # Detention 9 gün: 6-9 arası 4 gün x 35 x 2 = 280
        # Ardiye 12 gün: 4-10 arası 7 gün x 12 x 2 = 168, 11-12 arası 2 gün x 25 x 2 = 100
        self.assertEqual((s.demuraj_gun, s.detention_gun, s.birlesik_gun, s.ardiye_gun), (12, 9, None, 12))
        self.assertEqual(s.ara_toplam, {"demuraj": 400, "detention": 280, "birlesik": None, "ardiye": 268})
        self.assertEqual(s.toplam, 948)
        self.assertIsNone(s.toplam_try)
        self.assertEqual(
            s.cumle,
            "2 adet 40' HC konteyner için tahmini toplam 948.00 USD. Demuraj süresi 12 gün, "
            "detention süresi 9 gün, ardiye süresi 12 gün. Bu bir tahmindir, sözleşmenizdeki ve terminal "
            "tarifenizdeki koşullar geçerlidir.")

    def test_excel_modu_metni_json_icinde_kabul_edilir(self):
        sozluk = dict(ORNEK_SOZLUK, mod="Birleşik (tek serbest süre)")
        self.assertEqual(d.girdi_oku(sozluk).mod, d.MOD_BIRLESIK)


class SureVeKademeTesti(unittest.TestCase):
    def test_serbest_sure_icinde_iade_sifir(self):
        for n in (1, 5, 7):
            s = d.hesapla(gunle(n))
            self.assertEqual(s.demuraj_gun, n)
            self.assertEqual(s.toplam, 0, n)

    def test_kademe_sinirlari(self):
        # İlk gün sayılmadığında gün sayısı tarih farkına eşittir.
        # 8. gün 40, 14. gün 7 x 40 = 280, 15. gün 280 + 80 = 360, 30. gün 280 + 16 x 80 = 1560.
        beklenen = {7: 0, 8: 40, 14: 280, 15: 360, 30: 1560}
        for n, tutar in beklenen.items():
            with self.subTest(gun=n):
                self.assertEqual(d.hesapla(gunle(n)).toplam, tutar)

    def test_son_kademe_acik_uclu(self):
        s = d.hesapla(gunle(60))
        satir = s.satirlar["demuraj"]
        self.assertEqual(satir[2], (46, 46 * 80))  # 15. ile 60. gün arası 46 gün

    def test_ilk_gun_evet_ve_hayir(self):
        # 1 Eylül ile 12 Eylül arası: Evet 12 gün, Hayır 11 gün.
        evet = d.hesapla(temel(cikis=date(2026, 9, 12), bos_iade=date(2026, 9, 12), ilk_gun_sayilir=True))
        hayir = d.hesapla(temel(cikis=date(2026, 9, 12), bos_iade=date(2026, 9, 12), ilk_gun_sayilir=False))
        self.assertEqual((evet.demuraj_gun, evet.toplam), (12, 5 * 40))
        self.assertEqual((hayir.demuraj_gun, hayir.toplam), (11, 4 * 40))

    def test_ayni_gun(self):
        evet = d.hesapla(gunle(0, ilk_gun_sayilir=True))
        hayir = d.hesapla(gunle(0, ilk_gun_sayilir=False))
        self.assertEqual(evet.demuraj_gun, 1)
        self.assertEqual(hayir.demuraj_gun, 0)
        self.assertEqual(hayir.toplam, 0)

    def test_bosluklu_satirlar_arada_kalan_gunu_ucretlendirmez(self):
        g = gunle(12, demuraj=[(1, 7, 0), (10, 14, 40)])
        # 8. ve 9. gün hiçbir satıra girmez. 10-12 arası 3 gün x 40.
        self.assertEqual(d.hesapla(g).toplam, 120)

    def test_bos_satirlar_atlanir(self):
        g = gunle(10, demuraj=[(None, None, None), (8, 14, 40), (None, None, None)])
        s = d.hesapla(g)
        self.assertEqual(s.satirlar["demuraj"][0], None)
        self.assertEqual(s.satirlar["demuraj"][1], (3, 120))
        self.assertEqual(s.toplam, 120)


class ModTesti(unittest.TestCase):
    def test_birlesik_mod(self):
        # 1 Eylül ile 20 Eylül arası, ilk gün sayılmaz: 19 gün. 11-19 arası 9 gün x 45 = 405.
        g = temel(mod=d.MOD_BIRLESIK, birlesik=BIRLESIK, cikis=None, bos_iade=date(2026, 9, 20))
        s = d.hesapla(g)
        self.assertEqual(s.birlesik_gun, 19)
        self.assertIsNone(s.demuraj_gun)
        self.assertIsNone(s.detention_gun)
        self.assertEqual(s.toplam, 405)
        # İlk gün sayılırsa 20 gün: 11-20 arası 10 gün x 45 = 450.
        self.assertEqual(d.hesapla(replace(g, ilk_gun_sayilir=True)).toplam, 450)

    def test_birlesik_modda_ayri_tablolar_yok_sayilir(self):
        g = temel(mod=d.MOD_BIRLESIK, birlesik=BIRLESIK, cikis=None, bos_iade=date(2026, 9, 20),
                  demuraj=DEMURAJ, detention=DEMURAJ)
        s = d.hesapla(g)
        self.assertEqual(s.satirlar["demuraj"], [None] * 5)
        self.assertEqual(s.ara_toplam["demuraj"], None)

    def test_ayri_modda_birlesik_tablo_yok_sayilir(self):
        s = d.hesapla(replace(ORNEK, birlesik=BIRLESIK))
        self.assertEqual(s.toplam, 948)
        self.assertIsNone(s.birlesik_gun)

    def test_ardiye_bos_ise_sifir(self):
        s = d.hesapla(replace(ORNEK, ardiye=[]))
        self.assertIsNone(s.ardiye_gun)
        self.assertIsNone(s.ara_toplam["ardiye"])
        self.assertEqual(s.toplam, 680)
        self.assertNotIn("ardiye süresi", s.cumle)

    def test_ardiye_birlesik_modda_cikis_tarihiyle_hesaplanir(self):
        g = replace(ORNEK, mod=d.MOD_BIRLESIK, birlesik=BIRLESIK)
        s = d.hesapla(g)
        # Birleşik süre 1-20 Eylül, ilk gün sayılır: 20 gün, 11-20 arası 10 x 45 x 2 = 900. Ardiye 268.
        self.assertEqual((s.birlesik_gun, s.ardiye_gun), (20, 12))
        self.assertEqual(s.toplam, 900 + 268)

    def test_konteyner_sayisi_carpar(self):
        bir = d.hesapla(replace(ORNEK, konteyner_sayisi=1)).toplam
        uc = d.hesapla(replace(ORNEK, konteyner_sayisi=3)).toplam
        self.assertEqual(bir, 474)
        self.assertEqual(uc, 3 * 474)


class KurTesti(unittest.TestCase):
    def test_try_karsiligi(self):
        s = d.hesapla(replace(ORNEK, kur_try=40.0))
        self.assertAlmostEqual(s.toplam_try, 948 * 40)
        self.assertIn("Verilen kurla yaklaşık 37920.00 TRY.", s.cumle)

    def test_kur_girilmezse_try_toplam_yok(self):
        self.assertIsNone(d.hesapla(ORNEK).toplam_try)

    def test_para_birimi_try_ise_kur_yok_sayilir(self):
        s = d.hesapla(replace(ORNEK, para_birimi="TRY", kur_try=99.0))
        self.assertEqual(s.toplam_try, s.toplam)
        self.assertNotIn("Verilen kurla", s.cumle)

    def test_avro(self):
        s = d.hesapla(replace(ORNEK, para_birimi="EUR", kur_try=38.5))
        self.assertAlmostEqual(s.toplam_try, 948 * 38.5)
        self.assertIn("948.00 EUR", s.cumle)


class HataTesti(unittest.TestCase):
    def hata(self, g):
        with self.assertRaises(d.GirdiHatasi) as e:
            d.hesapla(g)
        return str(e.exception)

    def test_tarih_sirasi(self):
        self.assertEqual(self.hata(replace(ORNEK, cikis=date(2026, 8, 31))), d.M_SIRA_CIKIS)
        self.assertEqual(self.hata(replace(ORNEK, bos_iade=date(2026, 9, 10))), d.M_SIRA_IADE_CIKIS)
        self.assertEqual(self.hata(replace(ORNEK, cikis=date(2026, 9, 12), bos_iade=date(2026, 8, 30))),
                         d.M_SIRA_IADE_TAHLIYE)

    def test_eksik_tarih(self):
        self.assertEqual(self.hata(replace(ORNEK, tahliye=None)), d.M_TAHLIYE)
        self.assertEqual(self.hata(replace(ORNEK, cikis=None)), d.M_CIKIS)
        self.assertEqual(self.hata(replace(ORNEK, bos_iade=None)), d.M_IADE)

    def test_birlesik_modda_cikis_yalniz_ardiye_icin_gerekir(self):
        g = replace(ORNEK, mod=d.MOD_BIRLESIK, birlesik=BIRLESIK, cikis=None)
        self.assertEqual(self.hata(g), d.M_ARDIYE_CIKIS)
        self.assertEqual(d.hesapla(replace(g, ardiye=[])).birlesik_gun, 20)

    def test_gecersiz_ayarlar(self):
        self.assertEqual(self.hata(replace(ORNEK, konteyner_sayisi=0)), d.M_SAYI)
        self.assertEqual(self.hata(replace(ORNEK, konteyner_sayisi=1.5)), d.M_SAYI)
        self.assertEqual(self.hata(replace(ORNEK, para_birimi="GBP")), d.M_PARA)
        self.assertEqual(self.hata(replace(ORNEK, mod="karisik")), d.M_MOD)
        self.assertEqual(self.hata(replace(ORNEK, kur_try=-1)), d.M_KUR)
        self.assertEqual(self.hata(replace(ORNEK, kur_try=0)), d.M_KUR)

    def test_kur_hatasi_try_para_biriminde_yok_sayilir(self):
        self.assertEqual(d.hesapla(replace(ORNEK, para_birimi="TRY", kur_try=-1)).toplam, 948)

    def test_tarife_satiri_hatalari(self):
        kotu = {
            "bitis_baslangictan_kucuk": [(8, 5, 40)],
            "ucret_eksi": [(1, 7, -1)],
            "ucret_yok": [(1, 7, None)],
            "baslangic_sifir": [(0, 7, 10)],
            "baslangic_ondalik": [(1.5, 7, 10)],
            "bitis_ondalik": [(1, 7.5, 10)],
        }
        for ad, satirlar in kotu.items():
            with self.subTest(ad):
                self.assertEqual(self.hata(replace(ORNEK, demuraj=satirlar)), d.M_TARIFE)

    def test_alti_satir_sessizce_kesilmez(self):
        satirlar = [(i, i, 1) for i in range(1, 7)]
        self.assertEqual(self.hata(replace(ORNEK, demuraj=satirlar)), d.M_COK_SATIR)

    def test_kullanilmayan_tablodaki_hata_yok_sayilir(self):
        g = replace(ORNEK, birlesik=[(8, 5, 40)])
        self.assertEqual(d.hesapla(g).toplam, 948)

    def test_satir_notlari(self):
        self.assertEqual(d.satir_notu((None, None, None)), "")
        self.assertEqual(d.satir_notu((1, None, 10)), "")
        self.assertEqual(d.satir_notu((8, 5, 10)), d.R_BITIS_KUCUK)
        self.assertEqual(d.satir_notu((1, 7, None)), d.R_UCRET_YOK)
        self.assertEqual(d.satir_notu((1, 7, -3)), d.R_UCRET_EKSI)
        self.assertEqual(d.satir_notu((0, 7, 3)), d.R_BASLANGIC)


class KomutSatiriTesti(unittest.TestCase):
    def calistir(self, argumanlar):
        cikti, hata = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(cikti), contextlib.redirect_stderr(hata):
            kod = d.ana(argumanlar)
        return kod, cikti.getvalue(), hata.getvalue()

    def test_ornek_dosyasi(self):
        kod, cikti, _ = self.calistir([os.path.join(KLASOR, "ornek.json")])
        self.assertEqual(kod, 0)
        self.assertIn("Genel toplam: 948.00 USD", cikti)
        self.assertIn("TAHMİNİ HESAP", cikti)
        self.assertIn("15. gün ve sonrası", cikti)

    def test_eksik_tarih_hata_kodu_verir(self):
        sozluk = dict(ORNEK_SOZLUK, tahliye=None)
        with tempfile.TemporaryDirectory() as klasor:
            yol = os.path.join(klasor, "girdi.json")
            with open(yol, "w", encoding="utf-8") as f:
                json.dump(sozluk, f)
            kod, cikti, hata = self.calistir([yol])
        self.assertEqual(kod, 2)
        self.assertEqual(cikti, "")
        self.assertIn(d.M_TAHLIYE, hata)

    def test_olmayan_dosya(self):
        kod, _, hata = self.calistir(["yok.json"])
        self.assertEqual(kod, 2)
        self.assertIn("Hata", hata)


if __name__ == "__main__":
    unittest.main()
