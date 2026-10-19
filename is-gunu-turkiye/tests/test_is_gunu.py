"""is_gunu.py için birim testleri. Beklenen değerler elle takvimden kontrol edilmiştir."""
import contextlib
import csv
import io
import os
import sys
import unittest
from datetime import date

KLASOR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, KLASOR)

import is_gunu as ig  # noqa: E402


class CsvVeriTesti(unittest.TestCase):
    def test_csv_alanlari_gecerli(self):
        with open(ig.VARSAYILAN_CSV, encoding="utf-8", newline="") as f:
            satirlar = list(csv.DictReader(f))
        self.assertGreater(len(satirlar), 40)
        for s in satirlar:
            tarih = date.fromisoformat(s["tarih"])
            self.assertTrue(ig.KAPSAM_BASLANGIC <= tarih <= ig.KAPSAM_BITIS, s)
            self.assertIn(s["sure"], ("tam", "yarim"))
            self.assertTrue(s["kaynak"].startswith("https://vakithesaplama.diyanet.gov.tr/icerik.php?icerik="))

    def test_2027_05_19_tek_kayitta_birlesir(self):
        with open(ig.VARSAYILAN_CSV, encoding="utf-8", newline="") as f:
            ham = list(csv.DictReader(f))
        self.assertEqual(sum(1 for s in ham if s["tarih"] == "2027-05-19"), 2)
        tatiller = ig.tatilleri_yukle()
        kayitlar = [t for t in tatiller if t.tarih == date(2027, 5, 19)]
        self.assertEqual(len(kayitlar), 1)
        self.assertIn("Kurban Bayramı 4. gün", kayitlar[0].ad)
        self.assertIn("Gençlik ve Spor Bayramı", kayitlar[0].ad)
        self.assertEqual(kayitlar[0].sure, "tam")
        self.assertEqual(len(tatiller), len(ham) - 1)

    def test_yarim_gun_listesi(self):
        yarim = sorted(t.tarih for t in ig.tatilleri_yukle() if t.sure == "yarim")
        self.assertIn(date(2026, 3, 19), yarim)
        self.assertIn(date(2026, 10, 28), yarim)
        self.assertEqual(len(yarim), 9)


class IsGunuEkleTesti(unittest.TestCase):
    def test_kurban_bayrami_uzerinden_uc_is_gunu(self):
        # 25.05.2026 Pazartesi. 26 Salı yarım gün sayılır (1), 27-29 tatil, 30-31 hafta sonu,
        # 1 Haziran Pazartesi (2), 2 Haziran Salı (3).
        self.assertEqual(ig.is_gunu_ekle(date(2026, 5, 25), 3), date(2026, 6, 2))

    def test_negatif_sayi_geriye_gider(self):
        self.assertEqual(ig.is_gunu_ekle(date(2026, 6, 2), -3), date(2026, 5, 25))

    def test_sifir_baslangici_aynen_dondurur(self):
        self.assertEqual(ig.is_gunu_ekle(date(2026, 5, 27), 0), date(2026, 5, 27))

    def test_cumartesi_hafta_sonu_ise_bayram_cumartesisi_atlanir(self):
        # 14.05.2027 Cuma. 15 Cumartesi, 16 Pazar, 17-19 tatil. Sonraki iş günü 20 Mayıs Perşembe.
        # 19 Mayıs aynı günde iki tatil olsa da bir gün sayılır.
        self.assertEqual(ig.is_gunu_ekle(date(2027, 5, 14), 1), date(2027, 5, 20))

    def test_cumartesi_calisiliyorsa_yarim_gun_cumartesi_sayilir(self):
        # 15.05.2027 Cumartesi Kurban Bayramı arifesi, yarım gün olduğu için çalışılır.
        self.assertEqual(ig.is_gunu_ekle(date(2027, 5, 14), 1, cumartesi_tatil=False), date(2027, 5, 15))
        self.assertEqual(ig.is_gunu_ekle(date(2027, 5, 14), 2, cumartesi_tatil=False), date(2027, 5, 20))

    def test_cumartesi_calisilsa_da_bayram_cumartesisi_atlanir(self):
        # 24.02.2028 Perşembe. 25 Cuma yarım gün (çalışılır). 26 Cumartesi, 27 Pazar ve 28 Pazartesi
        # Ramazan Bayramı olduğu için sonraki iş günü 29 Şubat Salı (2028 artık yıldır).
        self.assertEqual(ig.is_gunu_ekle(date(2028, 2, 24), 1, cumartesi_tatil=False), date(2028, 2, 25))
        self.assertEqual(ig.is_gunu_ekle(date(2028, 2, 24), 2, cumartesi_tatil=False), date(2028, 2, 29))
        self.assertEqual(ig.is_gunu_ekle(date(2028, 2, 24), 2), date(2028, 2, 29))

    def test_yarim_gun_is_gunu_sayilir(self):
        self.assertTrue(ig.is_gunu_mu(date(2026, 10, 28)))
        self.assertFalse(ig.is_gunu_mu(date(2026, 10, 29)))
        # 26.10.2026 Pazartesi + 4: 27, 28 (yarım), 30, 02.11
        self.assertEqual(ig.is_gunu_ekle(date(2026, 10, 26), 4), date(2026, 11, 2))

    def test_kapsamin_son_gunu_icinde_kalan_sonuc(self):
        # 28.12.2028 Perşembe + 1 iş günü = 29 Aralık Cuma, kapsamın içinde.
        self.assertEqual(ig.is_gunu_ekle(date(2028, 12, 28), 1), date(2028, 12, 29))

    def test_kapsam_disi_baslangic(self):
        with self.assertRaises(ig.KapsamDisiHata) as hata:
            ig.is_gunu_ekle(date(2025, 12, 31), 1)
        self.assertIn("kapsamı dışında", str(hata.exception))
        with self.assertRaises(ig.KapsamDisiHata):
            ig.is_gunu_ekle(date(2029, 1, 2), -1)

    def test_kapsam_disi_sonuc(self):
        with self.assertRaises(ig.KapsamDisiHata):
            ig.is_gunu_ekle(date(2028, 12, 29), 5)
        with self.assertRaises(ig.KapsamDisiHata):
            ig.is_gunu_ekle(date(2026, 1, 5), -3)
        # 02.01.2026 Cuma. Geriye 1 iş günü: 1 Ocak tatil, 31 Aralık 2025 kapsam dışı.
        with self.assertRaises(ig.KapsamDisiHata):
            ig.is_gunu_ekle(date(2026, 1, 2), -1)

    def test_tam_sayi_olmayan_sayi(self):
        with self.assertRaises(ValueError):
            ig.is_gunu_ekle(date(2026, 5, 25), 2.5)


class IsGunuSayTesti(unittest.TestCase):
    def test_ramazan_bayrami_2026(self):
        # 16-27 Mart 2026: 10 hafta içi günden 20 Mart Cuma düşer. 19 Mart yarım gün sayılır.
        self.assertEqual(ig.is_gunu_say(date(2026, 3, 16), date(2026, 3, 27)), 9)

    def test_cumartesi_calisilir_modu(self):
        # 14-28 Mart 2026. Hafta içi: 16-19 (4) ve 23-27 (5). Çalışılan Cumartesiler: 14 ve 28.
        # 21 Mart Cumartesi bayram olduğu için çalışılmaz.
        self.assertEqual(ig.is_gunu_say(date(2026, 3, 14), date(2026, 3, 28)), 9)
        self.assertEqual(ig.is_gunu_say(date(2026, 3, 14), date(2026, 3, 28), cumartesi_tatil=False), 11)

    def test_mayis_2026(self):
        # 21 hafta içi günden 1, 19, 27, 28, 29 Mayıs düşer.
        self.assertEqual(ig.is_gunu_say(date(2026, 5, 1), date(2026, 5, 31)), 16)

    def test_tek_gun(self):
        self.assertEqual(ig.is_gunu_say(date(2026, 5, 25), date(2026, 5, 25)), 1)
        self.assertEqual(ig.is_gunu_say(date(2026, 5, 27), date(2026, 5, 27)), 0)

    def test_ayni_gune_iki_tatil_bir_kez_sayilir(self):
        # 17-19 Mayıs 2027 üç gün de tatil. 19 Mayıs iki kayıttan oluşsa da düşen gün üç.
        o = ig.ozet(date(2027, 5, 17), date(2027, 5, 19))
        self.assertEqual(o["is_gunu"], 0)
        self.assertEqual(o["tam_gun_tatil"], 3)
        self.assertEqual(o["tatil_nedeniyle_dusen_is_gunu"], 3)

    def test_ozet_mayis_2026(self):
        o = ig.ozet(date(2026, 5, 1), date(2026, 5, 31))
        self.assertEqual(o, {"is_gunu": 16, "takvim_gunu": 31, "tam_gun_tatil": 6,
                             "yarim_gun_tatil": 1, "tatil_nedeniyle_dusen_is_gunu": 5})

    def test_ozet_cumartesi_calisilir_bayram_cumartesisi(self):
        # 14-21 Mayıs 2027, Cumartesi çalışılır. Tatil nedeniyle düşenler: 17, 18, 19. 16 Pazar zaten tatil.
        o = ig.ozet(date(2027, 5, 14), date(2027, 5, 21), cumartesi_tatil=False)
        self.assertEqual(o["is_gunu"], 4)  # 14, 15 (yarım gün), 20, 21
        self.assertEqual(o["tatil_nedeniyle_dusen_is_gunu"], 3)
        self.assertEqual(o["yarim_gun_tatil"], 1)

    def test_takvim_gunu(self):
        self.assertEqual(ig.takvim_gunu(date(2026, 5, 1), date(2026, 5, 31)), 31)
        self.assertEqual(ig.takvim_gunu(date(2026, 5, 1), date(2026, 5, 1)), 1)

    def test_kapsam_disi(self):
        with self.assertRaises(ig.KapsamDisiHata):
            ig.is_gunu_say(date(2025, 12, 30), date(2026, 1, 5))
        with self.assertRaises(ig.KapsamDisiHata):
            ig.is_gunu_say(date(2028, 12, 30), date(2029, 1, 3))

    def test_ters_sira(self):
        with self.assertRaises(ig.TarihSirasiHatasi):
            ig.is_gunu_say(date(2026, 5, 31), date(2026, 5, 1))


class TatilListesiTesti(unittest.TestCase):
    def test_aradaki_tatiller_yarim_gunleri_de_icerir(self):
        liste = ig.aradaki_tatiller(date(2026, 10, 26), date(2026, 10, 30))
        self.assertEqual([(t.tarih, t.sure) for t in liste],
                         [(date(2026, 10, 28), "yarim"), (date(2026, 10, 29), "tam")])

    def test_aradaki_tatiller_hafta_sonuna_dusenleri_de_listeler(self):
        liste = ig.aradaki_tatiller(date(2026, 3, 19), date(2026, 3, 22))
        self.assertEqual(len(liste), 4)

    def test_aradaki_tatiller_kapsam_disi(self):
        with self.assertRaises(ig.KapsamDisiHata):
            ig.aradaki_tatiller(date(2025, 12, 1), date(2026, 1, 1))

    def test_gun_adlari_ve_tarih_okuma(self):
        self.assertEqual(ig.gun_adi(date(2026, 6, 2)), "Salı")
        self.assertEqual(ig.gun_adi(date(2026, 5, 30)), "Cumartesi")
        self.assertEqual(ig.tarih_oku("2026-05-25"), date(2026, 5, 25))
        self.assertEqual(ig.tarih_oku("25.05.2026"), date(2026, 5, 25))
        self.assertEqual(ig.tarih_yaz(date(2026, 5, 5)), "05.05.2026")
        with self.assertRaises(ValueError):
            ig.tarih_oku("25/05/2026")


class KomutSatiriTesti(unittest.TestCase):
    def calistir(self, argumanlar):
        cikti, hata = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(cikti), contextlib.redirect_stderr(hata):
            kod = ig.ana(argumanlar)
        return kod, cikti.getvalue(), hata.getvalue()

    def test_ekle(self):
        kod, cikti, _ = self.calistir(["ekle", "2026-05-25", "3"])
        self.assertEqual(kod, 0)
        self.assertIn("Sonuç: 02.06.2026 Salı", cikti)
        self.assertIn("Kurban Bayramı arifesi", cikti)

    def test_say_cumartesi_secenegi(self):
        kod, cikti, _ = self.calistir(["say", "2026-03-14", "2026-03-28", "--cumartesi-calisilir"])
        self.assertEqual(kod, 0)
        self.assertIn("İş günü sayısı (iki uç dahil): 11", cikti)

    def test_kapsam_disi_hata_kodu(self):
        kod, cikti, hata = self.calistir(["ekle", "2029-01-02", "1"])
        self.assertEqual(kod, 2)
        self.assertEqual(cikti, "")
        self.assertIn("kapsamı dışında", hata)

    def test_gecersiz_tarih(self):
        kod, _, hata = self.calistir(["say", "dun", "2026-05-01"])
        self.assertEqual(kod, 2)
        self.assertIn("Tarih anlaşılamadı", hata)


if __name__ == "__main__":
    unittest.main()
