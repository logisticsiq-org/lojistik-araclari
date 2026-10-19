"""kontrol_hanesi.py birim testleri.

Çalıştırma: python3 -m unittest discover -s konteyner-awb-kontrol/tests -v
"""
import io
import os
import sys
import tempfile
import unittest
import contextlib

ARAC_KLASORU = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ARAC_KLASORU)
import kontrol_hanesi as kh  # noqa: E402

GECERLI_KONTEYNER = "CSQU3054383"
GECERLI_AWB = "02012345675"


def referans_konteyner_hanesi(no10):
    """Bağımsız ikinci hesap: harf değerlerini elle yazılmış tablodan alır."""
    tablo = dict(zip("ABCDEFGHIJKLMNOPQRSTUVWXYZ",
                     [10, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 23, 24, 25, 26, 27, 28, 29,
                      30, 31, 32, 34, 35, 36, 37, 38]))
    toplam = sum((tablo[k] if k.isalpha() else int(k)) * 2 ** i for i, k in enumerate(no10))
    return toplam % 11 % 10


class HarfTablosuTesti(unittest.TestCase):
    def test_harf_degerleri_ozellige_uygun(self):
        beklenen = {"A": 10, "B": 12, "C": 13, "D": 14, "E": 15, "F": 16, "G": 17, "H": 18, "I": 19,
                    "J": 20, "K": 21, "L": 23, "M": 24, "N": 25, "O": 26, "P": 27, "Q": 28, "R": 29,
                    "S": 30, "T": 31, "U": 32, "V": 34, "W": 35, "X": 36, "Y": 37, "Z": 38}
        self.assertEqual(kh.HARF_DEGERLERI, beklenen)
        self.assertFalse(any(v % 11 == 0 for v in kh.HARF_DEGERLERI.values()))


class KonteynerTesti(unittest.TestCase):
    def test_bilinen_gecerli_ornek(self):
        self.assertEqual(kh.konteyner_kontrol(GECERLI_KONTEYNER), (kh.GECERLI, 3))
        self.assertEqual(kh.konteyner_kontrol_hanesi("CSQU305438"), 3)

    def test_yanlis_hane(self):
        self.assertEqual(kh.konteyner_kontrol("CSQU3054384"), (kh.HANE_HATALI, 3))
        self.assertEqual(kh.konteyner_kontrol("CSQU3054380"), (kh.HANE_HATALI, 3))

    def test_tek_rakam_degisikligi_yakalanir(self):
        """Altı seri rakamının ve kontrol hanesinin her tek rakam değişikliği yakalanır."""
        sayac = 0
        for konum in range(4, 11):
            for yeni in "0123456789":
                if yeni == GECERLI_KONTEYNER[konum]:
                    continue
                numara = GECERLI_KONTEYNER[:konum] + yeni + GECERLI_KONTEYNER[konum + 1:]
                with self.subTest(numara=numara):
                    self.assertEqual(kh.konteyner_kontrol(numara)[0], kh.HANE_HATALI)
                sayac += 1
        self.assertEqual(sayac, 63)

    def test_komsu_rakam_yer_degistirme_cogunlukla_yakalanir(self):
        yakalanmayan = 0
        toplam = 0
        for konum in range(4, 10):
            a, b = GECERLI_KONTEYNER[konum], GECERLI_KONTEYNER[konum + 1]
            if a == b:
                continue
            numara = (GECERLI_KONTEYNER[:konum] + b + a + GECERLI_KONTEYNER[konum + 2:])
            toplam += 1
            if kh.konteyner_kontrol(numara)[0] == kh.GECERLI:
                yakalanmayan += 1
        self.assertGreater(toplam, 0)
        self.assertLess(yakalanmayan, toplam)

    def test_normallestirme(self):
        for girdi in ("csqu3054383", "CSQU 305438 3", "csqu-305438-3", "CSQU.305438.3", "CSQU/3054383",
                      "  csqu 3054383 "):
            with self.subTest(girdi=girdi):
                self.assertEqual(kh.konteyner_kontrol(girdi), (kh.GECERLI, 3))
        self.assertEqual(kh.normalle("csqu 305-438.3/"), "CSQU3054383")
        self.assertEqual(kh.normalle(None), "")

    def test_kategori_harfi(self):
        for kategori in "UJZ":
            numara10 = "CSQ" + kategori + "305438"
            numara = numara10 + str(kh.konteyner_kontrol_hanesi(numara10))
            with self.subTest(kategori=kategori):
                self.assertEqual(kh.konteyner_kontrol(numara)[0], kh.GECERLI)
        for kategori in "ABCDEFGHIKLMNOPQRSTVWXY":
            with self.subTest(kategori=kategori):
                self.assertEqual(kh.konteyner_kontrol("CSQ" + kategori + "3054383"), (kh.BICIM_HATALI, None))

    def test_uzunluk_ve_icerik(self):
        for girdi in ("", " ", "CSQU305438", "CSQU30543833", "CSQU305438", "CSQ3054383", "CS3U3054383",
                      "CSQU30543A3", "1234567890A", "0201234567", "CSQUü054383"):
            with self.subTest(girdi=girdi):
                self.assertEqual(kh.konteyner_kontrol(girdi), (kh.BICIM_HATALI, None))

    def test_noktasiz_i_gecersiz(self):
        self.assertEqual(kh.konteyner_kontrol("CSQUı054383")[0], kh.BICIM_HATALI)

    def test_on_artik_10_olur_0(self):
        """Mod 11 sonucu 10 çıkarsa kontrol hanesi 0 olur."""
        bulunan = None
        for sayi in range(1000000):
            no10 = "MSKU%06d" % sayi
            toplam = sum((kh.HARF_DEGERLERI[k] if k.isalpha() else int(k)) * 2 ** i for i, k in enumerate(no10))
            if toplam % 11 == 10:
                bulunan = no10
                break
        self.assertIsNotNone(bulunan)
        self.assertEqual(kh.konteyner_kontrol_hanesi(bulunan), 0)
        self.assertEqual(kh.konteyner_kontrol(bulunan + "0"), (kh.GECERLI, 0))

    def test_ikinci_hesapla_ayni(self):
        for sahip in ("MSKU", "OOLU", "TRIU", "ZIMU", "HLXJ", "ABCZ"):
            for sayi in range(0, 999999, 12347):
                no10 = "%s%06d" % (sahip, sayi)
                with self.subTest(no10=no10):
                    self.assertEqual(kh.konteyner_kontrol_hanesi(no10), referans_konteyner_hanesi(no10))

    def test_hane_hesabi_gecersiz_girdi(self):
        with self.assertRaises(ValueError):
            kh.konteyner_kontrol_hanesi("CSQU30543")
        with self.assertRaises(ValueError):
            kh.konteyner_kontrol_hanesi("CSQX305438")


class AwbTesti(unittest.TestCase):
    def test_bilinen_gecerli_ornek(self):
        self.assertEqual(kh.awb_kontrol("020-12345675"), (kh.GECERLI, 5))
        self.assertEqual(kh.awb_kontrol(GECERLI_AWB), (kh.GECERLI, 5))

    def test_yanlis_hane(self):
        self.assertEqual(kh.awb_kontrol("020-12345676"), (kh.HANE_HATALI, 5))

    def test_kontrol_hanesi_degisikligi_hep_yakalanir(self):
        for yeni in "0123456789":
            if yeni == "5":
                continue
            with self.subTest(yeni=yeni):
                self.assertEqual(kh.awb_kontrol("020-1234567" + yeni)[0], kh.HANE_HATALI)

    def test_seri_rakam_degisikligi_yedi_fark_disinda_yakalanir(self):
        """İlk 7 seri rakamında yapılan her değişiklik yakalanır, farkı tam 7 olanlar hariç.

        Mod 7 kontrolü bu farkı göremez. Bu, yöntemin bilinen bir sınırıdır.
        """
        yakalanmayan = []
        for konum in range(3, 10):
            for yeni in "0123456789":
                eski = GECERLI_AWB[konum]
                if yeni == eski:
                    continue
                numara = GECERLI_AWB[:konum] + yeni + GECERLI_AWB[konum + 1:]
                durum = kh.awb_kontrol(numara)[0]
                if abs(int(yeni) - int(eski)) == 7:
                    self.assertEqual(durum, kh.GECERLI, numara)
                    yakalanmayan.append(numara)
                else:
                    self.assertEqual(durum, kh.HANE_HATALI, numara)
        self.assertGreater(len(yakalanmayan), 0)

    def test_onek_kontrol_hanesine_dahil_degil(self):
        for onek in ("000", "176", "999"):
            self.assertEqual(kh.awb_kontrol(onek + GECERLI_AWB[3:])[0], kh.GECERLI)

    def test_normallestirme(self):
        for girdi in ("020-12345675", "020 12345675", "020.12345675", "020/12345675", " 020 - 1234 5675 "):
            with self.subTest(girdi=girdi):
                self.assertEqual(kh.awb_kontrol(girdi), (kh.GECERLI, 5))

    def test_sifir_kalani(self):
        self.assertEqual(kh.awb_kontrol("000-00000000"), (kh.GECERLI, 0))
        self.assertEqual(kh.awb_kontrol("000-00000007"), (kh.HANE_HATALI, 0))

    def test_uzunluk_ve_icerik(self):
        for girdi in ("", " ", "0201234567", "020123456750", "02A12345675", "CSQU3054383", "020-1234567X"):
            with self.subTest(girdi=girdi):
                self.assertEqual(kh.awb_kontrol(girdi), (kh.BICIM_HATALI, None))


class SonucMetniTesti(unittest.TestCase):
    def test_metinler(self):
        self.assertEqual(kh.sonuc_metni(kh.GECERLI, 3), "Geçerli")
        self.assertEqual(kh.sonuc_metni(kh.BICIM_HATALI, None), "Biçim hatalı")
        self.assertEqual(kh.sonuc_metni(kh.HANE_HATALI, 3), "Kontrol hanesi hatalı (doğrusu: 3)")


class KomutSatiriTesti(unittest.TestCase):
    def calistir(self, argumanlar):
        cikti = io.StringIO()
        with contextlib.redirect_stdout(cikti):
            kod = kh.main(argumanlar)
        return kod, cikti.getvalue()

    def test_gecerli_numaralar_sifir_kodu(self):
        kod, cikti = self.calistir([GECERLI_KONTEYNER, "020-12345675"])
        self.assertEqual(kod, 0)
        self.assertIn("Geçerli", cikti)
        self.assertIn("konteyner", cikti)
        self.assertIn("awb", cikti)

    def test_hatali_numara_bir_kodu(self):
        kod, cikti = self.calistir(["CSQU3054384"])
        self.assertEqual(kod, 1)
        self.assertIn("doğrusu: 3", cikti)

    def test_csv(self):
        with tempfile.TemporaryDirectory() as klasor:
            girdi = os.path.join(klasor, "in.csv")
            cikti = os.path.join(klasor, "out.csv")
            with open(girdi, "w", encoding="utf-8", newline="") as f:
                f.write("referans,numara\nA1,csqu 305438 3\nA2,CSQU3054384\nA3,020-12345675\nA4,\nA5,CSQX3054383\n")
            kod, _ = self.calistir(["--csv", girdi, "--sutun", "2", "--cikti", cikti, "--baslik"])
            self.assertEqual(kod, 0)
            with open(cikti, encoding="utf-8") as f:
                satirlar = [s.split(",") for s in f.read().strip().split("\n")]
        self.assertEqual(satirlar[0], ["referans", "numara", "tur", "normalle_no", "durum", "beklenen_hane"])
        self.assertEqual(satirlar[1][2:], ["konteyner", "CSQU3054383", "gecerli", "3"])
        self.assertEqual(satirlar[2][2:], ["konteyner", "CSQU3054384", "hane_hatali", "3"])
        self.assertEqual(satirlar[3][2:], ["awb", "02012345675", "gecerli", "5"])
        self.assertEqual(satirlar[4][2:], ["", "", "", ""])
        self.assertEqual(satirlar[5][2:], ["konteyner", "CSQX3054383", "bicim_hatali", ""])


if __name__ == "__main__":
    unittest.main()
