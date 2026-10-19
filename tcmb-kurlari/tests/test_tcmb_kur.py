"""tcmb_kur.py testleri. Çalıştırma: python3 -m unittest discover -s tcmb-kurlari/tests -v

Canlı test (TCMB sunucusuna bağlanır) yalnızca LIVE=1 ile çalışır.
"""
import csv
import datetime
import io
import os
import sys
import tempfile
import unittest
import contextlib
from decimal import Decimal

ESKI = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ESKI)
import tcmb_kur  # noqa: E402

FIXTURE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures", "tcmb-16102026.xml")
with open(FIXTURE, "rb") as _f:
    BULTEN_16 = _f.read()

D = datetime.date


class SahteAg:
    """Sadece verilen adresler için bayt döndürür, diğerleri için 404 gibi davranır."""

    def __init__(self, bulutenler):
        self.bulutenler = bulutenler
        self.istekler = []

    def __call__(self, url):
        self.istekler.append(url)
        if url in self.bulutenler:
            return self.bulutenler[url]
        raise tcmb_kur.BultenBulunamadi(url)


def sahte_ag_16():
    return SahteAg({tcmb_kur.bulten_url(D(2026, 10, 16)): BULTEN_16})


class AdresTesti(unittest.TestCase):
    def test_adres(self):
        self.assertEqual(tcmb_kur.bulten_url(D(2026, 10, 16)),
                         "https://www.tcmb.gov.tr/kurlar/202610/16102026.xml")
        self.assertEqual(tcmb_kur.bulten_url(D(2026, 1, 5)),
                         "https://www.tcmb.gov.tr/kurlar/202601/05012026.xml")


class CozmeTesti(unittest.TestCase):
    def setUp(self):
        self.tarih, self.no, self.kurlar = tcmb_kur.bulten_coz(BULTEN_16)

    def test_baslik(self):
        self.assertEqual(self.tarih, D(2026, 10, 16))
        self.assertEqual(self.no, "2026/196")

    def test_doviz_sayisi_ve_usd(self):
        self.assertEqual(len(self.kurlar), 22)
        usd = self.kurlar["USD"]
        self.assertEqual(usd["birim"], 1)
        self.assertEqual(usd["alanlar"]["ForexSelling"], Decimal("49.4463"))
        self.assertEqual(usd["alanlar"]["ForexBuying"], Decimal("49.3573"))
        self.assertEqual(usd["alanlar"]["BanknoteBuying"], Decimal("49.3228"))
        self.assertEqual(usd["alanlar"]["BanknoteSelling"], Decimal("49.5204"))

    def test_bos_alan_none(self):
        xdr = self.kurlar["XDR"]["alanlar"]
        self.assertEqual(xdr["ForexBuying"], Decimal("66.9840"))
        self.assertIsNone(xdr["ForexSelling"])
        self.assertIsNone(xdr["BanknoteSelling"])

    def test_bozuk_xml(self):
        with self.assertRaises(tcmb_kur.TCMBHata):
            tcmb_kur.bulten_coz(b"<html>bozuk")
        with self.assertRaises(tcmb_kur.TCMBHata):
            tcmb_kur.bulten_coz(b"<baska/>")


class KurGetirTesti(unittest.TestCase):
    def test_usd_ve_eur(self):
        ag = sahte_ag_16()
        kur, tarih, no = tcmb_kur.kur_getir(D(2026, 10, 16), "USD", opener=ag)
        self.assertEqual(kur, Decimal("49.4463"))
        self.assertEqual(tarih, D(2026, 10, 16))
        self.assertEqual(no, "2026/196")
        kur, _, _ = tcmb_kur.kur_getir(D(2026, 10, 16), "eur", opener=ag)
        self.assertEqual(kur, Decimal("55.6180"))

    def test_jpy_yuze_bolunur(self):
        kur, _, _ = tcmb_kur.kur_getir(D(2026, 10, 16), "JPY", opener=sahte_ag_16())
        self.assertEqual(kur, Decimal("0.313773"))

    def test_takma_adlar(self):
        ag = sahte_ag_16()
        beklenen = {
            "doviz-alis": Decimal("49.3573"),
            "doviz-satis": Decimal("49.4463"),
            "efektif-alis": Decimal("49.3228"),
            "efektif-satis": Decimal("49.5204"),
            "ForexSelling": Decimal("49.4463"),
            "BanknoteBuying": Decimal("49.3228"),
        }
        for tur, deger in beklenen.items():
            with self.subTest(tur=tur):
                kur, _, _ = tcmb_kur.kur_getir(D(2026, 10, 16), "USD", tur=tur, opener=ag)
                self.assertEqual(kur, deger)

    def test_bilinmeyen_tur(self):
        with self.assertRaises(tcmb_kur.TCMBHata) as bag:
            tcmb_kur.kur_getir(D(2026, 10, 16), "USD", tur="alis", opener=sahte_ag_16())
        self.assertIn("Kur türü anlaşılamadı", str(bag.exception))

    def test_bilinmeyen_doviz(self):
        with self.assertRaises(tcmb_kur.TCMBHata) as bag:
            tcmb_kur.kur_getir(D(2026, 10, 16), "ABC", opener=sahte_ag_16())
        self.assertIn("ABC kodu", str(bag.exception))
        self.assertIn("bültende yok", str(bag.exception))

    def test_bos_alan_hatasi(self):
        with self.assertRaises(tcmb_kur.TCMBHata) as bag:
            tcmb_kur.kur_getir(D(2026, 10, 16), "XDR", tur="doviz-satis", opener=sahte_ag_16())
        self.assertIn("boş", str(bag.exception))

    def test_hafta_sonu_geri_gider(self):
        ag = sahte_ag_16()
        # 18 Ekim 2026 Pazar, 17 Ekim Cumartesi: ikisi de 404, 16 Ekim bülteni kullanılır.
        kur, tarih, no = tcmb_kur.kur_getir(D(2026, 10, 18), "USD", opener=ag)
        self.assertEqual(kur, Decimal("49.4463"))
        self.assertEqual(tarih, D(2026, 10, 16))
        self.assertEqual(no, "2026/196")
        self.assertEqual([u[-12:] for u in ag.istekler], ["18102026.xml", "17102026.xml", "16102026.xml"])

    def test_geri_gun_bitince_hata(self):
        ag = SahteAg({})
        with self.assertRaises(tcmb_kur.TCMBHata) as bag:
            tcmb_kur.kur_getir(D(2026, 10, 18), "USD", geri_gun=3, opener=ag)
        self.assertIn("bülteni bulunamadı", str(bag.exception))
        self.assertEqual(len(ag.istekler), 4)  # istenen gün ve 3 gün geriye

    def test_onbellek_ayni_adresi_bir_kez_ister(self):
        ag = sahte_ag_16()
        onbellek = {}
        for doviz in ("USD", "EUR", "GBP"):
            tcmb_kur.kur_getir(D(2026, 10, 16), doviz, opener=ag, onbellek=onbellek)
        self.assertEqual(len(ag.istekler), 1)


class CsvTesti(unittest.TestCase):
    def test_aralik_csv(self):
        ag = sahte_ag_16()
        with tempfile.TemporaryDirectory() as klasor:
            yol = os.path.join(klasor, "kurlar.csv")
            cikti = io.StringIO()
            with contextlib.redirect_stdout(cikti):
                kod = tcmb_kur.main(
                    ["--bas", "2026-10-16", "--bit", "2026-10-18", "USD", "JPY", "--csv", yol], opener=ag)
            self.assertEqual(kod, 0)
            with open(yol, encoding="utf-8", newline="") as f:
                satirlar = list(csv.DictReader(f))
        self.assertEqual(list(satirlar[0].keys()),
                         ["istenen_tarih", "bulten_tarihi", "bulten_no", "doviz", "tur", "kur"])
        self.assertEqual(len(satirlar), 6)  # 3 gün x 2 döviz
        self.assertEqual(
            [(s["istenen_tarih"], s["doviz"]) for s in satirlar],
            [("2026-10-16", "USD"), ("2026-10-16", "JPY"), ("2026-10-17", "USD"),
             ("2026-10-17", "JPY"), ("2026-10-18", "USD"), ("2026-10-18", "JPY")])
        for s in satirlar:
            self.assertEqual(s["bulten_tarihi"], "2026-10-16")
            self.assertEqual(s["bulten_no"], "2026/196")
            self.assertEqual(s["tur"], "ForexSelling")
        self.assertEqual(satirlar[0]["kur"], "49.4463")
        self.assertEqual(satirlar[1]["kur"], "0.313773")
        self.assertEqual(satirlar[5]["kur"], "0.313773")

    def test_tek_gun_tablosu(self):
        cikti = io.StringIO()
        with contextlib.redirect_stdout(cikti):
            kod = tcmb_kur.main(["2026-10-16", "USD", "EUR", "--tur", "doviz-satis"], opener=sahte_ag_16())
        self.assertEqual(kod, 0)
        metin = cikti.getvalue()
        self.assertIn("49.4463", metin)
        self.assertIn("55.6180", metin)
        self.assertIn("2026/196", metin)

    def test_hata_kodu(self):
        hata = io.StringIO()
        with contextlib.redirect_stderr(hata):
            kod = tcmb_kur.main(["2026-10-16", "ABC"], opener=sahte_ag_16())
        self.assertEqual(kod, 1)
        self.assertIn("Hata:", hata.getvalue())

    def test_tarih_biçimleri(self):
        ag = sahte_ag_16()
        cikti = io.StringIO()
        with contextlib.redirect_stdout(cikti):
            self.assertEqual(tcmb_kur.main(["16.10.2026", "USD"], opener=ag), 0)
        self.assertIn("49.4463", cikti.getvalue())


@unittest.skipUnless(os.environ.get("LIVE") == "1", "Canlı test: LIVE=1 ile çalıştırın")
class CanliTest(unittest.TestCase):
    def test_tcmb_16_ekim_2026(self):
        kur, tarih, no = tcmb_kur.kur_getir(D(2026, 10, 16), "USD", tur="doviz-satis")
        self.assertEqual(tarih, D(2026, 10, 16))
        self.assertEqual(no, "2026/196")
        self.assertEqual(kur, Decimal("49.4463"))
        jpy, _, _ = tcmb_kur.kur_getir(D(2026, 10, 16), "JPY")
        self.assertEqual(jpy, Decimal("0.313773"))
        # Hafta sonu istenince Cuma bülteni gelir.
        _, tarih, _ = tcmb_kur.kur_getir(D(2026, 10, 18), "USD")
        self.assertEqual(tarih, D(2026, 10, 16))


if __name__ == "__main__":
    unittest.main()
