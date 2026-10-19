"""Bir .xlsx dosyasını LibreOffice ile yeniden hesaplatıp her sayfanın değerlerini döndürür.

Testlerde Excel formüllerinin sonucunu Python'daki referans hesapla karşılaştırmak için kullanılır.
LibreOffice yoksa `LibreOfficeYok` hatası verir, testler bu durumda atlanır.
"""
import csv
import glob
import os
import shutil
import subprocess
import tempfile


class LibreOfficeYok(RuntimeError):
    pass


def soffice_yolu():
    for aday in ("soffice", "/opt/homebrew/bin/soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        yol = shutil.which(aday) or (aday if os.path.exists(aday) else None)
        if yol:
            return yol
    raise LibreOfficeYok("LibreOffice (soffice) bulunamadı")


def hesapla(xlsx_yolu):
    """Sayfa adı -> satır listesi (her satır hücre metinlerinin listesi) döndürür."""
    soffice = soffice_yolu()
    gecici = tempfile.mkdtemp(prefix="lo-hesap-")
    try:
        profil = os.path.join(gecici, "profil")
        cikti = os.path.join(gecici, "cikti")
        os.makedirs(cikti)
        # 44 = virgül, 34 = tırnak, 76 = UTF-8; son -1 tüm sayfaları ayrı CSV olarak yazar.
        filtre = 'csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,false,false,false,-1'
        subprocess.run(
            [soffice, f"-env:UserInstallation=file://{profil}", "--headless", "--norestore",
             "--convert-to", filtre, "--outdir", cikti, os.path.abspath(xlsx_yolu)],
            check=True, capture_output=True, timeout=180,
        )
        taban = os.path.splitext(os.path.basename(xlsx_yolu))[0]
        sonuc = {}
        for dosya in glob.glob(os.path.join(cikti, "*.csv")):
            ad = os.path.splitext(os.path.basename(dosya))[0]
            if ad.startswith(taban + "-"):
                ad = ad[len(taban) + 1:]
            with open(dosya, encoding="utf-8", newline="") as f:
                sonuc[ad] = list(csv.reader(f))
        if not sonuc:
            raise RuntimeError("LibreOffice CSV üretmedi")
        return sonuc
    finally:
        shutil.rmtree(gecici, ignore_errors=True)


def hucre(sayfalar, sayfa, adres):
    """'B7' gibi bir adresin hesaplanmış metnini döndürür, boşsa ''."""
    sutun = 0
    i = 0
    while adres[i].isalpha():
        sutun = sutun * 26 + (ord(adres[i].upper()) - 64)
        i += 1
    satir = int(adres[i:])
    satirlar = sayfalar[sayfa]
    if satir - 1 >= len(satirlar) or sutun - 1 >= len(satirlar[satir - 1]):
        return ""
    return satirlar[satir - 1][sutun - 1]
