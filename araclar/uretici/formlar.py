# -*- coding: utf-8 -*-
"""Yeni / yeniden tasarlanan formları üretir:  python3 formlar.py"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from accessgen import *

HEDEF = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'kaynak', 'formlar')
MUSTERI_SQL = 'SELECT MusteriID, AdSoyad FROM tbl_Musteri ORDER BY AdSoyad;'
PERSONEL_SQL = 'SELECT PersonelID, AdSoyad FROM tbl_Personel ORDER BY AdSoyad;'
SATIR = 420


def satirlar(bas, adet, aralik=SATIR):
    return [bas + i * aralik for i in range(adet)]


# ---------------------------------------------------------------
def mal_kabul():
    f = Form('frm_MalKabul', 'Mal Kabul', 15800, 10250, onload='=MalKabulYukle()', bant='MAL KABUL  (ZEYTİN GİRİŞİ)')
    r = satirlar(1250, 5)
    f.bolum('lblB1', 'PARTİ', 300, 850, 4000)
    f.bolum('lblB2', 'ZEYTİN', 5300, 850, 4000)
    f.bolum('lblB3', 'TARTIM', 10500, 850, 4000)
    # A sütunu
    f.alan('Parti No', 'txtPartiKodu', 300, r[0], 1700, 1200, kalin=True)
    f.etiket('lblSezonE', 'Sezon', 3280, r[0] + 30, 680)
    f.kutu('txtSezon', 3950, r[0], 950, salt=True, hiza=2)
    f.alan('Kabul Tarihi', 'txtTarih', 300, r[1], 1700, 1500, bicim='dd\\.mm\\.yyyy', varsayilan='=Date()',
           sonra='=MalKabulTarihDegisti()')
    f.etiket('lblMusteri', 'Müşteri', 300, r[2] + 30, 1640)
    f.acilir('cmbMusteri', 2000, r[2], 2000, MUSTERI_SQL, tablo=True, sutun=2, genislikler='0;3200', listeGen=3500,
             sonra='=MalKabulMusteriSecildi()')
    f.dugme('btnYeniCari', '+ Yeni', 4050, r[2], 850, h=330, tik='=MalKabulYeniCari()', stil='altin', boyut=9)
    f.alan('Müşteri Bilgisi', 'txtMusteriBilgi', 300, r[3], 1700, 2900, salt=True, boyut=9)
    f.alanAcilir('Giriş Personeli', 'cmbGirisPersonel', 300, r[4], 1700, 2900, PERSONEL_SQL, tablo=True, sutun=2,
                 genislikler='0;2900')
    # B sütunu
    f.alanAcilir('Cinsi', 'cmbCinsi', 5300, r[0], 1500, 3200, 'Yağlık Zeytin;Sofralık Zeytin', sinirli=False)
    f.alanAcilir('Özellik', 'cmbOzellik', 5300, r[1], 1500, 3200, 'Dip;Tepe;Karma', sinirli=False)
    f.alanAcilir('Ödeme Tipi', 'cmbOdemeTipi', 5300, r[2], 1500, 3200, 'Hakyağı;Sıkım Ücreti')
    f.alan('Bidonlar', 'txtBidonDetay', 5300, r[3], 1500, 3200, sonra='=MalKabulBidonDuzenle()')
    f.alan('Açıklama', 'txtKabulAciklama', 5300, r[4], 1500, 3200)
    # C sütunu
    f.alan('Brüt Kg', 'txtBrut', 10500, r[0], 1800, 1600, sonra='=MalKabulHesapla()', hiza=3, bicim='General Number')
    f.alan('Dara Kg', 'txtDara', 10500, r[1], 1800, 1600, sonra='=MalKabulHesapla()', hiza=3, bicim='General Number')
    f.alan('Net Zeytin Kg', 'txtNet', 10500, r[2], 1800, 1600, salt=True, kalin=True, hiza=3, bicim='#,##0.0 "kg"')
    f.alan('Durum', 'txtDurum', 10500, r[3], 1800, 1600, salt=True)
    f.etiket('lblBidonIpucu', 'Bidon no: 11, 12, 13 veya 11-14 yazılabilir.', 10500, r[4] + 40, 4900, boyut=9,
             renk=GRI_METIN)
    f.kutu('txtPartiNo', 100, 250, 100, gorunur=False, tabstop=False)
    # Düğmeler
    x, t = 300, 3550
    for ad, metin, w, tik, stil in [
        ('btnYeni', 'Yeni Kayıt', 1650, '=MalKabulYeni()', 'altin'),
        ('btnKaydet', 'Kaydet', 1500, '=MalKabulKaydet()', 'yesil'),
        ('btnGuncelle', 'Güncelle', 1500, '=MalKabulGuncelle()', 'yesil'),
        ('btnSil', 'Sil', 1200, '=MalKabulSil()', 'kirmizi'),
        ('btnTemizle', 'Temizle', 1400, '=MalKabulTemizle()', 'gri'),
        ('btnFisOnizle', 'Fiş Önizle', 1650, '=MalKabulFisi("onizle")', 'acik'),
        ('btnFisYazdir', 'Fiş Yazdır', 1650, '=MalKabulFisi("yazdir")', 'acik'),
        ('btnKapat', 'Kapat', 1300, '=AktifFormuKapat()', 'gri')]:
        f.dugme(ad, metin, x, t, w, tik=tik, stil=stil)
        x += w + 100
    f.etiket('lblListe', 'Liste', 300, 4280, 700)
    f.acilir('cmbListeFiltre', 1000, 4250, 2600, 'Üretim Bekleyenler;Teslim Bekleyenler;Bu Sezon Tümü;Tüm Sezonlar',
             sonra='=MalKabulListeFiltrele()')
    f.etiket('lblIpucu', 'Listeden çift tıklayarak parti seçebilirsiniz.', 3800, 4290, 5500, boyut=9, renk=GRI_METIN)
    f.liste('lstPartiler', 300, 4700, 15200, 5300, 'SELECT * FROM qry_MalKabulListe WHERE 1=0;', 15,
            '0;900;1100;2600;2200;900;800;1000;1300;1700;1300;1100;0;0;0', ciftTik='=MalKabulListedenSec()')
    return f


# ---------------------------------------------------------------
def uretim():
    f = Form('frm_Uretim', 'Üretim', 15800, 10350, onload='=UretimYukle()', bant='ÜRETİM  (SIKIM)')
    r = satirlar(1250, 7)
    f.bolum('lblB1', 'PARTİ', 300, 850, 4000)
    f.bolum('lblB2', 'ÜRETİM', 5600, 850, 4000)
    f.bolum('lblB3', 'HESAP', 10600, 850, 2000)
    f.etiket('lblSezon', '', 12700, 860, 2800, kalin=True, renk=ALTIN, hiza=3, kontrol=False)
    f.etiket('lblParti', 'Parti', 300, r[0] + 30, 1640)
    f.acilir('cmbParti', 2000, r[0], 3000, 'SELECT PartiNo FROM tbl_Parti WHERE 1=0;', tablo=True, sutun=5,
             genislikler='0;900;2300;1100;1100', listeGen=5500, basliklar=True, sonra='=UretimPartiSecildi()', kalin=True)
    for i, (et, ad) in enumerate([('Müşteri', 'txtMusteri'), ('Kabul Tarihi', 'txtKabulTarihi'),
                                   ('Cinsi / Özellik', 'txtCinsi'), ('Net Zeytin (kg)', 'txtNetZeytin'),
                                   ('Bidonlar', 'txtBidonlar'), ('Durum', 'txtDurum')]):
        f.alan(et, ad, 300, r[i + 1], 1700, 3000, salt=True, bicim='#,##0.0 "kg"' if ad == 'txtNetZeytin' else None)
    f.alan('Sıkım Tarihi', 'txtSikimTarihi', 5600, r[0], 2250, 2150, bicim='dd\\.mm\\.yyyy')
    f.alan('Çıkan Yağ (kg)', 'txtCikanYag', 5600, r[1], 2250, 2150, sonra='=UretimHesapla()', kalin=True, hiza=3,
           bicim='General Number')
    f.alanAcilir('Ödeme Tipi', 'cmbOdemeTipi', 5600, r[2], 2250, 2150, 'Hakyağı;Sıkım Ücreti', sonra='=UretimHesapla()')
    f.alan('Hakyağı Oranı (%)', 'txtOran', 5600, r[3], 2250, 2150, sonra='=UretimHesapla()', hiza=3)
    f.alan('1 kg Yağ Fiyatı (TL)', 'txtYagFiyati', 5600, r[4], 2250, 2150, sonra='=UretimHesapla()', hiza=3,
           bicim='#,##0.00')
    f.etiket('lblFormul', 'Sıkım ücreti = çıkan yağ × oran × 1 kg yağ fiyatı', 5600, r[5] + 40, 4800, boyut=9,
             renk=GRI_METIN)
    for i, (et, ad) in enumerate([('Kızıl Su / Posa', 'txtPosa'), ('Randıman', 'txtRandiman'), ('Verim', 'txtVerim'),
                                   ('Hakyağı', 'txtHakyagi'), ('Sıkım Ücreti', 'txtUcret'),
                                   ('Müşteriye Kalan', 'txtKalanYag')]):
        f.alan(et, ad, 10600, r[i], 1800, 3100, salt=True, kalin=ad in ('txtKalanYag', 'txtUcret'))
    x, t = 300, 4400
    for ad, metin, w, tik, stil in [
        ('btnKaydet', 'Üretimi Kaydet', 2100, '=UretimKaydet()', 'yesil'),
        ('btnHesapla', 'Hesapla', 1500, '=UretimHesapla()', 'altin'),
        ('btnGeriAl', 'Üretimi Geri Al', 2200, '=UretimGeriAl()', 'kirmizi'),
        ('btnTemizle', 'Temizle', 1400, '=UretimTemizle()', 'gri'),
        ('btnKapat', 'Kapat', 1300, '=AktifFormuKapat()', 'gri')]:
        f.dugme(ad, metin, x, t, w, tik=tik, stil=stil)
        x += w + 100
    f.etiket('lblListe', 'Liste', 300, 5130, 700)
    f.acilir('cmbListeFiltre', 1000, 5100, 2600, 'Üretim Bekleyenler;Üretilenler;Bu Sezon Tümü', sonra='=UretimListeFiltrele()')
    f.etiket('lblIpucu', 'Listeden çift tıklayarak parti seçebilirsiniz.', 3800, 5140, 5500, boyut=9, renk=GRI_METIN)
    f.liste('lstUretim', 300, 5550, 15200, 4600, 'SELECT * FROM qry_UretimListe WHERE 1=0;', 17,
            '0;800;1050;2200;1100;1050;1000;1000;1250;700;1000;1100;1000;1200;0;0;0', ciftTik='=UretimListedenSec()')
    return f


# ---------------------------------------------------------------
def teslim():
    f = Form('frm_CicerTeslim', 'Yağ Teslim', 15800, 10500, onload='=TeslimFormYukle()', bant='YAĞ TESLİM')
    f.bolum('lblB1', 'PARTİ BİLGİLERİ', 300, 850, 4000)
    f.etiket('lblDurum', '', 5600, 860, 3900, kalin=True, renk=2785131, hiza=3, kontrol=False)
    f.bolum('lblB2', 'TESLİM BİLGİLERİ', 10000, 850, 4000)
    rA = satirlar(1250, 7, 400)
    f.etiket('lblParti', 'Parti', 300, rA[0] + 30, 1540)
    f.acilir('cmbParti', 1900, rA[0], 2600, 'SELECT PartiNo FROM tbl_Parti WHERE 1=0;', tablo=True, sutun=5,
             genislikler='0;900;2300;1100;1200', listeGen=5600, basliklar=True, sonra='=TeslimPartiSecildi()', kalin=True)
    for i, (et, ad) in enumerate([('Müşteri', 'txtMusteri'), ('Telefon', 'txtTelefon'), ('Kabul Tarihi', 'txtPartiTarihi'),
                                   ('Cinsi / Özellik', 'txtCinsi'), ('Ödeme Tipi', 'txtOdemeTipi'), ('Bidonlar', 'txtBidonlar')]):
        f.alan(et, ad, 300, rA[i + 1], 1600, 2600, salt=True)
    rB = satirlar(1250, 9, 400)
    for i, (et, ad) in enumerate([('Gelen Zeytin', 'txtZeytin'), ('Çıkan Yağ', 'txtYag'), ('Randıman', 'txtRandiman'),
                                   ('Hakyağı', 'txtHakyagi'), ('Kalan Net Yağ', 'txtNetYag'), ('Sıkım Ücreti', 'txtUcret'),
                                   ('Tahsil Edilen', 'txtOdenen'), ('Kalan Borç', 'txtKalanBorc'),
                                   ('Cari Bakiye', 'txtCariBakiye')]):
        f.alan(et, ad, 4800, rB[i], 1600, 2900, salt=True, kalin=ad in ('txtNetYag', 'txtKalanBorc'))
    rC = [1250, 1650, 2050, 2450, 2850]
    f.alan('Teslim Tarihi', 'txtTeslimTarihi', 10000, rC[0], 2000, 1500, bicim='dd\\.mm\\.yyyy')
    f.alan('Teslim Alan', 'txtTeslimAlan', 10000, rC[1], 2000, 3400)
    f.alanAcilir('Teslim Eden', 'cmbTeslimEden', 10000, rC[2], 2000, 3400, PERSONEL_SQL, tablo=True, sutun=2,
                 genislikler='0;3500')
    f.alan('Teslim Edilen Yağ', 'txtTeslimYag', 10000, rC[3], 2000, 1500, kalin=True, hiza=3, bicim='General Number')
    f.etiket('lblKg', 'kg', 13550, rC[3] + 30, 400)
    f.alan('Açıklama', 'txtTeslimAciklama', 10000, rC[4], 2000, 3400)
    f.bolum('lblB3', 'TAHSİLAT  (sıkım ücreti)', 10000, 3330, 5000)
    f.etiket('lblTahsilat', 'Tahsilat', 10000, 3730, 1940)
    f.acilir('cmbTahsilat', 12000, 3700, 1800, 'Tahsil Edilmedi;Peşin;Kısmi', sonra='=TahsilatAlanDurumu()')
    f.etiket('lblTutar', 'Tutar', 13850, 3730, 600)
    f.kutu('txtTahsilatTutar', 14450, 3700, 950, hiza=3, bicim='#,##0.00')
    f.etiket('lblYontem', 'Yöntem', 10000, 4130, 1940)
    f.acilir('cmbYontem', 12000, 4100, 1500, 'Nakit;Havale;Kart', sonra='=TahsilatYontemDegisti()')
    f.etiket('lblHesap', 'Kasa / Banka', 10000, 4530, 1940)
    f.acilir('cmbHesap', 12000, 4500, 3400, 'SELECT KasaID, KasaAdi FROM tbl_Kasa WHERE 1=0;', tablo=True, sutun=2,
             genislikler='0;3500')
    x, t = 300, 5100
    for ad, metin, w, tik, stil in [
        ('btnKaydet', 'Teslim Kaydet', 1950, '=TeslimKaydet()', 'yesil'),
        ('btnGeriAl', 'Teslimi Geri Al', 2050, '=TeslimGeriAl()', 'kirmizi'),
        ('btnTemizle', 'Temizle', 1350, '=TeslimTemizle()', 'gri'),
        ('btnOnizle', 'Fiş Önizle', 1650, '=TeslimFisi("onizle")', 'acik'),
        ('btnPdf', 'PDF Kaydet', 1750, '=TeslimFisi("pdf")', 'acik'),
        ('btnYazdir', 'Yazdır (A5)', 1700, '=TeslimFisi("yazdir")', 'acik'),
        ('btnKapat', 'Kapat', 1300, '=AktifFormuKapat()', 'gri')]:
        f.dugme(ad, metin, x, t, w, tik=tik, stil=stil)
        x += w + 100
    f.etiket('lblListe', 'Liste', 300, 5830, 700)
    f.acilir('cmbDurumFiltre', 1000, 5800, 2600, 'Teslim Bekleyenler;Teslim Edilenler;Borcu Kalanlar;Bu Sezon Tümü',
             sonra='=TeslimListeFiltrele()')
    f.etiket('lblIpucu', 'Listeden çift tıklayarak parti seçebilirsiniz.', 3800, 5840, 5000, boyut=9, renk=GRI_METIN)
    f.etiket('lblSezon', '', 11500, 5830, 4000, kalin=True, renk=ALTIN, hiza=3, kontrol=False)
    f.liste('lstTeslim', 300, 6250, 15200, 4100, 'SELECT * FROM qry_TeslimListe WHERE 1=0;', 17,
            '0;800;1050;2200;1250;1000;1150;1900;1150;1100;1100;1250;0;0;0;0;0', ciftTik='=TeslimListedenSec()')
    return f


# ---------------------------------------------------------------
def hizli_cari():
    f = Form('frm_HizliCari', 'Yeni Müşteri', 7200, 3900, onload='=HizliCariYukle()', popup=True, modal=True,
             kenarlik=3, minmax=False)
    f._ekle('Label', [('BackStyle', 1), ('OverlapFlags', 93), ('TextAlign', 1), ('Left', 0), ('Top', 0),
                      ('Width', 7200), ('Height', 560), ('FontSize', 13), ('FontWeight', 700), ('LeftMargin', 300),
                      ('TopMargin', 110), ('BackColor', KOYU), ('ForeColor', BEYAZ), ('Name', 'lblBaslik'),
                      ('Caption', q('YENİ MÜŞTERİ')), ('FontName', q('Segoe UI')), ('ThemeFontIndex', -1),
                      ('BackThemeColorIndex', -1), ('ForeThemeColorIndex', -1)], tab=False)
    f.cizgi('rctCizgi', 0, 560, 7200, 45, ALTIN)
    r = satirlar(800, 5, 410)
    f.alan('Ad Soyad', 'txtAdSoyad', 300, r[0], 1700, 4900, kalin=True)
    f.alan('Baba Adı', 'txtBabaAdi', 300, r[1], 1700, 2600)
    f.alan('TCKN / VKN', 'txtTCKN', 300, r[2], 1700, 2600)
    f.alan('Telefon', 'txtTelefon', 300, r[3], 1700, 2600)
    f.alan('Adres', 'txtAdres', 300, r[4], 1700, 4900)
    f.dugme('btnKaydet', 'Kaydet', 3700, 3000, 1650, tik='=HizliCariKaydet()', stil='yesil')
    f.dugme('btnVazgec', 'Vazgeç', 5450, 3000, 1450, tik='=HizliCariVazgec()', stil='gri')
    return f


# ---------------------------------------------------------------
def ana_menu():
    W = 15000
    f = Form('frm_CicerAnaMenu', 'Ana Menü', W, 7300, onload='=MenuYukle()')
    f._ekle('Rectangle', [('BackStyle', 1), ('OldBorderStyle', 0), ('OverlapFlags', 85), ('Left', 0), ('Top', 0),
                          ('Width', W), ('Height', 1500), ('BackColor', KOYU), ('Name', 'rctBant'), ('Tag', q('tema')),
                          ('BackThemeColorIndex', -1)], tab=False)
    f._ekle('LOGO_BUYUK', [('Name', 'imgLogo')], tab=False)
    f._ekle('TextBox', [('Enabled', 'NotDefault'), ('Locked', 'NotDefault'), ('TabStop', 'NotDefault'),
                        ('OldBorderStyle', 0), ('OverlapFlags', 215), ('TextAlign', 1), ('BackStyle', 0),
                        ('IMESentenceMode', 3), ('Left', 1750), ('Top', 300), ('Width', 9500), ('Height', 600),
                        ('FontSize', 22), ('FontWeight', 700), ('ForeColor', BEYAZ), ('Name', 'txtMarka'),
                        ('ControlSource', q('=FirmaAdiBuyuk()')), ('FontName', q('Segoe UI')), ('Tag', q('tema')),
                        ('ThemeFontIndex', -1), ('ForeThemeColorIndex', -1)])
    f.etiket('lblAltBaslik', 'ZEYTİNYAĞI ÜRETİM OTOMASYONU', 1800, 880, 7000, 400, 13, True, ALTIN, tag='tema')
    f.etiket('lblSezon', '', 10500, 950, 4200, 360, 11, True, ALTIN, hiza=3, tag='tema', kontrol=False)
    f.cizgi('rctCizgi', 0, 1500, W, 60, ALTIN)
    dugmeler = [
        [('btnMalKabul', 'Mal Kabul', 'frm_MalKabul', 'yesil'), ('btnUretim', 'Üretim', 'frm_Uretim', 'yesil'),
         ('btnTeslim', 'Yağ Teslim', 'frm_CicerTeslim', 'yesil'), ('btnMusteri', 'Müşteriler', 'frm_CicerMusteri', 'yesil'),
         ('btnCari', 'Cari Hareketler', 'frm_CicerCari', 'yesil')],
        [('btnKasa', 'Kasa / Banka', 'frm_CicerKasa', 'yesil'), ('btnStok', 'Stok Hareketleri', 'frm_CicerStok', 'yesil'),
         ('btnUrun', 'Ürün Kartları', 'frm_CicerUrun', 'yesil'), ('btnAlimSatim', 'Alım / Satım', 'frm_CicerAlimSatim', 'yesil'),
         ('btnRaporlar', 'Raporlar', 'frm_Raporlar', 'altin')],
        [('btnPersonel', 'Personel', 'frm_CicerPersonel', 'yesil'), ('btnKullanicilar', 'Kullanıcılar', 'frm_Kullanicilar', 'acik'),
         ('btnAyarlar', 'Ayarlar', 'frm_Ayarlar', 'acik')]]
    bw, bh, gap, x0, y0 = 2700, 850, 200, 350, 1850
    for i, satir in enumerate(dugmeler):
        for j, (ad, metin, form, stil) in enumerate(satir):
            f.dugme(ad, metin, x0 + j * (bw + gap), y0 + i * (bh + 150), bw, bh, tik='=MenuAc("%s")' % form,
                    stil=stil, boyut=12)
    y = y0 + 3 * (bh + 150) + 250
    f.etiket('lblOturum', '', 450, y, 6000, 330, 11, True, KOYU, kontrol=False)
    f.etiket('lblAltBilgi', 'Sürüm 3.0  •  Geliştirici: Ali Apaydın', 450, y + 350, 6000, 300, 10, renk=GRI_METIN)
    f.dugme('btnHakkinda', 'Hakkında', 8500, y, 1800, 600, tik='=MenuAc("frm_Hakkinda")', stil='altin')
    f.dugme('btnOturum', 'Oturumu Kapat', 10400, y, 2350, 600, tik='=OturumuKapat()', stil='gri')
    f.dugme('btnCikis', 'Çıkış', 12850, y, 1800, 600, tik='=AktifFormuKapat()', stil='gri')
    f.H = y + 900
    return f


# ---------------------------------------------------------------
def ayarlar():
    W = 15000
    f = Form('frm_Ayarlar', 'Ayarlar', W, 9200, onload='=AyarlarYukle()', bant='AYARLAR')
    f.bolum('lblB1', 'FİRMA BİLGİLERİ  (başlıklarda ve raporlarda görünür)', 300, 850, 7000)
    f.alan('Firma Adı', 'txtFirmaAdi', 300, 1250, 1900, 5000)
    f.alan('Telefon', 'txtFirmaTelefon', 300, 1650, 1900, 2600)
    f.alan('Adres', 'txtFirmaAdres', 300, 2050, 1900, 5000)
    f.bolum('lblB2', 'ÜRETİM VE SEZON', 300, 2600, 7000)
    f.alan('Hakyağı Oranı (%)', 'txtHakyagiOrani', 300, 3000, 2300, 1300, hiza=3)
    f.alan('1 kg Yağ Fiyatı (TL)', 'txtYagFiyati', 300, 3400, 2300, 1300, hiza=3, bicim='#,##0.00')
    f.etiket('lblAktifSezon', 'Aktif Sezon', 300, 3830, 2240)
    f.acilir('cmbAktifSezon', 2600, 3800, 1600, '')
    f.etiket('lblSezonNot', '(sezon 1 Eylül\'de başlar)', 4300, 3840, 3000, boyut=9, renk=GRI_METIN)
    f.etiket('lblNot1', 'Sıkım ücreti = çıkan yağ × hakyağı oranı × 1 kg yağ fiyatı.', 300, 4250, 6900, boyut=9,
             renk=GRI_METIN)
    f.etiket('lblNot2', 'Oran ve fiyat değişince yalnızca yeni açılan partilere uygulanır.', 300, 4500, 6900, boyut=9,
             renk=GRI_METIN)
    f.dugme('btnAyarKaydet', 'Ayarları Kaydet', 300, 4900, 2400, tik='=AyarlarKaydet()', stil='yesil')
    f.bolum('lblB3', 'YÖNETİCİ ŞİFRESİ', 300, 5600, 7000)
    f.alan('Mevcut Şifre', 'txtEskiSifre', 300, 6000, 2300, 2200, girisMaskesi='Password')
    f.alan('Yeni Şifre', 'txtYeniSifre', 300, 6400, 2300, 2200, girisMaskesi='Password')
    f.alan('Yeni Şifre (tekrar)', 'txtYeniSifre2', 300, 6800, 2300, 2200, girisMaskesi='Password')
    f.dugme('btnSifre', 'Şifreyi Değiştir', 5000, 6750, 2400, tik='=SifreDegistir()', stil='gri')
    # Sağ sütun: hesaplar
    X = 7800
    f.bolum('lblB4', 'KASA VE BANKA HESAPLARI', X, 850, 6900)
    f.liste('lstHesaplar', X, 1250, 6900, 1650, 'SELECT * FROM qry_HesapListe;', 6, '0;2100;800;1500;2500;0',
            ciftTik='=HesapListedenSec()')
    f.alan('Hesap Adı', 'txtHesapAdi', X, 3050, 1400, 3000)
    f.etiket('lblTur', 'Tür', 12350, 3080, 450)
    f.acilir('cmbHesapTuru', 12850, 3050, 1850, 'Nakit;Banka')
    f.alan('Banka', 'txtHesapBanka', X, 3450, 1400, 3000)
    f.alan('IBAN', 'txtHesapIBAN', X, 3850, 1400, 5500)
    f.kutu('txtHesapID', 100, 250, 100, gorunur=False, tabstop=False)
    f.dugme('btnHesapKaydet', 'Hesabı Kaydet', X, 4350, 2100, tik='=HesapKaydet()', stil='yesil')
    f.dugme('btnHesapSil', 'Hesabı Sil', X + 2200, 4350, 1800, tik='=HesapSil()', stil='kirmizi')
    f.dugme('btnHesapTemizle', 'Temizle', X + 4100, 4350, 1500, tik='=HesapTemizle()', stil='gri')
    f.bolum('lblB5', 'YEDEKLEME', X, 5000, 6900)
    f.dugme('btnYedekAl', 'Yedek Al', X, 5400, 3350, tik='=YedekEkraniAc()', stil='yesil')
    f.dugme('btnGeriYukle', 'Yedekten Geri Yükle', X + 3450, 5400, 3450, tik='=GeriYukleEkraniAc()', stil='altin')
    f.etiket('lblSonYedek', '', X, 5950, 6900, h=520, boyut=9, renk=GRI_METIN, kontrol=False)
    f.bolum('lblB6', 'YÖNETİCİ ARAÇLARI', X, 6600, 6900)
    f.dugme('btnKaynak', 'Kaynakları Dışa Aktar', X, 7000, 3350, tik='=KaynakAktarDugme()', stil='acik')
    f.dugme('btnYeniden', 'Kayıtları Yeniden Oluştur', X + 3450, 7000, 3450, tik='=KayitlariYenidenOlustur()', stil='gri')
    f.dugme('btnGelistirici', 'Geliştirici Modu (Shift)', X, 7550, 3350, tik='=GelistiriciModu()', stil='gri')
    f.etiket('lblVeriYolu', '', 300, 8550, 11800, boyut=9, renk=GRI_METIN, kontrol=False)
    f.dugme('btnKapat', 'Kapat', 13200, 8450, 1500, tik='=AktifFormuKapat()', stil='gri')
    return f


# ---------------------------------------------------------------
def kasa():
    f = Form('frm_CicerKasa', 'Kasa / Banka Hareketleri', 15800, 10200, onload='=KasaFormYukle()',
             bant='KASA / BANKA HAREKETLERİ')
    f.etiket('lblKasa', 'Hesap', 300, 930, 1100, kalin=True)
    f.acilir('cmbKasa', 1500, 900, 3600,
             "SELECT KasaID, KasaAdi & IIf(Nz(KasaTuru,'Nakit')='Banka','  (Banka)','') AS Hesap FROM tbl_Kasa "
             "WHERE Aktif=True ORDER BY IIf(Nz(KasaTuru,'Nakit')='Nakit',0,1), KasaAdi;",
             tablo=True, sutun=2, genislikler='0;3600', sonra='=KasaHesapDegisti()', kalin=True)
    f.etiket('lblBakiye', 'Hesap Bakiyesi', 5300, 930, 1620)
    f.kutu('txtKasaBakiye', 6950, 900, 1900, salt=True, kalin=True, hiza=3, bicim='#,##0.00')
    f.etiket('lblAcilis', 'Açılış Bakiyesi', 9150, 930, 1620)
    f.kutu('txtKasaAcilis', 10800, 900, 1500, hiza=3, bicim='#,##0.00')
    f.dugme('btnAcilis', 'Açılışı Kaydet', 12400, 870, 2000, h=400, tik='=KasaAcilisGuncelle()', stil='gri', boyut=10)
    f.cizgi('rctAyrac', 300, 1430, 15200)
    f.bolum('lblB1', 'HAREKET', 300, 1550)
    f.alan('Tarih', 'txtTarih', 300, 1950, 1200, 1500, bicim='dd\\.mm\\.yyyy')
    f.alanAcilir('Hareket', 'cmbHareketTuru', 3300, 1950, 1000, 1600, 'Tahsilat;Ödeme')
    f.alanAcilir('Yöntem', 'cmbOdemeYontemi', 6200, 1950, 1000, 1600, 'Nakit;Havale;Kart;EFT;Çek', sinirli=False)
    f.alan('Tutar', 'txtTutar', 9100, 1950, 800, 1600, hiza=3, bicim='#,##0.00', kalin=True)
    f.alanAcilir('Müşteri', 'cmbMusteri', 300, 2370, 1200, 3500, MUSTERI_SQL, tablo=True, sutun=2, genislikler='0;3500',
                 sonra='=KasaMusteriSecildi()')
    f.etiket('lblParti', 'Parti', 5300, 2400, 640)
    f.acilir('cmbParti', 6000, 2370, 2300, 'SELECT PartiNo FROM tbl_Parti WHERE 1=0;', tablo=True, sutun=4,
             genislikler='0;900;1100;1300', listeGen=3500, basliklar=True)
    f.alan('Açıklama', 'txtAciklama', 8600, 2370, 1100, 5800)
    f.etiket('lblIpucu', 'Sonradan gelen havaleyi borcu olan partiye bağlamak için müşteri ve parti seçin; tutar partinin kalan borcundan düşer.',
             300, 2790, 15000, boyut=9, renk=GRI_METIN)
    f.kutu('txtKasaHareketID', 100, 250, 100, gorunur=False, tabstop=False)
    x, t = 300, 3150
    for ad, metin, w, tik, stil in [
        ('btnYeni', 'Yeni Kayıt', 1600, '=KasaYeni()', 'altin'),
        ('btnKaydet', 'Kaydet', 1400, '=KasaHareketKaydet()', 'yesil'),
        ('btnGuncelle', 'Güncelle', 1500, '=KasaHareketGuncelle()', 'yesil'),
        ('btnSil', 'Sil', 1100, '=KasaHareketSil()', 'kirmizi'),
        ('btnTemizle', 'Temizle', 1400, '=KasaTemizle()', 'gri'),
        ('btnKapat', 'Kapat', 1300, '=AktifFormuKapat()', 'gri')]:
        f.dugme(ad, metin, x, t, w, tik=tik, stil=stil)
        x += w + 100
    f.alan('Başlangıç', 'txtBaslangic', 300, 3850, 1200, 1500, bicim='dd\\.mm\\.yyyy')
    f.alan('Bitiş', 'txtBitis', 3300, 3850, 700, 1500, bicim='dd\\.mm\\.yyyy')
    f.dugme('btnFiltre', 'Tarih Filtrele', 5700, 3810, 1900, h=410, tik='=KasaFiltrele()', stil='acik', boyut=10)
    f.dugme('btnTumu', 'Tümü', 7700, 3810, 1100, h=410, tik='=KasaTumu()', stil='gri', boyut=10)
    f.liste('lstHareketler', 300, 4400, 15200, 5600, 'SELECT * FROM qry_KasaEkstre WHERE 1=0;', 14,
            '0;1050;0;2200;700;1000;900;3900;1250;1250;1350;1100;0;0', ciftTik='=KasaListedenSec()')
    return f


# ---------------------------------------------------------------
# YEDEKLEME
# ---------------------------------------------------------------
KIRMIZI_METIN = 3158192
KAPANIS_KODU = """
Option Compare Database
Option Explicit

' Gizli form: Access penceresi X ile kapatilirken cikis sorusunu gosterir (bkz. modYedek)
Private Sub Form_Unload(Cancel As Integer)
    If Not CikisOnaylandi() Then
        Cancel = True
        Me.TimerInterval = 150
    End If
End Sub

Private Sub Form_Timer()
    Me.TimerInterval = 0
    CikisSor
End Sub
"""


def kucuk_baslik(f, metin):
    f._ekle('Label', [('BackStyle', 1), ('OverlapFlags', 93), ('TextAlign', 1), ('Left', 0), ('Top', 0),
                      ('Width', f.W), ('Height', 560), ('FontSize', 13), ('FontWeight', 700), ('LeftMargin', 300),
                      ('TopMargin', 110), ('BackColor', KOYU), ('ForeColor', BEYAZ), ('Name', 'lblBaslik'),
                      ('Caption', q(metin)), ('FontName', q('Segoe UI')), ('ThemeFontIndex', -1),
                      ('BackThemeColorIndex', -1), ('ForeThemeColorIndex', -1)], tab=False)
    f.cizgi('rctCizgi', 0, 560, f.W, 45, ALTIN)


def cikis():
    W = 7400
    f = Form('frm_Cikis', 'Çıkış', W, 2900, onload='=CikisFormYukle()', popup=True, modal=True, kenarlik=3,
             minmax=False)
    kucuk_baslik(f, 'ÇIKIŞ')
    f.etiket('lblSoru', 'Programdan çıkmak istiyor musunuz?', 400, 850, 6600, h=360, boyut=12, kalin=True)
    f.etiket('lblSonYedek', '', 400, 1300, 6600, h=520, boyut=9, renk=GRI_METIN, kontrol=False)
    f.dugme('btnEvet', 'Evet, Çık', 400, 2050, 1800, tik='=CikisEvet()', stil='kirmizi')
    f.dugme('btnYedek', 'Yedek Alıp Kapat', 2300, 2050, 2800, tik='=CikisYedekle()', stil='yesil')
    f.dugme('btnHayir', 'Hayır', 5200, 2050, 1800, tik='=CikisHayir()', stil='gri')
    return f


def yedek():
    W = 11400
    f = Form('frm_Yedek', 'Yedek Al', W, 7500, onload='=YedekFormYukle()', bant='YEDEK AL', popup=True,
             modal=True, kenarlik=3, minmax=False)
    f.etiket('lblAciklama', "Veri dosyası seçtiğiniz sürücüde 'ApaydinOtomasyon_Yedek' klasörüne tarihli adla "
             "kopyalanır. Klasör yoksa oluşturulur; eski yedekler silinmez.", 300, 850, 10800, h=520, boyut=9,
             renk=GRI_METIN, kontrol=False)
    f.bolum('lblB1', 'YEDEĞİN ALINACAĞI YER', 300, 1450, 6000)
    f.liste('lstSuruculer', 300, 1850, 10800, 2550, '', 6, '0;900;3300;2300;1700;1100', tik='=HedefGoster()',
            degerListesi=True, boyut=10)
    f.dugme('btnYenile', 'Sürücüleri Yenile', 300, 4550, 2500, h=400, tik='=SuruculeriListele()', stil='acik',
            boyut=10)
    f.dugme('btnKlasor', 'Klasör Seç...', 2900, 4550, 2000, h=400, tik='=YedekKlasorSec()', stil='gri', boyut=10)
    f.etiket('lblUyari', '', 5100, 4560, 6000, h=420, boyut=9, renk=KIRMIZI_METIN, kontrol=False)
    f.etiket('lblHedef', '', 300, 5100, 10800, h=760, boyut=9, kontrol=False)
    f.etiket('lblSonYedek', '', 300, 5950, 10800, h=300, boyut=9, renk=GRI_METIN, kontrol=False)
    f.dugme('btnYedekle', 'Yedeği Al', 6200, 6600, 3000, tik='=YedekAl()', stil='yesil')
    f.dugme('btnVazgec', 'Vazgeç', 9300, 6600, 1800, tik='=YedekVazgec()', stil='gri')
    return f


def geri_yukle():
    W = 12400
    f = Form('frm_GeriYukle', 'Yedekten Geri Yükle', W, 9000, onload='=GeriYukleFormYukle()',
             bant='YEDEKTEN GERİ YÜKLE', popup=True, modal=True, kenarlik=3, minmax=False)
    f.etiket('lblAciklama', "Takılı flaş bellekler, bu bilgisayardaki sürücüler ve seçtiğiniz klasörler "
             "'ApaydinOtomasyon_Yedek' klasörü için taranır. Yüklenen yedek mevcut verinin yerine geçer; mevcut "
             "veri silinmez, bu bilgisayardaki Yedek klasörüne taşınır.", 300, 850, 11800, h=520, boyut=9,
             renk=GRI_METIN, kontrol=False)
    f.bolum('lblB1', 'ÖNERİLEN YEDEK  (en güncel)', 300, 1450, 6000)
    f.cerceve('rctOneri', 300, 1850, 11800, 1050, arka=SALT)
    f.etiket('lblOneri', '', 500, 1950, 8500, h=400, boyut=12, kalin=True, kontrol=False)
    f.etiket('lblOneriYol', '', 500, 2380, 8500, h=480, boyut=9, renk=GRI_METIN, kontrol=False)
    f.dugme('btnOneriYukle', 'Bu Yedeği Yükle', 9200, 2100, 2700, tik='=OneriYukle()', stil='yesil')
    f.bolum('lblSayi', 'DİĞER YEDEKLER', 300, 3200, 6000)
    f.liste('lstYedekler', 300, 3600, 11800, 3700, '', 5, '0;1900;3400;4300;1100', ciftTik='=SeciliYukle()',
            degerListesi=True, boyut=10)
    f.dugme('btnSeciliYukle', 'Seçili Yedeği Yükle', 300, 7500, 3000, tik='=SeciliYukle()', stil='altin')
    f.dugme('btnTara', 'Yeniden Tara', 3400, 7500, 2000, tik='=YedekleriTara()', stil='gri')
    f.dugme('btnKlasor', 'Klasör Seç...', 5500, 7500, 2000, tik='=GeriYukleKlasorSec()', stil='gri')
    f.dugme('btnKapat', 'Kapat', 10500, 7500, 1600, tik='=GeriYukleKapat()', stil='gri')
    f.etiket('lblNot', 'Listeden daha eski bir yedek seçerseniz yüklemeden önce uyarı gösterilir.', 300, 8200,
             11800, boyut=9, renk=GRI_METIN, kontrol=False)
    return f


def kapanis():
    f = Form('frm_Kapanis', 'Kapanış', 3000, 600, kod=KAPANIS_KODU, olaylar=['OnUnload', 'OnTimer'])
    f.etiket('lblBilgi', 'Çıkış kontrolü', 100, 100, 2800, kontrol=False)
    return f


if __name__ == '__main__':
    for fn in (mal_kabul, uretim, teslim, hizli_cari, ana_menu, ayarlar, kasa, cikis, yedek, geri_yukle,
               kapanis):
        frm = fn()
        print(frm.yaz(HEDEF))
