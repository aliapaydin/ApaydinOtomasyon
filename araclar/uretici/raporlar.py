# -*- coding: utf-8 -*-
"""rpt_MalKabulFisi: rpt_TeslimFisi'nin başlık/altlık bölümleriyle yeni ayrıntı bölümü"""
import os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from accessgen import q, KOYU, ALTIN, METIN, GRI_METIN, CIZGI, BEYAZ

R = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'kaynak', 'raporlar')
ETIKET_RENK = 6584440


class Detay:
    def __init__(self):
        self.k = []
        self.n = 0
        self.tab = 0

    def _ad(self, on):
        self.n += 1
        return '%s%d' % (on, self.n)

    def metin(self, kaynak, l, t, w, h=260, boyut=8, kalin=False, renk=METIN, hiza=None, buyu=False):
        o = []
        if buyu:
            o.append(('CanGrow', 'NotDefault'))
        o += [('OldBorderStyle', 0), ('OverlapFlags', 4)]
        if hiza:
            o.append(('TextAlign', hiza))
        o += [('BackStyle', 0), ('IMESentenceMode', 3), ('Left', l), ('Top', t), ('Width', w), ('Height', h),
              ('FontSize', boyut)]
        if kalin:
            o.append(('FontWeight', 700))
        o += [('TabIndex', self.tab), ('ForeColor', renk), ('Name', q(self._ad('Text'))),
              ('ControlSource', q(kaynak)), ('FontName', q('Segoe UI')),
              ('ThemeFontIndex', -1), ('ForeThemeColorIndex', -1)]
        self.tab += 1
        self.k.append(('TextBox', o))

    def kutu(self, l, t, w, h, arka=None, kenar=None):
        o = [('OldBorderStyle', 1 if kenar is not None else 0), ('OverlapFlags', 4)]
        if arka is not None:
            o.insert(0, ('BackStyle', 1))
        else:
            o.insert(0, ('BackStyle', 0))
        o += [('Left', l), ('Top', t), ('Width', w), ('Height', h)]
        if arka is not None:
            o.append(('BackColor', arka))
        if kenar is not None:
            o.append(('BorderColor', kenar))
        o += [('Name', q(self._ad('Box'))), ('BackThemeColorIndex', -1), ('BorderThemeColorIndex', -1)]
        self.k.append(('Rectangle', o))

    def bolum(self, metin, t):
        self.metin('="%s"' % metin, 0, t, 7200, 280, 9, True, ALTIN)
        self.kutu(0, t + 280, 7200, 12, arka=CIZGI)

    def cift(self, t, e1, k1, e2=None, k2=None):
        self.metin('="%s"' % e1, 0, t, 1250, renk=ETIKET_RENK)
        self.metin(k1, 1250, t, 2300 if e2 else 5950, kalin=True, buyu=e2 is None)
        if e2:
            self.metin('="%s"' % e2, 3650, t, 1250, renk=ETIKET_RENK)
            self.metin(k2, 4900, t, 2300, kalin=True)

    def blok(self):
        out = []
        for tur, o in self.k:
            out.append(' ' * 16 + 'Begin ' + tur)
            for k, v in o:
                out.append(' ' * 20 + ('%s = NotDefault' % k if v == 'NotDefault' else '%s =%s' % (k, v)))
            out.append(' ' * 16 + 'End')
        return '\n'.join(out) + '\n'


def mal_kabul_fisi():
    src = open(os.path.join(R, 'rpt_TeslimFisi.txt'), encoding='utf-8').read().replace('\r\n', '\n')
    src = re.sub(r'\n\s*GUID = Begin\n(?:\s*0x[0-9a-f]+\n)?\s*End', '', src)
    src = re.sub(r'\n    NameMap = Begin\n.*?\n    End', '', src, flags=re.S)
    src = re.sub(r'^Checksum =.*\n', '', src, flags=re.M)
    src = src.replace('Caption ="Yağ Teslim Fişi"', 'Caption ="Mal Kabul Fişi"')
    # Ayrıntı bölümünün kontrol listesi
    a = src.index('        Begin Section\n')
    b = src.index('            Begin\n', a) + len('            Begin\n')
    c = src.index('        Begin PageFooter')
    son = src.rindex('            End\n        End\n', a, c)

    d = Detay()
    d.kutu(0, 60, 7200, 420, arka=KOYU)
    d.metin('="MAL KABUL FİŞİ"', 120, 120, 3800, 320, 12, True, BEYAZ)
    d.metin('="Fiş No: " & [PartiKodu]', 4000, 140, 3080, 300, 9, True, ALTIN, hiza=3)
    d.bolum('MÜŞTERİ', 580)
    d.cift(910, 'Müşteri', '=[Musteri]', 'TCKN / VKN', '=Nz([TCKN_VKN],"-")')
    d.cift(1190, 'Telefon', '=Nz([Telefon],"-")', 'Adres', '=Nz([MusteriAdres],"-")')
    d.bolum('ZEYTİN', 1550)
    d.cift(1880, 'Parti No', '=[PartiKodu]', 'Kabul Tarihi', '=Format([Tarih],"dd.mm.yyyy")')
    d.cift(2160, 'Cinsi', '=Nz([Cinsi],"") & IIf(IsNull([Ozellik]),""," / " & [Ozellik])', 'Giriş Personeli',
           '=Nz([GirisPersoneli],"-")')
    d.cift(2440, 'Ödeme Tipi', '=[OdemeTipi]', 'Sezon', '=Nz([Sezon],"")')
    d.cift(2720, 'Bidonlar', '=Nz([BidonDetay],"-")')
    d.cift(3000, 'Açıklama', '=Nz([KabulAciklama],"")')
    d.bolum('TARTIM', 3360)
    d.cift(3690, 'Brüt', '=KgMetin(Nz([BrutKg],[ZeytinMiktari_Kg]))', 'Dara', '=KgMetin(Nz([DaraKg],0))')
    d.kutu(0, 4050, 7200, 520, arka=15921906)
    d.metin('="NET ZEYTİN:  " & KgMetin(Nz([ZeytinMiktari_Kg],0))', 120, 4150, 6960, 330, 12, True, KOYU)
    d.metin('=IIf([OdemeTipi] Like "S*","Sıkım ücreti: çıkan yağ x %" & Nz([HakYagiOrani_Yuzde],0) & " x " & '
            'Format(Nz([YagFiyati],0),"#,##0.00") & " TL (1 kg yağ fiyatı)","Hakyağı: çıkan yağın %" & '
            'Nz([HakYagiOrani_Yuzde],0) & " kadarı")', 0, 4680, 7200, 260, 8, False, GRI_METIN)
    for x, baslik, kim in [(0, 'TESLİM EDEN (MÜŞTERİ)', '=Nz([Musteri],"")'),
                           (3700, 'TESLİM ALAN', '=Nz([GirisPersoneli],"")')]:
        d.kutu(x, 5100, 3500, 1150, kenar=CIZGI)
        d.metin('="%s"' % baslik, x + 100, 5160, 3300, 260, 8, True, KOYU)
        d.metin(kim, x + 100, 5430, 3300, 260, 8)
        d.metin('="İmza"', x + 100, 5920, 3300, 240, 7, False, GRI_METIN)

    yeni = src[:b] + d.blok() + src[son:]
    yeni = re.sub(r'(        Begin Section\n(?:.*\n)*?            Height =)\d+', lambda m: m.group(1) + '6400', yeni, count=1)
    open(os.path.join(R, 'rpt_MalKabulFisi.txt'), 'w', encoding='utf-8', newline='').write(yeni.replace('\n', '\r\n'))
    print('rpt_MalKabulFisi yazıldı')


if __name__ == '__main__':
    mal_kabul_fisi()
