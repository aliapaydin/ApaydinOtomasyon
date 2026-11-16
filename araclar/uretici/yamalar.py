# -*- coding: utf-8 -*-
"""Var olan formlara küçük değişiklikler:  python3 yamalar.py"""
import os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import accessgen as g

K = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'kaynak', 'formlar')


def oku(ad):
    return open(os.path.join(K, ad + '.txt'), encoding='utf-8').read().replace('\r\n', '\n')


def yaz(ad, s):
    open(os.path.join(K, ad + '.txt'), 'w', encoding='utf-8', newline='').write(s.replace('\n', '\r\n'))


def blok_bul(s, ad):
    m = re.search(r'\n( {16}Begin \w+\n(?:(?! {16}End\n).*\n)*? {20}Name ="%s"\n(?:(?! {16}End\n).*\n)*? {16}End\n)' % re.escape(ad), s)
    if not m:
        raise KeyError(ad)
    return m.start(1), m.end(1)


def ozellik(s, ad, k, v):
    a, b = blok_bul(s, ad)
    blok = s[a:b]
    if v is None:
        blok2 = re.sub(r'\n {20}%s =[^\n]*' % k, '', blok)
    elif re.search(r'\n {20}%s =' % k, blok):
        blok2 = re.sub(r'(\n {20}%s =)[^\n]*' % k, lambda m: m.group(1) + str(v), blok)
        # LayoutCached değerlerini de güncelle
    else:
        blok2 = blok.replace('\n' + ' ' * 20 + 'Name =', '\n' + ' ' * 20 + '%s =%s\n' % (k, v) + ' ' * 20 + 'Name =', 1)
    if k in ('Left', 'Top', 'Width', 'Height'):
        blok2 = re.sub(r'\n {20}LayoutCached\w+ =[^\n]*', '', blok2)
    return s[:a] + blok2 + s[b:]


def en_buyuk_tab(s):
    return max([int(x) for x in re.findall(r'TabIndex =(\d+)', s)] + [0])


def ekle(s, form_nesnesi):
    """Geçici bir Form nesnesindeki kontrolleri (ilk tab sırasını düzelterek) bölümün sonuna ekler."""
    t = en_buyuk_tab(s) + 1
    parcalar = []
    for tur, ozl in form_nesnesi.kontroller:
        yeni = []
        for k, v in ozl:
            if k == 'TabIndex':
                v = t
                t += 1
            yeni.append((k, v))
        parcalar.append(form_nesnesi._blok(tur, yeni).replace('\r\n', '\n'))
    son = '            End\n        End\n    End\nEnd'
    s = s.rstrip('\n')
    assert s.endswith(son), 'beklenmeyen form sonu'
    return s[:-len(son)] + ''.join(parcalar) + son + '\n'


def musteri():
    s = oku('frm_CicerMusteri')
    if 'txtBabaAdi' in s:
        return
    f = g.Form('x', 'x', 100, 100)
    f.etiket('lblBabaAdi', 'Baba Adı', 300, 2950, 1700)
    f.kutu('txtBabaAdi', 2100, 2850, 2300)
    s = ekle(s, f)
    s = ozellik(s, 'lstMusteriler', 'ColumnCount', 9)
    s = ozellik(s, 'lstMusteriler', 'ColumnWidths', '"0;1300;2700;1400;1400;1400;2900;1200;1300"')
    yaz('frm_CicerMusteri', s)


def raporlar():
    s = oku('frm_Raporlar')
    if 'cmbRSezon' in s:
        return
    for ad, l, w in [('btnBugun', 5900, 1100), ('btnHafta', 7100, 1200), ('btnAy', 8400, 1000), ('btnYil', 9500, 1000)]:
        s = ozellik(s, ad, 'Left', l)
        s = ozellik(s, ad, 'Width', w)
    f = g.Form('x', 'x', 100, 100)
    f.etiket('lblRSezon', 'Sezon', 10600, 2640, 650)
    f.acilir('cmbRSezon', 11250, 2620, 1450, '', sonra='=RaporSezonSecildi()')
    s = ekle(s, f)
    yaz('frm_Raporlar', s)


# Etiketlerin sığıp sığmadığını raporlar (yalnızca bilgi)
def etiket_denetimi():
    for fn in sorted(os.listdir(K)):
        s = oku(fn[:-4])
        for m in re.finditer(r' {16}Begin Label\n((?:(?! {16}End\n).*\n)*?) {16}End\n', s):
            b = m.group(1)
            cap = re.search(r'Caption ="((?:[^"\\]|\\.)*)"', b)
            w = re.search(r'\n {20}Width =(\d+)', '\n' + b)
            fs = re.search(r'FontSize =(\d+)', b)
            fw = 'FontWeight =700' in b
            if not cap or not w:
                continue
            tahmin = g.metin_genislik(cap.group(1), int(fs.group(1)) if fs else 11, fw)
            if tahmin > int(w.group(1)) + 80:
                ad = re.search(r'Name ="([^"]+)"', b).group(1)
                print('%-24s %-14s %5d > %5d  %s' % (fn, ad, tahmin, int(w.group(1)), cap.group(1)))


def etiketler():
    """Kesik görünen etiketler"""
    for form, ad, k, v in [
            ('frm_CicerAlimSatim', 'Etiket17', 'Width', 3200),
            ('frm_CicerCari', 'lblCariTCKN', 'Left', 3300),
            ('frm_CicerCari', 'lblCariTCKN', 'Width', 1200),
            ('frm_Kullanicilar', 'lblNot', 'Width', 5250),
            ('frm_CicerMusteri', 'Etiket6', 'Width', 1780),
            ('frm_Hakkinda', 'lblUrun', 'Width', 7000)]:
        s = oku(form)
        s = ozellik(s, ad, k, v)
        yaz(form, s)


if __name__ == '__main__':
    musteri()
    raporlar()
    etiketler()
    etiket_denetimi()
