# -*- coding: utf-8 -*-
"""Access SaveAsText biçiminde form üretici (Apaydın Otomasyon teması)."""
import os, re

GEN = os.path.dirname(os.path.abspath(__file__))

# Renkler (Access BGR long)
KOYU = 2050619      # koyu zeytin (başlık bandı)
ALTIN = 2597577     # altın
METIN = 1520430     # koyu metin
GRI_METIN = 6584440
CIZGI = 8435888
ZEMIN = 14939127
SALT = 13691116     # salt okunur kutu zemini
BEYAZ = 16777215

BUTON = {
    'yesil':  (2785131, 1990736, 3446400, BEYAZ),
    'kirmizi': (2637462, 1978488, 3691186, BEYAZ),
    'gri':    (6058096, 4741722, 7506056, BEYAZ),
    'acik':   (4103820, 3179635, 5288608, BEYAZ),
    'altin':  (2597577, 1673130, 4110558, KOYU),
}


def q(s):
    """Access metin özelliği: çift tırnaklar \\" olur."""
    return '"' + s.replace('\\', '\\\\').replace('"', '\\"') + '"'


def metin_genislik(s, boyut=10, kalin=False):
    """Segoe UI için kaba genişlik tahmini (twip)."""
    w = 0.0
    for c in s:
        if c in 'iIl.,:;|!\'ı':
            w += 0.32
        elif c in 'fjtr()/ ':
            w += 0.42
        elif c.isupper() or c in 'mwMWĞŞÖÜÇİ%':
            w += 0.78
        else:
            w += 0.56
    w *= boyut * 20  # 1pt = 20 twip
    if kalin:
        w *= 1.08
    return int(w) + 60


class Form:
    def __init__(self, ad, baslik, genislik, yukseklik, onload=None, bant=None,
                 popup=False, modal=False, kenarlik=None, onclose=None, minmax=True):
        self.ad = ad
        self.baslik = baslik
        self.W = genislik
        self.H = yukseklik
        self.onload = onload
        self.onclose = onclose
        self.popup = popup
        self.modal = modal
        self.kenarlik = kenarlik
        self.minmax = minmax
        self.kontroller = []
        self.tab = 0
        self.uyarilar = []
        self.adlar = set()
        if bant:
            self.bant(bant)

    # ---------------------------------------------------------
    def _ekle(self, tur, ozl, tab=True):
        ad = [v for k, v in ozl if k == 'Name'][0]
        if ad in self.adlar:
            raise ValueError('Aynı ad: ' + ad)
        self.adlar.add(ad)
        if tab:
            ozl.append(('TabIndex', self.tab))
            self.tab += 1
        self.kontroller.append((tur, ozl))

    def bant(self, metin):
        """Üst koyu başlık bandı, logo, firma adı ve altın çizgi."""
        self._ekle('Label', [
            ('BackStyle', 1), ('OverlapFlags', 93), ('TextAlign', 1),
            ('Left', 0), ('Top', 0), ('Width', self.W), ('Height', 640),
            ('FontSize', 15), ('FontWeight', 700), ('LeftMargin', 950), ('TopMargin', 140),
            ('BackColor', KOYU), ('ForeColor', BEYAZ), ('Name', 'lblBaslik'),
            ('Caption', q(metin)), ('FontName', q('Segoe UI')), ('Tag', q('tema')),
            ('ThemeFontIndex', -1), ('BackThemeColorIndex', -1), ('ForeThemeColorIndex', -1)], tab=False)
        self.logo = True
        self._ekle('Rectangle', [
            ('BackStyle', 1), ('OldBorderStyle', 0), ('OverlapFlags', 87),
            ('Left', 0), ('Top', 640), ('Width', self.W), ('Height', 55),
            ('BackColor', ALTIN), ('Name', 'rctCizgi'), ('Tag', q('tema')),
            ('BackThemeColorIndex', -1)], tab=False)
        self._ekle('TextBox', [
            ('Enabled', 'NotDefault'), ('Locked', 'NotDefault'), ('TabStop', 'NotDefault'),
            ('OldBorderStyle', 0), ('OverlapFlags', 215), ('TextAlign', 3), ('BackStyle', 0),
            ('IMESentenceMode', 3), ('Left', self.W - 5700), ('Top', 210), ('Width', 5400), ('Height', 300),
            ('FontSize', 9), ('FontWeight', 700), ('ForeColor', ALTIN), ('Name', 'txtMarka'),
            ('ControlSource', q('=FirmaAdiBuyuk()')), ('FontName', q('Segoe UI')), ('Tag', q('tema')),
            ('ThemeFontIndex', -1), ('ForeThemeColorIndex', -1)])
        self._ekle('LOGO', [('Name', 'imgLogo')], tab=False)

    def etiket(self, ad, metin, l, t, w=None, h=300, boyut=10, kalin=False, renk=METIN,
               hiza=None, tag=None, gorunur=True, arka=None, kontrol=True):
        tahmin = metin_genislik(metin, boyut, kalin)
        if w is None:
            w = tahmin
        elif kontrol and tahmin > w:
            self.uyarilar.append('%s.%s: "%s" sığmayabilir (%d > %d)' % (self.ad, ad, metin, tahmin, w))
        o = []
        if not gorunur:
            o.append(('Visible', 'NotDefault'))
        if arka is not None:
            o.append(('BackStyle', 1))
        o.append(('OverlapFlags', 85))
        if hiza:
            o.append(('TextAlign', hiza))
        o += [('Left', l), ('Top', t), ('Width', w), ('Height', h), ('FontSize', boyut)]
        if kalin:
            o.append(('FontWeight', 700))
        if arka is not None:
            o.append(('BackColor', arka))
        o += [('ForeColor', renk), ('Name', ad), ('Caption', q(metin)), ('FontName', q('Segoe UI'))]
        if tag:
            o.append(('Tag', q(tag)))
        o += [('ThemeFontIndex', -1), ('ForeThemeColorIndex', -1)]
        if arka is not None:
            o.append(('BackThemeColorIndex', -1))
        self._ekle('Label', o, tab=False)

    def bolum(self, ad, metin, l, t, w=5000):
        self.etiket(ad, metin, l, t, w, 320, 10, True, ALTIN, tag='tema')

    def kutu(self, ad, l, t, w, h=330, salt=False, bicim=None, varsayilan=None, sonra=None,
             kalin=False, hiza=None, gorunur=True, kaynak=None, boyut=10, cokSatir=False,
             girisMaskesi=None, degisince=None, tabstop=True, odak_kayip=None):
        o = []
        if not gorunur:
            o.append(('Visible', 'NotDefault'))
        if salt:
            o += [('Locked', 'NotDefault'), ('TabStop', 'NotDefault')]
        elif not tabstop:
            o.append(('TabStop', 'NotDefault'))
        if cokSatir:
            o.append(('ScrollBars', 2))
        o.append(('OverlapFlags', 85))
        if hiza:
            o.append(('TextAlign', hiza))
        o += [('IMESentenceMode', 3), ('Left', l), ('Top', t), ('Width', w), ('Height', h), ('FontSize', boyut)]
        if kalin:
            o.append(('FontWeight', 700))
        if salt:
            o.append(('BackColor', SALT))
        o += [('BorderColor', CIZGI), ('ForeColor', METIN), ('Name', ad)]
        if kaynak:
            o.append(('ControlSource', q(kaynak)))
        if bicim:
            o.append(('Format', q(bicim)))
        if girisMaskesi:
            o.append(('InputMask', q(girisMaskesi)))
        if varsayilan:
            o.append(('DefaultValue', q(varsayilan)))
        if sonra:
            o.append(('AfterUpdate', q(sonra)))
        if degisince:
            o.append(('OnChange', q(degisince)))
        if odak_kayip:
            o.append(('OnLostFocus', q(odak_kayip)))
        o.append(('FontName', q('Segoe UI')))
        if cokSatir:
            o.append(('EnterKeyBehavior', 'NotDefault'))
        o += [('BackThemeColorIndex', -1), ('BorderThemeColorIndex', -1),
              ('ThemeFontIndex', -1), ('ForeThemeColorIndex', -1)]
        self._ekle('TextBox', o)

    def acilir(self, ad, l, t, w, kaynak, h=330, tablo=False, sutun=1, genislikler=None, sonra=None,
               listeGen=None, basliklar=False, sinirli=True, salt=False, gorunur=True, kalin=False):
        o = []
        if not gorunur:
            o.append(('Visible', 'NotDefault'))
        if basliklar:
            o.append(('ColumnHeads', 'NotDefault'))
        if sinirli:
            o.append(('LimitToList', 'NotDefault'))
        if not tablo:
            o.append(('RowSourceTypeInt', 1))
        if salt:
            o += [('Locked', 'NotDefault')]
        o += [('OverlapFlags', 85), ('IMESentenceMode', 3), ('ColumnCount', sutun)]
        if listeGen:
            o.append(('ListWidth', listeGen))
        o += [('Left', l), ('Top', t), ('Width', w), ('Height', h), ('FontSize', 10)]
        if kalin:
            o.append(('FontWeight', 700))
        o += [('BorderColor', CIZGI), ('ForeColor', METIN), ('Name', ad),
              ('RowSourceType', q('Table/Query' if tablo else 'Value List')),
              ('RowSource', q(kaynak))]
        if genislikler:
            o.append(('ColumnWidths', q(genislikler)))
        if sonra:
            o.append(('AfterUpdate', q(sonra)))
        o += [('FontName', q('Segoe UI')), ('AllowValueListEdits', 0),
              ('ThemeFontIndex', -1), ('BackThemeColorIndex', -1),
              ('BorderThemeColorIndex', -1), ('ForeThemeColorIndex', -1)]
        self._ekle('ComboBox', o)

    def liste(self, ad, l, t, w, h, kaynak, sutun, genislikler, ciftTik=None, tik=None, basliklar=True, boyut=9):
        o = []
        if basliklar:
            o.append(('ColumnHeads', 'NotDefault'))
        o += [('OverlapFlags', 85), ('IMESentenceMode', 3), ('ColumnCount', sutun),
              ('Left', l), ('Top', t), ('Width', w), ('Height', h), ('FontSize', boyut),
              ('ForeColor', METIN), ('BorderColor', CIZGI), ('Name', ad),
              ('RowSourceType', q('Table/Query')), ('RowSource', q(kaynak)),
              ('ColumnWidths', q(genislikler))]
        if ciftTik:
            o.append(('OnDblClick', q(ciftTik)))
        if tik:
            o.append(('AfterUpdate', q(tik)))
        o += [('FontName', q('Segoe UI')), ('ThemeFontIndex', -1), ('BackThemeColorIndex', -1),
              ('BorderThemeColorIndex', -1), ('ForeThemeColorIndex', -1)]
        self._ekle('ListBox', o)

    def onay(self, ad, l, t, sonra=None, metin=None, metinGen=None):
        o = [('OverlapFlags', 85), ('Left', l), ('Top', t + 30), ('Width', 260), ('Height', 240), ('Name', ad)]
        if sonra:
            o.append(('AfterUpdate', q(sonra)))
        o.append(('DefaultValue', q('False')))
        self._ekle('CheckBox', o)
        if metin:
            self.etiket('lbl_' + ad, metin, l + 330, t, metinGen)

    def dugme(self, ad, metin, l, t, w, h=450, tik=None, stil='yesil', boyut=11, gorunur=True):
        arka, kenar, uzerinde, yazi = BUTON[stil]
        tahmin = metin_genislik(metin, boyut, True) + 120
        if tahmin > w:
            self.uyarilar.append('%s.%s düğme "%s" sığmayabilir (%d > %d)' % (self.ad, ad, metin, tahmin, w))
        o = []
        if not gorunur:
            o.append(('Visible', 'NotDefault'))
        o += [('OverlapFlags', 85), ('Left', l), ('Top', t), ('Width', w), ('Height', h),
              ('FontSize', boyut), ('FontWeight', 700), ('ForeColor', yazi), ('Name', ad), ('Caption', q(metin))]
        if tik:
            o.append(('OnClick', q(tik)))
        o += [('FontName', q('Segoe UI')), ('ForeThemeColorIndex', -1), ('Gradient', 0),
              ('BackColor', arka), ('BackThemeColorIndex', -1), ('BorderColor', kenar),
              ('BorderThemeColorIndex', -1), ('ThemeFontIndex', -1), ('HoverColor', uzerinde),
              ('HoverThemeColorIndex', -1), ('PressedColor', kenar), ('PressedThemeColorIndex', -1),
              ('HoverForeColor', yazi), ('HoverForeThemeColorIndex', -1), ('PressedForeColor', yazi),
              ('PressedForeThemeColorIndex', -1)]
        self._ekle('CommandButton', o)

    def cizgi(self, ad, l, t, w, h=15, renk=CIZGI):
        self._ekle('Rectangle', [('BackStyle', 1), ('OldBorderStyle', 0), ('OverlapFlags', 85),
                                 ('Left', l), ('Top', t), ('Width', w), ('Height', h),
                                 ('BackColor', renk), ('Name', ad), ('BackThemeColorIndex', -1)], tab=False)

    def cerceve(self, ad, l, t, w, h, arka=None, kenarRenk=CIZGI):
        o = []
        if arka is not None:
            o.append(('BackStyle', 1))
        else:
            o.append(('BackStyle', 0))
        o += [('OverlapFlags', 85), ('Left', l), ('Top', t), ('Width', w), ('Height', h)]
        if arka is not None:
            o.append(('BackColor', arka))
        o += [('BorderColor', kenarRenk), ('Name', ad), ('BackThemeColorIndex', -1), ('BorderThemeColorIndex', -1)]
        self._ekle('Rectangle', o, tab=False)

    # Etiket + kutu satırı
    def alan(self, etiketMetni, ad, l, t, lw, w, **kw):
        self.etiket('lbl' + ad[3:], etiketMetni, l, t + 30, lw - 60)
        self.kutu(ad, l + lw, t, w, **kw)

    def alanAcilir(self, etiketMetni, ad, l, t, lw, w, kaynak, **kw):
        self.etiket('lbl' + ad[3:], etiketMetni, l, t + 30, lw - 60)
        self.acilir(ad, l + lw, t, w, kaynak, **kw)

    # ---------------------------------------------------------
    def _blok(self, tur, ozl):
        g = ' ' * 20
        satir = [' ' * 16 + 'Begin ' + tur]
        for k, v in ozl:
            if k == 'Name' and not str(v).startswith('"'):
                v = q(v)
            if v == 'NotDefault':
                satir.append(g + '%s = NotDefault' % k)
            else:
                satir.append(g + '%s =%s' % (k, v))
        satir.append(' ' * 16 + 'End')
        return '\n'.join(satir) + '\n'

    def metin(self):
        out = []
        out.append('Version =21\nVersionRequired =20\nPublishOption =1\nBegin Form\n')
        out.append('    RecordSelectors = NotDefault\n')
        out.append('    AutoCenter = NotDefault\n')
        out.append('    NavigationButtons = NotDefault\n')
        out.append('    DividingLines = NotDefault\n')
        out.append('    AllowDesignChanges = NotDefault\n')
        if self.popup:
            out.append('    PopUp = NotDefault\n')
        if self.modal:
            out.append('    Modal = NotDefault\n')
        out.append('    DefaultView =0\n')
        out.append('    ScrollBars =0\n')
        if self.kenarlik is not None:
            out.append('    BorderStyle =%d\n' % self.kenarlik)
        if not self.minmax:
            pass
        out.append('    TabularCharSet =162\n    PictureAlignment =2\n    DatasheetGridlinesBehavior =3\n    GridY =10\n')
        out.append('    Width =%d\n' % self.W)
        out.append('    DatasheetFontHeight =11\n    ItemSuffix =%d\n' % (len(self.kontroller) + 10))
        out.append('    DatasheetGridlinesColor =15132391\n')
        out.append('    Caption =%s\n' % q(self.baslik))
        out.append('    DatasheetFontName ="Calibri"\n')
        if self.onload:
            out.append('    OnLoad =%s\n' % q(self.onload))
        if self.onclose:
            out.append('    OnClose =%s\n' % q(self.onclose))
        out.append('    FilterOnLoad =0\n    ShowPageMargins =0\n    DisplayOnSharePointSite =1\n')
        out.append('    DatasheetAlternateBackColor =15921906\n    DatasheetGridlinesColor12 =0\n    FitToScreen =1\n')
        out.append('    DatasheetBackThemeColorIndex =1\n    BorderThemeColorIndex =3\n    ThemeFontIndex =1\n')
        out.append('    ForeThemeColorIndex =0\n    AlternateBackThemeColorIndex =1\n    AlternateBackShade =95.0\n')
        out.append('    NoSaveCTIWhenDisabled =1\n')
        out.append(open(os.path.join(GEN, 'form_defaults.txt'), encoding='utf-8').read())
        out.append('        Begin Section\n')
        out.append('            Height =%d\n            BackColor =%d\n            Name ="Ayrıntı"\n' % (self.H, ZEMIN))
        out.append('            AutoHeight =1\n            AlternateBackColor =15921906\n')
        out.append('            AlternateBackThemeColorIndex =1\n            AlternateBackShade =95.0\n')
        out.append('            Begin\n')
        for tur, ozl in self.kontroller:
            if tur == 'LOGO':
                logo = open(os.path.join(GEN, 'logo_kucuk.txt'), encoding='utf-8').read()
                logo = re.sub(r'\n\s*GUID = Begin\n.*?\n\s*End', '', logo, count=1, flags=re.S)
                logo = re.sub(r'TabIndex =\d+', 'TabIndex =%d' % self.tab, logo)
                self.tab += 1
                out.append(logo)
            elif tur == 'LOGO_BUYUK':
                logo = open(os.path.join(GEN, 'logo_buyuk.txt'), encoding='utf-8').read()
                logo = re.sub(r'\n\s*GUID = Begin\n.*?\n\s*End', '', logo, count=1, flags=re.S)
                logo = re.sub(r'TabIndex =\d+', 'TabIndex =%d' % self.tab, logo)
                self.tab += 1
                out.append(logo)
            else:
                out.append(self._blok(tur, ozl))
        out.append('            End\n        End\n    End\nEnd\n')
        return ''.join(out).replace('\n', '\r\n')

    def yaz(self, klasor):
        yol = os.path.join(klasor, self.ad + '.txt')
        with open(yol, 'w', encoding='utf-8', newline='') as f:
            f.write(self.metin())
        for u in self.uyarilar:
            print('UYARI', u)
        return yol
