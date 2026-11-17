# -*- coding: utf-8 -*-
"""GitHub okuma anahtarını (fine-grained token) modLisans'taki JETON_KODU için karıştırır.

Kullanım:  python jeton_gizle.py github_pat_xxx
Çıktıyı kaynak/vba/modLisans.bas içindeki JETON_KODU sabitine yapıştırın.
Bu bir şifreleme değildir; yalnızca anahtarın kaynakta düz metin olarak
durmasını (ve GitHub'ın onu otomatik iptal etmesini) önler. Anahtar yalnızca
app_licenses deposunu OKUMA yetkisine sahip olmalıdır.
"""
import sys

KARISTIR = "ApaydinOtomasyon"


def gizle(jeton):
    return ''.join('%02X' % (ord(c) ^ ord(KARISTIR[i % len(KARISTIR)])) for i, c in enumerate(jeton))


def coz(kod):
    return ''.join(chr(int(kod[i:i + 2], 16) ^ ord(KARISTIR[(i // 2) % len(KARISTIR)])) for i in range(0, len(kod), 2))


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    k = gizle(sys.argv[1].strip())
    assert coz(k) == sys.argv[1].strip()
    print(k)
