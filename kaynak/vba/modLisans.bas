Option Compare Database
Option Explicit

' =============================================================
' KULLANIM (LİSANS) KONTROLÜ   (form: frm_Kilit)
' -------------------------------------------------------------
' Program her açılışta (ve oturum kapatılıp giriş ekranına
' dönüldüğünde) GitHub deposundaki lisans.txt dosyasını okur:
'
'   HERKES = acik                 tüm kurulumlar için durum
'   AO-7K3M9P = kapali | mesaj    tek kurulum için durum (HERKES'in yerine geçer)
'   MESAJ = ...                   program kapalıyken gösterilecek mesaj
'   ILETISIM = ...                mesajın altındaki iletişim bilgisi
'
' Lisans kodu her kurulumda ilk açılışta üretilir ve Ayarlar ekranında
' görünür. Okunan durum veri dosyasına (tbl_Ayarlar) yazılır:
'   - Dosyaya ulaşılamazsa son bilinen durum geçerlidir; kapatılmış bir
'     kurulum internet kesilince açılmaz.
'   - 7 gün boyunca hiç kontrol yapılamazsa giriş durdurulur.
' Program kapalıyken yalnızca "Tekrar Dene", "Yedek Al" ve "Çıkış"
' kullanılabilir; müşteri kendi verisinin yedeğini her zaman alabilir.
' =============================================================

Private Const DEPO As String = "aliapaydin/ApaydinOtomasyon"
Private Const DAL As String = "main"
Private Const CEVRIMDISI_GUN As Long = 7
Private Const ZAMAN_ASIMI As Single = 6

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private LisansAcildi As Boolean
Private SonOkumaBasarili As Boolean

' ---------------------------------------------------------------
' Yardımcılar
' ---------------------------------------------------------------
Private Function IsoSimdi() As String
    IsoSimdi = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function

Private Function IsoTarihOku(ByVal s As String) As Date
    On Error Resume Next
    IsoTarihOku = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2))) + _
                  TimeSerial(CInt(Mid$(s, 12, 2)), CInt(Mid$(s, 15, 2)), CInt(Mid$(s, 18, 2)))
End Function

Private Sub AyarKisa(ByVal anahtar As String, ByVal deger As String)
    AyarYaz anahtar, Left$(deger, 250)
End Sub

Public Function LisansKodu() As String
    Dim k As String
    k = Ayar("LisansKodu")
    If k = "" Then
        k = YeniLisansKodu()
        AyarYaz "LisansKodu", k
    End If
    LisansKodu = k
End Function

Private Function YeniLisansKodu() As String
    Dim harfler As String, i As Long, s As String
    harfler = "ABCDEFGHJKLMNPRSTUVYZ23456789"
    Randomize
    For i = 1 To 6
        s = s & Mid$(harfler, Int(Rnd * Len(harfler)) + 1, 1)
    Next
    YeniLisansKodu = "AO-" & s
End Function

' Lisans dosyasını indirir. Önce GitHub API (en fazla 1 dk önbellek), olmazsa
' raw.githubusercontent.com (birkaç dk önbellek) denenir.
Private Function UzakMetin(ByRef metin As String) As Boolean
    If AdrestenOku("https://api.github.com/repos/" & DEPO & "/contents/lisans.txt?ref=" & DAL, _
                   "application/vnd.github.raw", metin) Then
        UzakMetin = True
    ElseIf AdrestenOku("https://raw.githubusercontent.com/" & DEPO & "/" & DAL & "/lisans.txt", "", metin) Then
        UzakMetin = True
    End If
End Function

' Windows'un internet / proxy ayarlarıyla okur, en fazla 6 sn bekler
Private Function AdrestenOku(ByVal adres As String, ByVal kabul As String, ByRef metin As String) As Boolean
    On Error GoTo Hata
    Dim h As Object, bas As Single
    Set h = CreateObject("MSXML2.XMLHTTP.6.0")
    h.Open "GET", adres & IIf(InStr(1, adres, "?") > 0, "&", "?") & "t=" & Format$(Now, "yyyymmddhhnnss"), True
    h.setRequestHeader "Cache-Control", "no-cache"
    h.setRequestHeader "Pragma", "no-cache"
    If kabul <> "" Then h.setRequestHeader "Accept", kabul
    h.send
    bas = Timer
    Do While h.readyState <> 4
        DoEvents
        Sleep 50
        If Timer - bas > ZAMAN_ASIMI Or Timer < bas Then
            h.abort
            Exit Function
        End If
    Loop
    If h.Status = 200 Then
        metin = h.responseText
        AdrestenOku = True
    End If
    Exit Function
Hata:
End Function

Private Function NormalDurum(ByVal s As String) As String
    s = Replace(Replace(Replace(s, "ı", "i"), "İ", "i"), "ç", "c")
    s = LCase$(Trim$(s))
    If s = "kapali" Or s = "kapat" Or s = "kapalidir" Then
        NormalDurum = "kapali"
    Else
        NormalDurum = "acik"
    End If
End Function

' lisans.txt içeriğini bu kurulumun koduna göre çözer
Private Function LisansCoz(ByVal metin As String, ByVal kod As String, ByRef durum As String, _
                           ByRef mesaj As String, ByRef iletisim As String) As Boolean
    Dim satirlar() As String, i As Long, s As String, p As Long, anahtar As String, deger As String
    Dim genel As String, ozel As String, genelVar As Boolean, ozelVar As Boolean
    metin = Replace(metin, vbCr, "")
    satirlar = Split(metin, vbLf)
    For i = 0 To UBound(satirlar)
        s = Trim$(satirlar(i))
        If s <> "" And Left$(s, 1) <> "#" Then
            p = InStr(1, s, "=")
            If p > 1 Then
                anahtar = UCase$(Trim$(Left$(s, p - 1)))
                deger = Trim$(Mid$(s, p + 1))
                If anahtar = "HERKES" Then
                    genel = deger: genelVar = True
                ElseIf anahtar = "MESAJ" Then
                    mesaj = deger
                ElseIf anahtar = "ILETISIM" Or anahtar = "İLETİŞİM" Then
                    iletisim = deger
                ElseIf anahtar = UCase$(kod) Then
                    ozel = deger: ozelVar = True
                End If
            End If
        End If
    Next
    If Not genelVar Then Exit Function   ' beklenen biçimde değil (ör. hata sayfası)
    If ozelVar Then s = ozel Else s = genel
    p = InStr(1, s, "|")
    If p > 0 Then
        If Trim$(Mid$(s, p + 1)) <> "" Then mesaj = Trim$(Mid$(s, p + 1))
        s = Left$(s, p - 1)
    End If
    durum = NormalDurum(s)
    LisansCoz = True
End Function

' ---------------------------------------------------------------
' Kontrol
' ---------------------------------------------------------------
' Dosyayı okuyup sonucu veri dosyasına yazar. Okunabildiyse True.
Public Function LisansGuncelle() As Boolean
    On Error GoTo Hata
    Dim metin As String, durum As String, mesaj As String, iletisim As String
    If Ayar("LisansIlkCalisma") = "" Then AyarYaz "LisansIlkCalisma", IsoSimdi()
    SonOkumaBasarili = False
    If Not UzakMetin(metin) Then Exit Function
    If Not LisansCoz(metin, LisansKodu(), durum, mesaj, iletisim) Then Exit Function
    AyarKisa "LisansDurum", durum
    AyarKisa "LisansMesaj", mesaj
    AyarKisa "LisansIletisim", iletisim
    AyarYaz "LisansSonKontrol", IsoSimdi()
    SonOkumaBasarili = True
    LisansGuncelle = True
    Exit Function
Hata:
End Function

' Son bilinen duruma göre program açık mı?
Public Function LisansAcik() As Boolean
    Dim son As String
    If Ayar("LisansDurum") = "kapali" Then Exit Function
    son = Ayar("LisansSonKontrol")
    If son = "" Then son = Ayar("LisansIlkCalisma")
    If son = "" Then
        LisansAcik = True
        Exit Function
    End If
    If DateDiff("d", IsoTarihOku(son), Now) > CEVRIMDISI_GUN Then Exit Function
    LisansAcik = True
End Function

' Giriş ekranı açılırken çağrılır. False dönerse program kapanır.
Public Function LisansKontrol() As Boolean
    On Error GoTo Hata
    If TestModu Then
        LisansKontrol = True
        Exit Function
    End If
    LisansKodu
    LisansGuncelle
    If LisansAcik() Then
        LisansKontrol = True
        Exit Function
    End If
    LisansAcildi = False
    DoCmd.OpenForm "frm_Kilit", WindowMode:=acDialog
    LisansKontrol = LisansAcildi
    Exit Function
Hata:
    ' Beklenmedik bir hata müşteriyi programsız bırakmasın
    LisansKontrol = True
End Function

Private Function SonKontrolMetni() As String
    Dim son As String
    son = Ayar("LisansSonKontrol")
    If son = "" Then
        SonKontrolMetni = "henüz yapılamadı"
    Else
        SonKontrolMetni = Format$(IsoTarihOku(son), "dd\.mm\.yyyy hh:nn")
    End If
End Function

Public Function LisansBilgiMetni() As String
    LisansBilgiMetni = "Lisans kodu: " & LisansKodu() & "   •   Durum: " & _
                       IIf(LisansAcik(), "açık", "kapalı") & "   •   Son kontrol: " & SonKontrolMetni()
End Function

' ---------------------------------------------------------------
' Kilit ekranı   (form: frm_Kilit)
' ---------------------------------------------------------------
Public Function KilitFormYukle()
    On Error Resume Next
    Dim frm As Form, mesaj As String, iletisim As String
    Set frm = Forms("frm_Kilit")
    If Ayar("LisansDurum") = "kapali" Then
        mesaj = Ayar("LisansMesaj")
        If mesaj = "" Then mesaj = "Programın kullanımı durdurulmuştur. Lütfen yazılımcı ile iletişime geçin."
    Else
        mesaj = "Kullanım durumu " & CEVRIMDISI_GUN & " gündür kontrol edilemedi." & vbCrLf & _
                "Bilgisayarın internete bağlı olduğundan emin olup 'Tekrar Dene'ye basın."
    End If
    iletisim = Ayar("LisansIletisim")
    If iletisim = "" Then iletisim = "Yazılımcı: " & GELISTIRICI & "   •   " & GELISTIRICI_TELEFON & "   •   " & GELISTIRICI_EPOSTA
    frm!lblMesaj.Caption = mesaj
    frm!lblIletisim.Caption = iletisim
    frm!lblKod.Caption = "Lisans kodu: " & LisansKodu() & "   (yazılımcıya bu kodu bildirin)"
    frm!lblDurum.Caption = "Son kontrol: " & SonKontrolMetni() & _
                           IIf(SonOkumaBasarili, "", "   •   Bu açılışta lisans dosyasına ulaşılamadı.")
End Function

Public Function KilitTekrarDene()
    On Error Resume Next
    DoCmd.Hourglass True
    LisansGuncelle
    DoCmd.Hourglass False
    If LisansAcik() Then
        LisansAcildi = True
        Mesaj "Program kullanıma açıldı."
        DoCmd.Close acForm, "frm_Kilit"
        Exit Function
    End If
    KilitFormYukle
    If Not SonOkumaBasarili Then
        Mesaj "Lisans dosyasına ulaşılamadı. İnternet bağlantısını kontrol edin.", vbExclamation
    Else
        Mesaj "Program hâlâ kullanıma kapalı. Lütfen yazılımcı ile iletişime geçin.", vbExclamation
    End If
End Function

Public Function KilitYedekAl()
    On Error Resume Next
    DoCmd.OpenForm "frm_Yedek", WindowMode:=acDialog
End Function

Public Function KilitCikis()
    LisansAcildi = False
    DoCmd.Close acForm, "frm_Kilit"
End Function
