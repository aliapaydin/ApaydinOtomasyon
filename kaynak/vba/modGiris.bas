Option Compare Database
Option Explicit

' =============================================================
' GİRİŞ, KULLANICILAR ve AYARLAR
' -------------------------------------------------------------
' frm_Giris        : profil kartlarıyla giriş (kartı seç, şifreyi yaz)
' frm_Kullanicilar : yalnızca yönetici; kullanıcı ekler / siler,
'                    şifrelerini belirler
' frm_Ayarlar      : yalnızca yönetici; firma adı ve bilgileri,
'                    varsayılan ücret / oran, yönetici şifresi
' İlk kurulumda "admin" kullanıcısı şifre 1234 ile oluşturulur.
' Şifreler tbl_Kullanici tablosunda tuzlu özet olarak saklanır.
' =============================================================

Private Const PROFIL_SAYISI As Integer = 5      ' bir sayfada gösterilen kart
Private Const RENK_SAYISI As Integer = 6        ' imgRenk0..imgRenk5

Private Profiller() As Long                     ' sıralı KullaniciID listesi
Private ProfilAdet As Long
Private Sayfa As Long
Private SeciliID As Long

' ---------------------------------------------------------------
' Kurulum: hiç kullanıcı yoksa admin / 1234 oluşturur
' ---------------------------------------------------------------
Public Function KullaniciVarsayilanOlustur()
    Dim rs As DAO.Recordset
    If DCount("*", "tbl_Kullanici") > 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("tbl_Kullanici", dbOpenDynaset)
    rs.AddNew
    rs!KullaniciAdi = "admin"
    rs!adSoyad = "Yönetici"
    rs!sifre = SifreSakla("1234")
    rs!Yonetici = True
    rs!Aktif = True
    rs!Renk = 0
    rs!OlusturmaTarihi = Now
    rs.Update
    rs.Close
End Function

Private Function bas(ByVal adSoyad As String) As String
    Dim p() As String, s As String
    p = Split(Trim$(adSoyad), " ")
    If UBound(p) < 0 Then bas = "?": Exit Function
    s = Left$(p(0), 1)
    If UBound(p) > 0 Then s = s & Left$(p(UBound(p)), 1)
    bas = BuyukHarf(s)
End Function

' ---------------------------------------------------------------
' GİRİŞ EKRANI
' ---------------------------------------------------------------
Private Function G() As Form
    Set G = Forms(GIRIS_FORMU)
End Function

Public Function GirisYukle()
    On Error GoTo Hata
    Dim rs As DAO.Recordset
    ' Yedekten geri yükleme bekliyorsa önce onu yap (veri tablolarına dokunmadan)
    BekleyenGeriYukleme
    ' Veri dosyası bağlantısı ve sürümü
    If Not VeriBaglantisiHazir() Then
        Mesaj "Veri dosyası olmadan program çalışamaz. Program kapatılacak.", vbCritical
        CikisOnay = True
        DoCmd.Quit acQuitSaveNone
        Exit Function
    End If
    ' Kullanım (lisans) kontrolü: kapalıysa uyarı ekranı, yalnızca yedek alınabilir
    If Not LisansKontrol() Then
        CikisOnay = True
        DoCmd.Quit acQuitSaveNone
        Exit Function
    End If
    ShiftKapat
    KapanisFormunuAc
    ' Bağlantı yeni kurulduysa firma adı gibi hesaplanan alanları yenile
    On Error Resume Next
    Forms("frm_Giris").Recalc
    On Error GoTo Hata
    KullaniciVarsayilanOlustur
    OturumID = 0: OturumAd = "": OturumAdmin = False
    ProfilAdet = 0: Sayfa = 0: SeciliID = 0
    ReDim Profiller(0)
    Set rs = CurrentDb.OpenRecordset("SELECT KullaniciID FROM tbl_Kullanici WHERE Aktif=True ORDER BY Yonetici, AdSoyad", dbOpenSnapshot)
    Do Until rs.EOF
        ReDim Preserve Profiller(ProfilAdet)
        Profiller(ProfilAdet) = rs!KullaniciID
        ProfilAdet = ProfilAdet + 1
        rs.MoveNext
    Loop
    rs.Close
    BaslikYenile
    ProfilleriGoster
    Exit Function
Hata:
    Mesaj "Giriş ekranı hazırlanamadı: " & Err.Description, vbExclamation
End Function

Private Sub ProfilleriGoster()
    On Error Resume Next
    Dim frm As Form, i As Integer, idx As Long, rs As DAO.Recordset, gorunen As Integer
    Set frm = G()
    For i = 1 To PROFIL_SAYISI
        idx = Sayfa * PROFIL_SAYISI + i - 1
        If idx < ProfilAdet Then
            Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Kullanici WHERE KullaniciID=" & Profiller(idx), dbOpenSnapshot)
            frm.Controls("imgProfil" & i).PictureData = frm.Controls("imgRenk" & (Nz(rs!Renk, 0) Mod RENK_SAYISI)).PictureData
            frm.Controls("lblBas" & i).Caption = bas(Nz(rs!adSoyad, rs!KullaniciAdi))
            frm.Controls("lblAd" & i).Caption = Nz(rs!adSoyad, rs!KullaniciAdi) & IIf(rs!Yonetici And InStr(Nz(rs!adSoyad, ""), "Yönetici") = 0, vbCrLf & "Yönetici", "")
            frm.Controls("imgSecim" & i).Visible = (Profiller(idx) = SeciliID)
            rs.Close
            gorunen = gorunen + 1
            ProfilGorunur i, True
        Else
            ProfilGorunur i, False
        End If
    Next
    ' kartları ortala
    Dim genislik As Long, bosluk As Long, sol As Long
    genislik = 2000: bosluk = 400
    sol = (frm.Width - (gorunen * genislik + (gorunen - 1) * bosluk)) / 2
    For i = 1 To gorunen
        ProfilKonum i, sol + (i - 1) * (genislik + bosluk)
    Next
    frm!btnOnceki.Visible = (Sayfa > 0)
    frm!btnSonraki.Visible = ((Sayfa + 1) * PROFIL_SAYISI < ProfilAdet)
    SifrePaneli (SeciliID <> 0)
End Sub

Private Sub ProfilGorunur(ByVal i As Integer, ByVal goster As Boolean)
    Dim frm As Form
    Set frm = G()
    frm.Controls("imgProfil" & i).Visible = goster
    frm.Controls("lblBas" & i).Visible = goster
    frm.Controls("lblAd" & i).Visible = goster
    frm.Controls("btnProfil" & i).Visible = goster
    If Not goster Then frm.Controls("imgSecim" & i).Visible = False
End Sub

Private Sub ProfilKonum(ByVal i As Integer, ByVal sol As Long)
    Dim frm As Form
    Set frm = G()
    frm.Controls("imgSecim" & i).Left = sol
    frm.Controls("imgProfil" & i).Left = sol + 100
    frm.Controls("lblBas" & i).Left = sol + 100
    frm.Controls("btnProfil" & i).Left = sol + 100
    frm.Controls("lblAd" & i).Left = sol - 200
End Sub

Private Sub SifrePaneli(ByVal goster As Boolean)
    On Error Resume Next
    Dim frm As Form, rs As DAO.Recordset
    Set frm = G()
    If goster Then
        frm!lblSecili.Caption = Nz(DLookup("AdSoyad", "tbl_Kullanici", "KullaniciID=" & SeciliID), "")
        frm!lblSifre.Visible = True
        frm!txtSifre.Visible = True
        frm!btnGiris.Visible = True
        frm!lblSecili.Visible = True
        frm!lblIpucu.Caption = "Şifrenizi yazıp Enter'a basın"
    Else
        frm!lblSecili.Visible = False
        frm!lblSifre.Visible = False
        frm!txtSifre.Visible = False
        frm!btnGiris.Visible = False
        frm!lblIpucu.Caption = "Devam etmek için profilinizi seçin"
    End If
End Sub

Public Function ProfilSec(ByVal i As Integer)
    On Error Resume Next
    Dim idx As Long, frm As Form
    idx = Sayfa * PROFIL_SAYISI + i - 1
    If idx >= ProfilAdet Then Exit Function
    SeciliID = Profiller(idx)
    Set frm = G()
    frm!txtSifre = Null
    ProfilleriGoster
    frm!txtSifre.SetFocus
End Function

Public Function ProfilSayfa(ByVal yon As Integer)
    Sayfa = Sayfa + yon
    If Sayfa < 0 Then Sayfa = 0
    ProfilleriGoster
End Function

Public Function GirisYap()
    On Error GoTo Hata
    Dim frm As Form, rs As DAO.Recordset
    Set frm = G()
    If SeciliID = 0 Then
        Mesaj "Önce profilinizi seçin.", vbExclamation: Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Kullanici WHERE KullaniciID=" & SeciliID, dbOpenDynaset)
    If rs.EOF Then rs.Close: Exit Function
    If Not SifreDogruMu(Nz(frm!txtSifre, ""), Nz(rs!sifre, "")) Then
        rs.Close
        frm!txtSifre = Null
        frm!lblIpucu.Caption = "Şifre hatalı, tekrar deneyin"
        Mesaj "Şifre hatalı.", vbExclamation
        frm!txtSifre.SetFocus
        Exit Function
    End If
    OturumID = rs!KullaniciID
    OturumAd = Nz(rs!adSoyad, rs!KullaniciAdi)
    OturumAdmin = Nz(rs!Yonetici, False)
    rs.Edit
    rs!SonGiris = Now
    rs.Update
    rs.Close
    DoCmd.Close acForm, GIRIS_FORMU
    DoCmd.OpenForm ANA_MENU
    Exit Function
Hata:
    Mesaj "Giriş hatası: " & Err.Description, vbExclamation
End Function

Public Function OturumuKapat()
    If Not Soru("Oturum kapatılsın mı?") Then Exit Function
    ' Açık tüm formları kapat
    Dim i As Long
    For i = Forms.Count - 1 To 0 Step -1
        If Forms(i).Name <> "frm_Kapanis" Then DoCmd.Close acForm, Forms(i).Name
    Next
    OturumID = 0: OturumAd = "": OturumAdmin = False
    DoCmd.OpenForm GIRIS_FORMU
End Function

' ---------------------------------------------------------------
' ANA MENÜ
' ---------------------------------------------------------------
Public Function MenuYukle()
    On Error Resume Next
    Dim frm As Form
    If OturumID = 0 And Not TestModu Then
        DoCmd.Close acForm, ANA_MENU
        DoCmd.OpenForm GIRIS_FORMU
        Exit Function
    End If
    Set frm = Forms(ANA_MENU)
    frm!btnKullanicilar.Visible = OturumAdmin Or TestModu
    frm!btnAyarlar.Visible = OturumAdmin Or TestModu
    frm!lblOturum.Caption = OturumBilgisi()
    frm!lblSezon.Caption = "Aktif sezon: " & AktifSezon()
    BaslikYenile
End Function

' ---------------------------------------------------------------
' KULLANICILAR (yalnızca yönetici)
' ---------------------------------------------------------------
Private Function k() As Form
    Set k = Forms("frm_Kullanicilar")
End Function

Public Function KullaniciTemizle()
    Dim frm As Form
    Set frm = k()
    KutulariTemizle frm, Array("txtKullaniciID", "txtKullaniciAdi", "txtAdSoyad", "txtSifre", "txtSifre2")
    frm!chkYonetici = False
    frm!chkAktif = True
    frm!cmbRenk = 1
    frm!lstKullanicilar = Null
End Function

Public Function KullaniciYeni()
    KullaniciTemizle
    k().Controls("txtKullaniciAdi").SetFocus
End Function

Public Function KullaniciKaydet()
    KullaniciYazIslem False
End Function

Public Function KullaniciGuncelle()
    KullaniciYazIslem True
End Function

Private Function YoneticiSayisi(ByVal haric As Long) As Long
    YoneticiSayisi = DCount("*", "tbl_Kullanici", "Yonetici=True AND Aktif=True AND KullaniciID<>" & haric)
End Function

Private Sub KullaniciYazIslem(ByVal guncelle As Boolean)
    On Error GoTo Hata
    Dim frm As Form, id As Long, rs As DAO.Recordset, sifre As String
    Set frm = k()
    If guncelle Then
        If Bos(frm!txtKullaniciID) Then
            Mesaj "Güncellemek için önce listeden bir kullanıcı seçin (çift tıklayın).", vbExclamation: Exit Sub
        End If
        id = CLng(frm!txtKullaniciID)
    ElseIf Not Bos(frm!txtKullaniciID) Then
        Mesaj "Bu kullanıcı zaten kayıtlı. Değişiklik için 'Güncelle' düğmesini kullanın.", vbExclamation: Exit Sub
    End If
    If Bos(frm!txtKullaniciAdi) Or Bos(frm!txtAdSoyad) Then
        Mesaj "Kullanıcı adı ve ad soyad giriniz.", vbExclamation: Exit Sub
    End If
    If DCount("*", "tbl_Kullanici", "KullaniciAdi=" & SqlMetin(Trim$(frm!txtKullaniciAdi)) & " AND KullaniciID<>" & id) > 0 Then
        Mesaj "Bu kullanıcı adı zaten kullanılıyor.", vbExclamation: Exit Sub
    End If
    sifre = Nz(frm!txtSifre, "")
    If Not guncelle And Len(sifre) < 4 Then
        Mesaj "Şifre en az 4 karakter olmalıdır.", vbExclamation: Exit Sub
    End If
    If sifre <> "" Then
        If Len(sifre) < 4 Then Mesaj "Şifre en az 4 karakter olmalıdır.", vbExclamation: Exit Sub
        If sifre <> Nz(frm!txtSifre2, "") Then Mesaj "Şifreler aynı değil.", vbExclamation: Exit Sub
    End If
    If guncelle Then
        If (Not Nz(frm!chkYonetici, False) Or Not Nz(frm!chkAktif, True)) And YoneticiSayisi(id) = 0 Then
            Mesaj "En az bir aktif yönetici kalmalıdır.", vbExclamation: Exit Sub
        End If
    End If

    Set rs = KayitAc("tbl_Kullanici", "KullaniciID", id)
    rs!KullaniciAdi = Trim$(frm!txtKullaniciAdi)
    rs!adSoyad = Trim$(frm!txtAdSoyad)
    If sifre <> "" Then rs!sifre = SifreSakla(sifre)
    rs!Yonetici = Nz(frm!chkYonetici, False)
    rs!Aktif = Nz(frm!chkAktif, True)
    rs!Renk = Nz(frm!cmbRenk, 0)
    If id = 0 Then rs!OlusturmaTarihi = Now
    rs.Update
    rs.Bookmark = rs.LastModified
    id = rs!KullaniciID
    rs.Close

    If id = OturumID Then
        OturumAd = Trim$(frm!txtAdSoyad)
        OturumAdmin = Nz(frm!chkYonetici, False)
    End If
    frm!txtKullaniciID = id
    frm!txtSifre = Null: frm!txtSifre2 = Null
    frm!lstKullanicilar.Requery
    Mesaj IIf(guncelle, "Kullanıcı güncellendi.", "Kullanıcı eklendi.")
    Exit Sub
Hata:
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Sub

Public Function KullaniciListedenSec()
    On Error GoTo Hata
    Dim frm As Form, rs As DAO.Recordset, id As Long
    Set frm = k()
    id = ListeID(frm!lstKullanicilar)
    If id = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Kullanici WHERE KullaniciID=" & id, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function
    frm!txtKullaniciID = rs!KullaniciID
    frm!txtKullaniciAdi = rs!KullaniciAdi
    frm!txtAdSoyad = rs!adSoyad
    frm!txtSifre = Null: frm!txtSifre2 = Null
    frm!chkYonetici = Nz(rs!Yonetici, False)
    frm!chkAktif = Nz(rs!Aktif, True)
    frm!cmbRenk = Nz(rs!Renk, 0)
    rs.Close
    Exit Function
Hata:
    Mesaj "Kullanıcı açılamadı: " & Err.Description, vbExclamation
End Function

Public Function KullaniciSil()
    On Error GoTo Hata
    Dim frm As Form, id As Long
    Set frm = k()
    If Bos(frm!txtKullaniciID) Then
        Mesaj "Silmek için önce listeden bir kullanıcı seçin.", vbExclamation: Exit Function
    End If
    id = CLng(frm!txtKullaniciID)
    If id = OturumID Then
        Mesaj "Oturum açık olan kendi hesabınızı silemezsiniz.", vbExclamation: Exit Function
    End If
    If Nz(DLookup("Yonetici", "tbl_Kullanici", "KullaniciID=" & id), False) And YoneticiSayisi(id) = 0 Then
        Mesaj "Son yönetici silinemez.", vbExclamation: Exit Function
    End If
    If Not Soru(Nz(frm!txtAdSoyad, "") & " kullanıcısı silinsin mi?") Then Exit Function
    CurrentDb.Execute "DELETE FROM tbl_Kullanici WHERE KullaniciID=" & id, dbFailOnError
    KullaniciTemizle
    frm!lstKullanicilar.Requery
    Mesaj "Kullanıcı silindi."
    Exit Function
Hata:
    Mesaj "Silme sırasında hata: " & Err.Description, vbExclamation
End Function

' ---------------------------------------------------------------
' AYARLAR (yalnızca yönetici)
' ---------------------------------------------------------------
Private Function a() As Form
    Set a = Forms("frm_Ayarlar")
End Function

Public Function AyarlarYukle()
    On Error Resume Next
    Dim frm As Form
    Set frm = a()
    frm!txtFirmaAdi = Ayar("FirmaAdi")
    frm!txtFirmaTelefon = MetinVeyaNull(Ayar("FirmaTelefon"))
    frm!txtFirmaAdres = MetinVeyaNull(Ayar("FirmaAdres"))
    frm!txtHakyagiOrani = AyarSayi("HakyagiOrani", 12)
    frm!txtYagFiyati = AyarSayi("YagFiyati", 200)
    frm!cmbAktifSezon.RowSource = SezonListesi()
    frm!cmbAktifSezon = AktifSezon()
    frm!lblVeriYolu.Caption = "Veri dosyası: " & VeriYolu() & "   (sürüm " & VeriSurumu() & ")"
    frm!lblSonYedek.Caption = SonYedekMetni()
    frm!lblLisans.Caption = LisansBilgiMetni()
    KutulariTemizle frm, Array("txtEskiSifre", "txtYeniSifre", "txtYeniSifre2")
    HesapTemizle
End Function

Public Function AyarlarKaydet()
    On Error GoTo Hata
    Dim frm As Form, eskiSezon As String
    Set frm = a()
    If Bos(frm!txtFirmaAdi) Then
        Mesaj "Firma adını giriniz.", vbExclamation: Exit Function
    End If
    If Not SayiMi(frm!txtHakyagiOrani) Then
        Mesaj "Hakyağı oranını (%) sayı olarak giriniz.", vbExclamation: Exit Function
    End If
    If Sayi(frm!txtHakyagiOrani) <= 0 Or Sayi(frm!txtHakyagiOrani) >= 100 Then
        Mesaj "Hakyağı oranı 0 ile 100 arasında olmalıdır.", vbExclamation: Exit Function
    End If
    If Not SayiMi(frm!txtYagFiyati) Or Sayi(frm!txtYagFiyati) <= 0 Then
        Mesaj "1 kg yağ fiyatını (TL) giriniz.", vbExclamation: Exit Function
    End If
    If Len(Nz(frm!cmbAktifSezon, "")) <> 9 Then
        Mesaj "Aktif sezonu seçiniz.", vbExclamation: Exit Function
    End If
    eskiSezon = AktifSezon()
    AyarYaz "FirmaAdi", Trim$(frm!txtFirmaAdi)
    AyarYaz "FirmaTelefon", frm!txtFirmaTelefon
    AyarYaz "FirmaAdres", frm!txtFirmaAdres
    AyarYaz "HakyagiOrani", CStr(Sayi(frm!txtHakyagiOrani))
    AyarYaz "YagFiyati", CStr(Sayi(frm!txtYagFiyati))
    AyarYaz "AktifSezon", frm!cmbAktifSezon
    BaslikYenile
    BasliklariYenile
    Mesaj "Ayarlar kaydedildi." & vbCrLf & vbCrLf & _
        "Hakyağı oranı ve yağ fiyatı bundan sonra açılan partilere uygulanır; kayıtlı partiler değişmez." & _
        IIf(eskiSezon <> AktifSezon(), vbCrLf & "Aktif sezon " & AktifSezon() & " oldu. Açık ekranları kapatıp yeniden açın.", "")
    Exit Function
Hata:
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Function

' ---------------------------------------------------------------
' Kasa ve banka hesapları (Ayarlar ekranı)
' ---------------------------------------------------------------
Public Function HesapTemizle()
    On Error Resume Next
    Dim frm As Form
    Set frm = a()
    KutulariTemizle frm, Array("txtHesapID", "txtHesapAdi", "txtHesapBanka", "txtHesapIBAN")
    frm!cmbHesapTuru = "Banka"
    frm!lstHesaplar.Requery
End Function

Public Function HesapListedenSec()
    On Error Resume Next
    Dim frm As Form, rs As DAO.Recordset, id As Long
    Set frm = a()
    id = ListeID(frm!lstHesaplar)
    If id = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Kasa WHERE KasaID=" & id, dbOpenSnapshot)
    If Not rs.EOF Then
        frm!txtHesapID = rs!KasaID
        frm!txtHesapAdi = rs!KasaAdi
        frm!cmbHesapTuru = Nz(rs!KasaTuru, "Nakit")
        frm!txtHesapBanka = rs!BankaAdi
        frm!txtHesapIBAN = rs.Fields("IBAN")
    End If
    rs.Close
End Function

Public Function HesapKaydet()
    On Error GoTo Hata
    Dim frm As Form, rs As DAO.Recordset, id As Long
    Set frm = a()
    If Bos(frm!txtHesapAdi) Then
        Mesaj "Hesap adını giriniz (ör. Ziraat Bankası TL).", vbExclamation: Exit Function
    End If
    id = Nz(frm!txtHesapID, 0)
    Set rs = KayitAc("tbl_Kasa", "KasaID", id)
    rs!KasaAdi = Trim$(frm!txtHesapAdi)
    rs!KasaTuru = IIf(Nz(frm!cmbHesapTuru, "Banka") = "Nakit", "Nakit", "Banka")
    rs!BankaAdi = MetinVeyaNull(frm!txtHesapBanka)
    rs.Fields("IBAN") = MetinVeyaNull(UCase$(Replace(Nz(frm!txtHesapIBAN, ""), " ", "")))
    rs!Aktif = True
    If id = 0 Then rs!AcilisBakiyesi = 0
    rs.Update
    rs.Close
    HesapTemizle
    KombolariYenile
    Mesaj "Hesap kaydedildi."
    Exit Function
Hata:
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Function

Public Function HesapSil()
    On Error GoTo Hata
    Dim frm As Form, id As Long
    Set frm = a()
    id = Nz(frm!txtHesapID, 0)
    If id = 0 Then
        Mesaj "Önce listeden bir hesap seçin (çift tıklayın).", vbExclamation: Exit Function
    End If
    If id = VarsayilanKasaID() Then
        Mesaj "Ana nakit kasa silinemez.", vbExclamation: Exit Function
    End If
    If DCount("*", "tbl_KasaHareket", "KasaID=" & id) > 0 Then
        If Not Soru("Bu hesapta hareket var; silinemez. Hesap pasif yapılsın mı? (Listelerde görünmez.)") Then Exit Function
        CurrentDb.Execute "UPDATE tbl_Kasa SET Aktif=False WHERE KasaID=" & id, dbFailOnError
    Else
        If Not Soru(Nz(frm!txtHesapAdi, "") & " silinsin mi?") Then Exit Function
        CurrentDb.Execute "DELETE FROM tbl_Kasa WHERE KasaID=" & id, dbFailOnError
    End If
    HesapTemizle
    KombolariYenile
    Exit Function
Hata:
    Mesaj "Silme sırasında hata: " & Err.Description, vbExclamation
End Function

' ---------------------------------------------------------------
' Yönetici araçları (Ayarlar ekranı)
' ---------------------------------------------------------------
' Shift ile açılışı kapatır (program her açılışta çağırır)
Public Sub ShiftKapat()
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb
    db.Properties("AllowBypassKey") = False
    If Err.Number <> 0 Then
        Err.Clear
        db.Properties.Append db.CreateProperty("AllowBypassKey", dbBoolean, False)
    End If
End Sub

' Bir sonraki açılışta Shift'e basılı tutarak tasarım görünümüne girilebilir.
' Program normal açılınca Shift yeniden kapanır.
Public Function GelistiriciModu()
    On Error Resume Next
    If Not OturumAdmin And Not TestModu Then Exit Function
    If Not Soru("Program bir sonraki açılışta Shift tuşuna basılı tutularak tasarım görünümünde açılabilecek." & vbCrLf & _
                "(Normal açılışta bu izin yeniden kapanır.) Devam edilsin mi?") Then Exit Function
    CurrentDb.Properties("AllowBypassKey") = True
    Mesaj "Tamam. Programı kapatın; açarken Shift tuşuna basılı tutun."
End Function

Public Function KaynakAktarDugme()
    Dim yol As String
    yol = KaynakDisaAktar()
    If yol <> "" Then Mesaj "Kaynaklar dışa aktarıldı:" & vbCrLf & yol
End Function

Public Function KayitlariYenidenOlustur()
    On Error GoTo Hata
    If Not Soru("Tüm partilerin otomatik cari / kasa / stok kayıtları ve bidon listeleri yeniden oluşturulsun mu?" & vbCrLf & _
                "(Tutarlar değişmez; elle girilen kayıtlar korunur.)") Then Exit Function
    DoCmd.Hourglass True
    TumPartiKayitlariniYenidenYaz
    DoCmd.Hourglass False
    Mesaj "Kayıtlar yeniden oluşturuldu."
    Exit Function
Hata:
    DoCmd.Hourglass False
    Mesaj "İşlem sırasında hata: " & Err.Description, vbExclamation
End Function

' Açık formlardaki firma adı yazılarını yeniler
Public Sub BasliklariYenile()
    On Error Resume Next
    Dim frm As Form
    For Each frm In Forms
        frm.Recalc
    Next
End Sub

Public Function SifreDegistir()
    On Error GoTo Hata
    Dim frm As Form, saklanan As String
    Set frm = a()
    If OturumID = 0 Then
        Mesaj "Şifre değiştirmek için giriş yapmış olmalısınız.", vbExclamation: Exit Function
    End If
    saklanan = Nz(DLookup("Sifre", "tbl_Kullanici", "KullaniciID=" & OturumID), "")
    If Not SifreDogruMu(Nz(frm!txtEskiSifre, ""), saklanan) Then
        Mesaj "Mevcut şifre hatalı.", vbExclamation: Exit Function
    End If
    If Len(Nz(frm!txtYeniSifre, "")) < 4 Then
        Mesaj "Yeni şifre en az 4 karakter olmalıdır.", vbExclamation: Exit Function
    End If
    If Nz(frm!txtYeniSifre, "") <> Nz(frm!txtYeniSifre2, "") Then
        Mesaj "Yeni şifreler aynı değil.", vbExclamation: Exit Function
    End If
    CurrentDb.Execute "UPDATE tbl_Kullanici SET Sifre=" & SqlMetin(SifreSakla(frm!txtYeniSifre)) & _
        " WHERE KullaniciID=" & OturumID, dbFailOnError
    KutulariTemizle frm, Array("txtEskiSifre", "txtYeniSifre", "txtYeniSifre2")
    Mesaj "Şifreniz değiştirildi."
    Exit Function
Hata:
    Mesaj "Şifre değiştirilemedi: " & Err.Description, vbExclamation
End Function