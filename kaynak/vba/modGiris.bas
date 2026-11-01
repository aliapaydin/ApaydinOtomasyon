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
    Do While Forms.Count > 0
        DoCmd.Close acForm, Forms(0).Name
    Loop
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
    frm!txtVarsayilanUcret = MetinVeyaNull(Ayar("VarsayilanSikimUcreti"))
    frm!txtVarsayilanOran = MetinVeyaNull(Ayar("VarsayilanHakyagiOrani"))
    KutulariTemizle frm, Array("txtEskiSifre", "txtYeniSifre", "txtYeniSifre2")
End Function

Public Function AyarlarKaydet()
    On Error GoTo Hata
    Dim frm As Form
    Set frm = a()
    If Bos(frm!txtFirmaAdi) Then
        Mesaj "Firma adını giriniz.", vbExclamation: Exit Function
    End If
    If Not Bos(frm!txtVarsayilanUcret) And Not SayiMi(frm!txtVarsayilanUcret) Then
        Mesaj "Varsayılan sıkım ücreti sayı olmalıdır.", vbExclamation: Exit Function
    End If
    If Not Bos(frm!txtVarsayilanOran) And Not SayiMi(frm!txtVarsayilanOran) Then
        Mesaj "Varsayılan hakyağı oranı sayı olmalıdır.", vbExclamation: Exit Function
    End If
    AyarYaz "FirmaAdi", Trim$(frm!txtFirmaAdi)
    AyarYaz "FirmaTelefon", frm!txtFirmaTelefon
    AyarYaz "FirmaAdres", frm!txtFirmaAdres
    AyarYaz "VarsayilanSikimUcreti", frm!txtVarsayilanUcret
    AyarYaz "VarsayilanHakyagiOrani", frm!txtVarsayilanOran
    BaslikYenile
    BasliklariYenile
    Mesaj "Ayarlar kaydedildi."
    Exit Function
Hata:
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
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