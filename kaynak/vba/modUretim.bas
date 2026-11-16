Option Compare Database
Option Explicit

' =============================================================
' ÜRETİM (SIKIM)   (form: frm_Uretim)
' -------------------------------------------------------------
' Mal kabulü yapılmış parti seçilir; sıkım tarihi ve çıkan yağ girilir.
'   Kızıl su / Posa = Net zeytin - Çıkan yağ
'   Randıman %      = Çıkan yağ / Net zeytin x 100
'   Hakyağı seçilirse : Hakyağı kg = Çıkan yağ x oran
'                       Müşteriye kalan = Çıkan yağ - Hakyağı
'                       Hakyağı stoğa girer.
'   Sıkım ücreti seçilirse:
'                       Ücret = Çıkan yağ x oran x 1 kg yağ fiyatı
'                       Müşteriye kalan = Çıkan yağ
'                       Ücret müşterinin carisine BORÇ yazılır.
' Oran ve yağ fiyatı mal kabulde partiye kopyalanmıştır; bu ekranda
' o parti için değiştirilebilir.
' =============================================================

Private Const FORM_ADI As String = "frm_Uretim"
Private Const BILGI As String = "txtMusteri,txtKabulTarihi,txtCinsi,txtNetZeytin,txtBidonlar,txtDurum"
Private Const GIRIS As String = "txtSikimTarihi,txtCikanYag,cmbOdemeTipi,txtOran,txtYagFiyati"
Private Const HESAP As String = "txtPosa,txtRandiman,txtVerim,txtHakyagi,txtUcret,txtKalanYag"

Private Function F() As Form
    Set F = Forms(FORM_ADI)
End Function

Public Function UretimYukle()
    On Error Resume Next
    F().Controls("cmbListeFiltre").Value = "Üretim Bekleyenler"
    UretimTemizle
End Function

Public Function UretimTemizle()
    Dim frm As Form
    Set frm = F()
    KutulariTemizle frm, Split("cmbParti," & BILGI & "," & GIRIS & "," & HESAP, ",")
    frm!lstUretim = Null
    frm!lblSezon.Caption = AktifSezonMetni()
    UretimListeFiltrele
    UretimPartiListesi
End Function

' Parti açılır kutusu: aktif sezonun teslim edilmemiş partileri
Private Sub UretimPartiListesi()
    F().Controls("cmbParti").RowSource = "SELECT PartiNo, [Parti No], Müşteri, Tarih, Durum FROM qry_MalKabulListe " & _
        "WHERE Sezon=" & SqlMetin(AktifSezon()) & " AND Durum<>'" & DURUM_TESLIM & "' ORDER BY PartiSira;"
End Sub

Public Function UretimListeFiltrele()
    On Error Resume Next
    Dim frm As Form, w As String
    Set frm = F()
    w = " WHERE Sezon=" & SqlMetin(AktifSezon())
    Select Case Nz(frm!cmbListeFiltre, "")
        Case "Üretim Bekleyenler": w = w & " AND Durum='" & DURUM_KABUL & "'"
        Case "Üretilenler": w = w & " AND Durum<>'" & DURUM_KABUL & "'"
    End Select
    frm!lstUretim.RowSource = "SELECT * FROM qry_UretimListe" & w & " ORDER BY TarihHam DESC, PartiSira DESC;"
End Function

Public Function UretimListedenSec()
    Dim id As Long
    id = ListeID(F().Controls("lstUretim"))
    If id = 0 Then Exit Function
    F().Controls("cmbParti").Value = id
    UretimPartiSecildi
End Function

Public Function UretimPartiSecildi()
    On Error GoTo Hata
    Dim frm As Form, rs As DAO.Recordset
    Set frm = F()
    KutulariTemizle frm, Split(BILGI & "," & GIRIS & "," & HESAP, ",")
    If Bos(frm!cmbParti) Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT p.*, m.AdSoyad FROM tbl_Parti AS p LEFT JOIN tbl_Musteri AS m ON p.MusteriID=m.MusteriID WHERE p.PartiNo=" & CLng(frm!cmbParti), dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function
    frm!txtMusteri = rs!adSoyad
    frm!txtKabulTarihi = Format$(rs!tarih, "dd.mm.yyyy")
    frm!txtCinsi = Nz(rs!Cinsi, "") & IIf(IsNull(rs!Ozellik), "", " / " & rs!Ozellik)
    frm!txtNetZeytin = rs!ZeytinMiktari_Kg
    frm!txtBidonlar = rs!BidonDetay
    frm!txtDurum = PartiDurumu(rs!SikimTarihi, rs!CikanYag_Kg, rs!TeslimTarihi)
    frm!txtSikimTarihi = Nz(rs!SikimTarihi, Date)
    frm!txtCikanYag = rs!CikanYag_Kg
    frm!cmbOdemeTipi = rs!odemeTipi
    frm!txtOran = Nz(rs!HakYagiOrani_Yuzde, VarsayilanOran())
    frm!txtYagFiyati = Nz(rs!YagFiyati, VarsayilanYagFiyati())
    rs.Close
    UretimHesapla
    If Bos(frm!txtCikanYag) Then frm!txtCikanYag.SetFocus
    Exit Function
Hata:
    Mesaj "Parti bilgileri okunamadı: " & Err.Description, vbExclamation
End Function

' Ekrandaki değerlerle hesap kutularını doldurur
Public Function UretimHesapla()
    On Error Resume Next
    Dim frm As Form, zeytin As Double, yag As Double, oran As Double, fiyat As Double, hak As Double
    Set frm = F()
    zeytin = Sayi(frm!txtNetZeytin)
    yag = Sayi(frm!txtCikanYag)
    oran = Sayi(frm!txtOran)
    fiyat = Sayi(frm!txtYagFiyati)
    KutulariTemizle frm, Split(HESAP, ",")
    frm!txtYagFiyati.Enabled = UcretliMi(frm!cmbOdemeTipi)
    If zeytin <= 0 Or yag <= 0 Then Exit Function
    frm!txtPosa = KgMetin(zeytin - yag)
    frm!txtRandiman = "%" & Format$(yag / zeytin * 100, "0.00")
    frm!txtVerim = Format$(zeytin / yag, "0.00") & " kg zeytin = 1 kg yağ"
    If HakyagiMi(frm!cmbOdemeTipi) Then
        hak = HakyagiKg(yag, oran)
        frm!txtHakyagi = KgMetin(hak)
        frm!txtUcret = "-"
        frm!txtKalanYag = KgMetin(yag - hak)
    ElseIf UcretliMi(frm!cmbOdemeTipi) Then
        frm!txtHakyagi = "-"
        frm!txtUcret = TLYaz(SikimUcreti(yag, oran, fiyat))
        frm!txtKalanYag = KgMetin(yag)
    End If
End Function

Public Function UretimKaydet()
    On Error GoTo Hata
    Dim frm As Form, id As Long, rs As DAO.Recordset, islemde As Boolean, zeytin As Double, yag As Double
    Set frm = F()
    If Bos(frm!cmbParti) Then
        Mesaj "Önce parti seçin.", vbExclamation: Exit Function
    End If
    id = CLng(frm!cmbParti)
    zeytin = Sayi(frm!txtNetZeytin)
    yag = Sayi(frm!txtCikanYag)
    If Not TarihMi(frm!txtSikimTarihi) Then
        Mesaj "Sıkım tarihini giriniz.", vbExclamation: Exit Function
    End If
    If CDate(frm!txtSikimTarihi) < CDate(DLookup("Tarih", "tbl_Parti", "PartiNo=" & id)) Then
        Mesaj "Sıkım tarihi mal kabul tarihinden önce olamaz.", vbExclamation: Exit Function
    End If
    If yag <= 0 Then
        Mesaj "Çıkan yağ miktarını (kg) giriniz.", vbExclamation: Exit Function
    End If
    If yag >= zeytin Then
        Mesaj "Çıkan yağ, net zeytin miktarından küçük olmalıdır.", vbExclamation: Exit Function
    End If
    If Not UcretliMi(frm!cmbOdemeTipi) And Not HakyagiMi(frm!cmbOdemeTipi) Then
        Mesaj "Ödeme tipini seçiniz.", vbExclamation: Exit Function
    End If
    If Sayi(frm!txtOran) <= 0 Or Sayi(frm!txtOran) >= 100 Then
        Mesaj "Hakyağı oranını (%) giriniz.", vbExclamation: Exit Function
    End If
    If UcretliMi(frm!cmbOdemeTipi) And Sayi(frm!txtYagFiyati) <= 0 Then
        Mesaj "1 kg yağ fiyatını giriniz.", vbExclamation: Exit Function
    End If
    If yag / zeytin > 0.4 Then
        If Not Soru("Randıman %" & Format$(yag / zeytin * 100, "0.0") & " görünüyor. Miktarlar doğru mu?") Then Exit Function
    End If
    If Not IsNull(DLookup("TeslimTarihi", "tbl_Parti", "PartiNo=" & id)) Then
        If Not Soru("Bu partinin yağı teslim edilmiş. Üretim bilgileri yine de değiştirilsin mi?") Then Exit Function
    End If

    DBEngine.BeginTrans: islemde = True
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Parti WHERE PartiNo=" & id, dbOpenDynaset)
    rs.Edit
    rs!SikimTarihi = CDate(frm!txtSikimTarihi)
    rs!CikanYag_Kg = yag
    rs!odemeTipi = IIf(UcretliMi(frm!cmbOdemeTipi), "Sıkım Ücreti", "Hakyağı")
    rs!HakYagiOrani_Yuzde = Sayi(frm!txtOran)
    rs!YagFiyati = Sayi(frm!txtYagFiyati)
    rs.Update
    rs.Close
    UretimYenidenHesapla id
    DBEngine.CommitTrans: islemde = False

    frm!lstUretim.Requery
    UretimPartiSecildi
    Mesaj "Üretim kaydedildi." & vbCrLf & _
        IIf(UcretliMi(frm!cmbOdemeTipi), "Sıkım ücreti (" & frm!txtUcret & ") müşterinin carisine borç yazıldı.", _
            "Hakyağı (" & frm!txtHakyagi & ") stoğa girdi.")
    Exit Function
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Function

' Kayıtlı oran / fiyat / yağ ile partinin hesaplarını ve otomatik kayıtlarını yeniler
Public Sub UretimYenidenHesapla(ByVal partiNo As Long)
    Dim rs As DAO.Recordset, zeytin As Double, yag As Double, oran As Double, hak As Double
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Parti WHERE PartiNo=" & partiNo, dbOpenDynaset)
    If rs.EOF Then rs.Close: Exit Sub
    zeytin = Nz(rs!ZeytinMiktari_Kg, 0)
    yag = Nz(rs!CikanYag_Kg, 0)
    oran = Nz(rs!HakYagiOrani_Yuzde, 0)
    rs.Edit
    If IsNull(rs!SikimTarihi) Or yag <= 0 Then
        rs!AtikMiktari_Kg = Null
        rs!HakYagi_Kg = Null
        rs!TLTutari = Null
        rs!MusteriKalanYag_Kg = Null
    Else
        rs!AtikMiktari_Kg = zeytin - yag
        If HakyagiMi(rs!odemeTipi) Then
            hak = HakyagiKg(yag, oran)
            rs!HakYagi_Kg = hak
            rs!TLTutari = 0
            rs!MusteriKalanYag_Kg = yag - hak
        Else
            rs!HakYagi_Kg = 0
            rs!TLTutari = SikimUcreti(yag, oran, Nz(rs!YagFiyati, 0))
            rs!MusteriKalanYag_Kg = yag
        End If
    End If
    rs.Update
    rs.Close
    PartiMuhasebeYaz partiNo
End Sub

Public Function UretimGeriAl()
    On Error GoTo Hata
    Dim frm As Form, id As Long, islemde As Boolean
    Set frm = F()
    If Bos(frm!cmbParti) Then
        Mesaj "Önce parti seçin.", vbExclamation: Exit Function
    End If
    id = CLng(frm!cmbParti)
    If IsNull(DLookup("SikimTarihi", "tbl_Parti", "PartiNo=" & id)) Then
        Mesaj "Bu partinin üretim kaydı yok.", vbExclamation: Exit Function
    End If
    If Not IsNull(DLookup("TeslimTarihi", "tbl_Parti", "PartiNo=" & id)) Then
        Mesaj "Yağı teslim edilmiş partinin üretimi geri alınamaz. Önce teslimi geri alın.", vbExclamation: Exit Function
    End If
    If Not Soru("Bu partinin üretim kaydı geri alınsın mı?" & vbCrLf & "(Sıkım ücreti borcu / hakyağı stok girişi silinir.)") Then Exit Function
    DBEngine.BeginTrans: islemde = True
    CurrentDb.Execute "UPDATE tbl_Parti SET SikimTarihi=Null, CikanYag_Kg=Null, AtikMiktari_Kg=Null, HakYagi_Kg=Null, " & _
        "TLTutari=Null, MusteriKalanYag_Kg=Null WHERE PartiNo=" & id, dbFailOnError
    OtomatikKayitlariSil KAYNAK_PARTI, id
    DBEngine.CommitTrans: islemde = False
    frm!lstUretim.Requery
    UretimPartiSecildi
    Mesaj "Üretim kaydı geri alındı."
    Exit Function
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "İşlem yapılamadı: " & Err.Description, vbExclamation
End Function
