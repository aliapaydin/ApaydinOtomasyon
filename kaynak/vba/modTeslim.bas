Option Compare Database
Option Explicit

' =============================================================
' YAĞ TESLİM   (form: frm_CicerTeslim)
' -------------------------------------------------------------
' Partinin yağı müşteriye teslim edilirken:
'   teslim tarihi, teslim alan, teslim eden (personel),
'   teslim edilen yağ miktarı ve açıklama kaydedilir.
' Ekranda parti özeti görünür: gelen zeytin, çıkan yağ, verim
' (kaç kg zeytinden 1 kg yağ çıktığı), randıman %, hakyağı,
' müşteriye kalan net yağ veya sıkım ücreti ve tahsil durumu.
' Ücretli partide ücret teslimde alındıysa işaretlenir; tahsilat
' teslim tarihiyle kasaya ve cariye işlenir.
' Teslim fişi A5 olarak önizlenir, PDF kaydedilir veya yazdırılır.
' =============================================================

Private Const FORM_ADI As String = "frm_CicerTeslim"
Private Const BILGI As String = "txtMusteri,txtTelefon,txtPartiTarihi,txtCinsi,txtOdemeTipi,txtZeytin,txtYag,txtVerim,txtRandiman,txtHakyagi,txtNetYag,txtUcret,txtTahsil,txtBidonlar"
Private Const TESLIM As String = "txtTeslimTarihi,txtTeslimAlan,cmbTeslimEden,txtTeslimYag,txtTeslimAciklama"

Private Function F() As Form
    Set F = Forms(FORM_ADI)
End Function

Public Function TeslimFormYukle()
    On Error Resume Next
    F().Controls("cmbDurumFiltre").Value = "Bekleyenler"
    TeslimListeFiltrele
    TeslimTemizle
End Function

Public Function TeslimTemizle()
    Dim frm As Form
    Set frm = F()
    KutulariTemizle frm, Split(BILGI & ",cmbParti," & TESLIM, ",")
    frm!chkTahsil = False
    frm!chkTahsil.Enabled = False
    frm!lstTeslim = Null
    frm!lblDurum.Caption = ""
End Function

Public Function TeslimListeFiltrele()
    On Error Resume Next
    Dim frm As Form, filtre As String
    Set frm = F()
    Select Case Nz(frm!cmbDurumFiltre, "")
        Case "Bekleyenler": filtre = " WHERE TeslimTarihiHam Is Null"
        Case "Teslim Edilenler": filtre = " WHERE TeslimTarihiHam Is Not Null"
    End Select
    frm!lstTeslim.RowSource = "SELECT * FROM qry_TeslimListe" & filtre & ";"
End Function

Public Function TeslimListedenSec()
    Dim id As Long
    id = ListeID(F().Controls("lstTeslim"))
    If id = 0 Then Exit Function
    F().Controls("cmbParti").Value = id
    TeslimPartiSecildi
End Function

' Parti seçilince özet bilgileri doldurur
Public Function TeslimPartiSecildi()
    On Error GoTo Hata
    Dim frm As Form, rs As DAO.Recordset, zeytin As Double, yag As Double, net As Double
    Set frm = F()
    KutulariTemizle frm, Split(BILGI & "," & TESLIM, ",")
    frm!chkTahsil = False
    If Bos(frm!cmbParti) Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT p.*, m.AdSoyad, m.Telefon FROM tbl_Parti AS p LEFT JOIN tbl_Musteri AS m ON p.MusteriID=m.MusteriID WHERE p.PartiNo=" & CLng(frm!cmbParti), dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function

    zeytin = Nz(rs!ZeytinMiktari_Kg, 0)
    yag = Nz(rs!CikanYag_Kg, 0)
    frm!txtMusteri = rs!adSoyad
    frm!txtTelefon = rs!Telefon
    frm!txtPartiTarihi = Format$(rs!tarih, "dd.mm.yyyy")
    frm!txtCinsi = Trim$(Nz(rs!Cinsi, "") & " / " & Nz(rs!Ozellik, ""))
    frm!txtOdemeTipi = rs!odemeTipi
    frm!txtZeytin = KgMetin(zeytin)
    frm!txtYag = KgMetin(yag)
    If zeytin > 0 And yag > 0 Then
        frm!txtVerim = Format$(zeytin / yag, "0.00") & " kg zeytin = 1 kg yağ"
        frm!txtRandiman = "%" & Format$(yag / zeytin * 100, "0.00")
    End If
    If HakyagiMi(rs!odemeTipi) Then
        net = Nz(rs!MusteriKalanYag_Kg, yag - Nz(rs!HakYagi_Kg, 0))
        frm!txtHakyagi = "%" & Nz(rs!HakYagiOrani_Yuzde, 0) & "  ›  " & KgMetin(Nz(rs!HakYagi_Kg, 0))
        frm!txtUcret = "-"
        frm!txtTahsil = "-"
        frm!chkTahsil.Enabled = False
    Else
        net = yag
        frm!txtHakyagi = "-"
        frm!txtUcret = TLYaz(rs!TLTutari.Value) & "  (" & Format$(zeytin, "#,##0") & " kg x " & Format$(Nz(rs!SimSikimUcreti, 0), "0.00") & ")"
        If Nz(rs!PesinTahsilEdildi, False) Then
            frm!txtTahsil = "Tahsil edildi (" & Format$(Nz(rs!TahsilatTarihi, rs!tarih), "dd.mm.yyyy") & ")"
            frm!chkTahsil = True
            frm!chkTahsil.Enabled = False
        Else
            frm!txtTahsil = "Tahsil edilmedi"
            frm!chkTahsil.Enabled = True
        End If
    End If
    frm!txtNetYag = KgMetin(net)
    frm!txtBidonlar = rs!BidonDetay

    ' Teslim bilgileri: kayıtlıysa onları, değilse varsayılanları getir
    If Not IsNull(rs!TeslimTarihi) Then
        frm!txtTeslimTarihi = rs!TeslimTarihi
        frm!txtTeslimAlan = rs!TeslimAlan
        frm!cmbTeslimEden = rs!TeslimPersonelID
        frm!txtTeslimYag = Nz(rs!TeslimYag_Kg, net)
        frm!txtTeslimAciklama = rs!TeslimAciklama
        frm!lblDurum.Caption = "TESLİM EDİLDİ"
    Else
        frm!txtTeslimTarihi = Date
        frm!txtTeslimAlan = rs!adSoyad
        frm!txtTeslimYag = net
        frm!lblDurum.Caption = "TESLİM BEKLİYOR"
    End If
    rs.Close
    Exit Function
Hata:
    Mesaj "Parti bilgileri okunamadı: " & Err.Description, vbExclamation
End Function

Public Function TeslimKaydet()
    On Error GoTo Hata
    Dim frm As Form, id As Long, rs As DAO.Recordset, islemde As Boolean, tahsilEt As Boolean
    Set frm = F()
    If Bos(frm!cmbParti) Then
        Mesaj "Önce parti seçin.", vbExclamation: Exit Function
    End If
    If Not TarihMi(frm!txtTeslimTarihi) Then
        Mesaj "Teslim tarihini giriniz.", vbExclamation: Exit Function
    End If
    If Bos(frm!txtTeslimAlan) Then
        Mesaj "Teslim alan kişiyi giriniz.", vbExclamation: Exit Function
    End If
    If Bos(frm!cmbTeslimEden) Then
        Mesaj "Teslim eden personeli seçiniz.", vbExclamation: Exit Function
    End If
    If Sayi(frm!txtTeslimYag) <= 0 Then
        Mesaj "Teslim edilen yağ miktarını giriniz.", vbExclamation: Exit Function
    End If
    id = CLng(frm!cmbParti)
    tahsilEt = Nz(frm!chkTahsil, False) And frm!chkTahsil.Enabled

    DBEngine.BeginTrans: islemde = True
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Parti WHERE PartiNo=" & id, dbOpenDynaset)
    rs.Edit
    rs!TeslimTarihi = CDate(frm!txtTeslimTarihi)
    rs!TeslimAlan = Trim$(frm!txtTeslimAlan)
    rs!TeslimPersonelID = CLng(frm!cmbTeslimEden)
    rs!TeslimYag_Kg = Sayi(frm!txtTeslimYag)
    rs!TeslimAciklama = MetinVeyaNull(frm!txtTeslimAciklama)
    If tahsilEt Then
        rs!PesinTahsilEdildi = True
        rs!TahsilatTarihi = CDate(frm!txtTeslimTarihi)
    End If
    rs.Update
    rs.Close
    If tahsilEt Then PartiMuhasebeYaz id
    BidonlariEsle id
    DBEngine.CommitTrans: islemde = False

    frm!lstTeslim.Requery
    TeslimPartiSecildi
    Mesaj "Teslim kaydedildi." & IIf(tahsilEt, vbCrLf & "Sıkım ücreti kasaya ve cariye tahsilat olarak işlendi.", "") & _
          vbCrLf & vbCrLf & "Teslim fişini 'Fiş Önizle', 'PDF Kaydet' veya 'Yazdır' ile alabilirsiniz."
    Exit Function
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "Teslim kaydedilemedi: " & Err.Description, vbExclamation
End Function

Public Function TeslimGeriAl()
    On Error GoTo Hata
    Dim frm As Form, id As Long, islemde As Boolean
    Set frm = F()
    If Bos(frm!cmbParti) Then
        Mesaj "Önce parti seçin.", vbExclamation: Exit Function
    End If
    id = CLng(frm!cmbParti)
    If IsNull(DLookup("TeslimTarihi", "tbl_Parti", "PartiNo=" & id)) Then
        Mesaj "Bu partinin teslim kaydı yok.", vbExclamation: Exit Function
    End If
    If Not Soru("Bu partinin teslim kaydı geri alınsın mı?" & vbCrLf & "(Alınmış tahsilat kayıtları korunur.)") Then Exit Function
    DBEngine.BeginTrans: islemde = True
    CurrentDb.Execute "UPDATE tbl_Parti SET TeslimTarihi=Null, TeslimAlan=Null, TeslimYag_Kg=Null, TeslimAciklama=Null, TeslimPersonelID=Null WHERE PartiNo=" & id, dbFailOnError
    BidonlariEsle id
    DBEngine.CommitTrans: islemde = False
    frm!lstTeslim.Requery
    TeslimPartiSecildi
    Mesaj "Teslim kaydı geri alındı."
    Exit Function
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "İşlem yapılamadı: " & Err.Description, vbExclamation
End Function

' Teslim fişi: mod = "onizle" / "pdf" / "yazdir"
Public Function TeslimFisi(ByVal cikti As String)
    Dim frm As Form
    Set frm = F()
    If Bos(frm!cmbParti) Then
        Mesaj "Önce parti seçin.", vbExclamation: Exit Function
    End If
    If IsNull(DLookup("TeslimTarihi", "tbl_Parti", "PartiNo=" & CLng(frm!cmbParti))) Then
        If Not Soru("Bu parti henüz teslim edilmedi. Fiş yine de hazırlansın mı?") Then Exit Function
    End If
    TeslimFisiAc CLng(frm!cmbParti), cikti
End Function