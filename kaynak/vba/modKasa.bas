Option Compare Database
Option Explicit

' =============================================================
' KASA HAREKETLERİ   (form: frm_CicerKasa)
' -------------------------------------------------------------
' Tahsilat / Ödeme girişi. Müşteri seçilirse cariye de işlenir:
'   Tahsilat -> müşterinin carisine ALACAK
'   Ödeme    -> müşterinin carisine BORÇ
' Liste: Tahsilat, Ödeme ve yürüyen Bakiye. İki tarih arası süzülebilir.
' Parti / Alım-Satım'dan gelen otomatik kasa kayıtları buradan
' değiştirilemez; kaynağından düzeltilir.
' =============================================================

Private Const FORM_ADI As String = "frm_CicerKasa"
Private Const KUTULAR As String = "txtKasaHareketID,cmbMusteri,cmbHareketTuru,cmbOdemeYontemi,txtTutar,txtAciklama"

Private Function F() As Form
    Set F = Forms(FORM_ADI)
End Function

Public Function KasaFormYukle()
    On Error Resume Next
    Dim frm As Form
    Set frm = F()
    frm!txtKasaAcilis = Nz(DLookup("AcilisBakiyesi", "tbl_Kasa", "KasaID=" & VarsayilanKasaID()), 0)
    KasaTemizle
    KasaBakiyeGoster
End Function

Private Sub KasaBakiyeGoster()
    On Error Resume Next
    F().Controls("txtKasaBakiye").Value = KasaBakiye()
End Sub

Public Function KasaTemizle()
    Dim frm As Form
    Set frm = F()
    KutulariTemizle frm, Split(KUTULAR, ",")
    frm!txtTarih = Date
    frm!cmbOdemeYontemi = "Nakit"
    frm!lstHareketler = Null
End Function

Public Function KasaYeni()
    KasaTemizle
    F().Controls("cmbHareketTuru").SetFocus
End Function

Public Function KasaHareketKaydet()
    KasaYazIslem False
End Function

Public Function KasaHareketGuncelle()
    KasaYazIslem True
End Function

Private Function OtomatikKasaMi(ByVal id As Long) As Boolean
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT OtomatikMi, KaynakTipi FROM tbl_KasaHareket WHERE KasaHareketID=" & id, dbOpenSnapshot)
    If Not rs.EOF Then
        If Nz(rs!OtomatikMi, False) Then
            Mesaj "Bu kasa hareketi '" & KaynakAdi(rs!kaynakTipi) & "' ekranından otomatik oluşturuldu." & vbCrLf & _
                  "Değiştirmek veya silmek için o ekrandan düzeltin.", vbExclamation
            OtomatikKasaMi = True
        End If
    End If
    rs.Close
End Function

Private Sub KasaYazIslem(ByVal guncelle As Boolean)
    On Error GoTo Hata
    Dim frm As Form, id As Long, tahsilatMi As Boolean, tutar As Currency, islemde As Boolean
    Dim musteri As Variant, acik As String
    Set frm = F()

    If guncelle Then
        If Bos(frm!txtKasaHareketID) Then
            Mesaj "Güncellemek için önce listeden bir kasa hareketi seçin (çift tıklayın).", vbExclamation
            Exit Sub
        End If
        id = CLng(frm!txtKasaHareketID)
        If OtomatikKasaMi(id) Then Exit Sub
    ElseIf Not Bos(frm!txtKasaHareketID) Then
        Mesaj "Bu hareket zaten kayıtlı. Değişiklik için 'Güncelle' düğmesini kullanın." & vbCrLf & _
              "Yeni hareket için önce 'Yeni Kayıt' düğmesine basın.", vbExclamation
        Exit Sub
    End If
    If Not TarihMi(frm!txtTarih) Then
        Mesaj "Tarih giriniz.", vbExclamation: Exit Sub
    End If
    If Bos(frm!cmbHareketTuru) Then
        Mesaj "Hareket türünü (Tahsilat / Ödeme) seçiniz.", vbExclamation: Exit Sub
    End If
    tutar = Sayi(frm!txtTutar)
    If tutar <= 0 Then
        Mesaj "Tutarı giriniz.", vbExclamation: Exit Sub
    End If
    tahsilatMi = (InStr(Sade(frm!cmbHareketTuru), "tahsil") > 0)
    musteri = SayiVeyaNull(frm!cmbMusteri)
    acik = Nz(frm!txtAciklama, "")

    DBEngine.BeginTrans: islemde = True
    id = KasaYaz(CDate(frm!txtTarih), musteri, Null, IIf(tahsilatMi, "Tahsilat", "Ödeme"), _
        Nz(frm!cmbOdemeYontemi, ""), acik, IIf(tahsilatMi, tutar, 0), IIf(tahsilatMi, 0, tutar), _
        KAYNAK_MANUEL, Null, False, id)
    ' Bağlı cari kaydını yeniden yaz
    OtomatikKayitlariSil KAYNAK_KASA, id
    If Not IsNull(musteri) Then
        CariYaz CDate(frm!txtTarih), musteri, Null, IIf(tahsilatMi, "Kasa Tahsilat", "Kasa Ödeme"), _
            IIf(acik = "", "Kasa " & IIf(tahsilatMi, "tahsilatı", "ödemesi"), acik), _
            IIf(tahsilatMi, 0, tutar), IIf(tahsilatMi, tutar, 0), KAYNAK_KASA, id, True
    End If
    DBEngine.CommitTrans: islemde = False

    frm!txtKasaHareketID = id
    frm!lstHareketler.Requery
    KasaBakiyeGoster
    Mesaj IIf(guncelle, "Kasa hareketi güncellendi.", "Kasa hareketi kaydedildi.") & _
          IIf(IsNull(musteri), "", vbCrLf & "Müşterinin carisine de işlendi.")
    Exit Sub
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Sub

Public Function KasaListedenSec()
    On Error GoTo Hata
    Dim frm As Form, rs As DAO.Recordset, id As Long
    Set frm = F()
    id = ListeID(frm!lstHareketler)
    If id = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_KasaHareket WHERE KasaHareketID=" & id, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function
    frm!txtKasaHareketID = rs!KasaHareketID
    frm!txtTarih = rs!tarih
    frm!cmbMusteri = rs!musteriID
    If Nz(rs!tahsilat, 0) > 0 Then
        frm!cmbHareketTuru = "Tahsilat"
        frm!txtTutar = rs!tahsilat
    Else
        frm!cmbHareketTuru = "Ödeme"
        frm!txtTutar = rs!odeme
    End If
    frm!cmbOdemeYontemi = rs!odemeYontemi
    frm!txtAciklama = rs!aciklama
    If Nz(rs!OtomatikMi, False) Then
        Mesaj "Bu hareket '" & KaynakAdi(rs!kaynakTipi) & "' ekranından otomatik oluşturuldu; sadece görüntüleme içindir."
    End If
    rs.Close
    Exit Function
Hata:
    Mesaj "Hareket açılamadı: " & Err.Description, vbExclamation
End Function

Public Function KasaHareketSil()
    On Error GoTo Hata
    Dim frm As Form, id As Long, islemde As Boolean
    Set frm = F()
    If Bos(frm!txtKasaHareketID) Then
        Mesaj "Silmek için önce listeden bir kasa hareketi seçin.", vbExclamation
        Exit Function
    End If
    id = CLng(frm!txtKasaHareketID)
    If OtomatikKasaMi(id) Then Exit Function
    If Not Soru("Seçili kasa hareketi (ve bağlı cari kaydı) silinsin mi?") Then Exit Function
    DBEngine.BeginTrans: islemde = True
    OtomatikKayitlariSil KAYNAK_KASA, id
    CurrentDb.Execute "DELETE FROM tbl_KasaHareket WHERE KasaHareketID=" & id, dbFailOnError
    DBEngine.CommitTrans: islemde = False
    KasaTemizle
    frm!lstHareketler.Requery
    KasaBakiyeGoster
    Mesaj "Kasa hareketi silindi."
    Exit Function
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "Silme sırasında hata: " & Err.Description, vbExclamation
End Function

' İki tarih arası süz (biri boş bırakılabilir)
Public Function KasaFiltrele()
    On Error GoTo Hata
    Dim frm As Form, filtre As String
    Set frm = F()
    If TarihMi(frm!txtBaslangic) Then filtre = "TarihHam>=" & SqlTarih(CDate(frm!txtBaslangic))
    If TarihMi(frm!txtBitis) Then
        If filtre <> "" Then filtre = filtre & " AND "
        filtre = filtre & "TarihHam<=" & SqlTarih(CDate(frm!txtBitis))
    End If
    If filtre <> "" Then filtre = " WHERE " & filtre
    frm!lstHareketler.RowSource = "SELECT * FROM qry_KasaEkstre" & filtre & " ORDER BY TarihHam, KasaHareketID;"
    Exit Function
Hata:
    Mesaj "Filtreleme hatası: " & Err.Description, vbExclamation
End Function

Public Function KasaTumu()
    Dim frm As Form
    Set frm = F()
    frm!txtBaslangic = Null
    frm!txtBitis = Null
    KasaFiltrele
End Function

Public Function KasaAcilisGuncelle()
    On Error GoTo Hata
    Dim frm As Form
    Set frm = F()
    If Not SayiMi(frm!txtKasaAcilis) Then
        Mesaj "Açılış bakiyesini sayı olarak giriniz.", vbExclamation: Exit Function
    End If
    CurrentDb.Execute "UPDATE tbl_Kasa SET AcilisBakiyesi=" & Replace(CStr(CDbl(frm!txtKasaAcilis)), ",", ".") & _
        " WHERE KasaID=" & VarsayilanKasaID(), dbFailOnError
    frm!lstHareketler.Requery
    KasaBakiyeGoster
    Mesaj "Kasa açılış bakiyesi kaydedildi."
    Exit Function
Hata:
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Function