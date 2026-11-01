Option Compare Database
Option Explicit

' =============================================================
' CARİ HAREKETLER   (form: frm_CicerCari)
' -------------------------------------------------------------
' Tüm carilerin hareketleri listelenir; müşteri seçilince o cariye
' ait hareketler (borç / alacak / kalan), bidonları ve kart bilgileri
' (telefon, TCKN, adres, banka, IBAN) gösterilir.
' Hareket girişi Müşteri, Kasa, Parti ve Alım/Satım ekranlarından yapılır.
' =============================================================

Private Const FORM_ADI As String = "frm_CicerCari"

Private Function F() As Form
    Set F = Forms(FORM_ADI)
End Function

Public Function CariFiltrele()
    On Error GoTo Hata
    Dim frm As Form, filtre As String
    Set frm = F()
    If Not Bos(frm!cmbMusteriFiltre) Then filtre = " WHERE MusteriID=" & frm!cmbMusteriFiltre
    frm!lstHareketler.RowSource = "SELECT * FROM qry_CariEkstre" & filtre & " ORDER BY TarihHam, CariHareketID;"
    frm!lstBidonlar.RowSource = "SELECT * FROM qry_BidonListesi" & filtre & ";"
    If Bos(frm!cmbMusteriFiltre) Then
        frm!txtBakiye = Null
    Else
        frm!txtBakiye = MusteriBakiye(CLng(frm!cmbMusteriFiltre))
    End If
    CariBilgiGoster
    Exit Function
Hata:
    Mesaj "Filtreleme hatası: " & Err.Description, vbExclamation
End Function

Public Function CariTumu()
    F().Controls("cmbMusteriFiltre").Value = Null
    CariFiltrele
End Function

' Seçili carinin kart bilgilerini (adres, banka, IBAN...) gösterir
Private Sub CariBilgiGoster()
    On Error Resume Next
    Dim frm As Form, rs As DAO.Recordset
    Set frm = F()
    KutulariTemizle frm, Array("txtCariTelefon", "txtCariTCKN", "txtCariAdres", "txtCariBanka", "txtCariIBAN")
    If Bos(frm!cmbMusteriFiltre) Then Exit Sub
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Musteri WHERE MusteriID=" & CLng(frm!cmbMusteriFiltre), dbOpenSnapshot)
    If Not rs.EOF Then
        frm!txtCariTelefon = rs!Telefon
        frm!txtCariTCKN = rs!TCKN_VKN
        frm!txtCariAdres = rs!Adres
        frm!txtCariBanka = rs!BankaAdi
        frm!txtCariIBAN = rs.Fields("IBAN")
    End If
    rs.Close
End Sub