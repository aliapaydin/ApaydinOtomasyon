Option Compare Database
Option Explicit

' =============================================================
' ALIM / SATIM   (form: frm_CicerAlimSatim)
' -------------------------------------------------------------
' Toplam = Miktar x Birim Fiyat
' Otomatik kayıtlar:
'   ALIM  -> Stok GİRİŞ, cariye ALACAK (satıcıya borçlanırız)
'            Peşin ödendiyse ayrıca cariye BORÇ + Kasa ÖDEME
'   SATIM -> Stok ÇIKIŞ, cariye BORÇ (müşteri bize borçlanır)
'            Peşin tahsil edildiyse ayrıca cariye ALACAK + Kasa TAHSİLAT
' Güncelleme eski otomatik kayıtları silip yeniden yazar.
' =============================================================

Private Const FORM_ADI As String = "frm_CicerAlimSatim"
Private Const KUTULAR As String = "txtIslemID,cmbIslemTuru,cmbMusteri,cmbUrun,txtMiktar,txtBirimFiyat,txtToplam,txtAciklama"

Private Function F() As Form
    Set F = Forms(FORM_ADI)
End Function

Private Function AlimMi(ByVal v As Variant) As Boolean
    AlimMi = (Left$(Sade(v), 3) = "ali")
End Function

Public Function AlimSatimTemizle()
    Dim frm As Form
    Set frm = F()
    KutulariTemizle frm, Split(KUTULAR, ",")
    frm!txtTarih = Date
    frm!chkPesin = False
    frm!lstIslemler = Null
End Function

Public Function AlimSatimYeni()
    AlimSatimTemizle
    F().Controls("cmbIslemTuru").SetFocus
End Function

Public Function AlimSatimHesapla()
    Dim frm As Form
    Set frm = F()
    If SayiMi(frm!txtMiktar) And SayiMi(frm!txtBirimFiyat) Then
        frm!txtToplam = Round(Sayi(frm!txtMiktar) * Sayi(frm!txtBirimFiyat), 2)
    Else
        frm!txtToplam = Null
    End If
End Function

' İşlem türü veya ürün değişince ürün kartındaki alış / satış fiyatını getirir
Public Function AlimSatimSecimDegisti()
    Dim frm As Form, fiyat As Variant
    Set frm = F()
    If Not Bos(frm!cmbUrun) And Not Bos(frm!cmbIslemTuru) Then
        fiyat = DLookup(IIf(AlimMi(frm!cmbIslemTuru), "AlisFiyati", "SatisFiyati"), "tbl_Urun", "UrunID=" & frm!cmbUrun)
        If Nz(fiyat, 0) > 0 Then frm!txtBirimFiyat = fiyat
    End If
    AlimSatimHesapla
End Function

Public Function AlimSatimKaydet()
    AlimSatimYazIslem False
End Function

Public Function AlimSatimGuncelle()
    AlimSatimYazIslem True
End Function

Private Sub AlimSatimYazIslem(ByVal guncelle As Boolean)
    On Error GoTo Hata
    Dim frm As Form, id As Long, rs As DAO.Recordset, islemde As Boolean
    Set frm = F()
    If guncelle Then
        If Bos(frm!txtIslemID) Then
            Mesaj "Güncellemek için önce listeden bir işlem seçin (çift tıklayın).", vbExclamation
            Exit Sub
        End If
        id = CLng(frm!txtIslemID)
    ElseIf Not Bos(frm!txtIslemID) Then
        Mesaj "Bu işlem zaten kayıtlı. Değişiklik için 'Güncelle' düğmesini kullanın." & vbCrLf & _
              "Yeni işlem için önce 'Yeni Kayıt' düğmesine basın.", vbExclamation
        Exit Sub
    End If
    AlimSatimHesapla
    If Not TarihMi(frm!txtTarih) Then
        Mesaj "Tarih giriniz.", vbExclamation: Exit Sub
    End If
    If Bos(frm!cmbIslemTuru) Then
        Mesaj "İşlem türünü (Alım / Satım) seçiniz.", vbExclamation: Exit Sub
    End If
    If Bos(frm!cmbMusteri) Then
        Mesaj "Cari seçiniz.", vbExclamation: Exit Sub
    End If
    If Bos(frm!cmbUrun) Then
        Mesaj "Ürün seçiniz.", vbExclamation: Exit Sub
    End If
    If Sayi(frm!txtMiktar) <= 0 Or Sayi(frm!txtBirimFiyat) <= 0 Then
        Mesaj "Miktar ve birim fiyatı giriniz.", vbExclamation: Exit Sub
    End If

    DBEngine.BeginTrans: islemde = True
    Set rs = KayitAc("tbl_AlimSatim", "IslemID", id)
    rs!tarih = CDate(frm!txtTarih)
    rs!IslemTuru = IIf(AlimMi(frm!cmbIslemTuru), "Alım", "Satım")
    rs!musteriID = CLng(frm!cmbMusteri)
    rs!urunID = CLng(frm!cmbUrun)
    rs!miktar = Sayi(frm!txtMiktar)
    rs!BirimFiyat = Sayi(frm!txtBirimFiyat)
    rs!ToplamTutar = Sayi(frm!txtToplam)
    rs!PesinOdendi = Nz(frm!chkPesin, False)
    rs!aciklama = MetinVeyaNull(frm!txtAciklama)
    rs!IptalMi = False
    rs.Update
    rs.Bookmark = rs.LastModified
    id = rs!IslemID
    rs.Close
    AlimSatimMuhasebeYaz id
    DBEngine.CommitTrans: islemde = False

    frm!txtIslemID = id
    frm!lstIslemler.Requery
    Mesaj IIf(guncelle, "İşlem güncellendi.", "İşlem kaydedildi.") & vbCrLf & _
          "Stok / cari / kasa kayıtları otomatik oluşturuldu."
    Exit Sub
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Sub

Public Sub AlimSatimMuhasebeYaz(ByVal id As Long)
    Dim rs As DAO.Recordset, alim As Boolean, acik As String, tutar As Currency
    Set rs = CurrentDb.OpenRecordset("SELECT a.*, u.UrunAdi FROM tbl_AlimSatim AS a INNER JOIN tbl_Urun AS u ON a.UrunID=u.UrunID WHERE a.IslemID=" & id, dbOpenSnapshot)
    OtomatikKayitlariSil KAYNAK_ALIMSATIM, id
    If rs.EOF Then rs.Close: Exit Sub
    alim = AlimMi(rs!IslemTuru)
    tutar = Nz(rs!ToplamTutar, 0)
    acik = rs!IslemTuru & ": " & rs!miktar & " x " & rs!UrunAdi & IIf(Bos(rs!aciklama), "", " - " & rs!aciklama)

    If alim Then
        StokYaz rs!tarih, rs!urunID, rs!musteriID, Null, "Alım", acik, rs!miktar, 0, KAYNAK_ALIMSATIM, id, True
        CariYaz rs!tarih, rs!musteriID, Null, "Alım", acik, 0, tutar, KAYNAK_ALIMSATIM, id, True
        If Nz(rs!PesinOdendi, False) Then
            CariYaz rs!tarih, rs!musteriID, Null, "Peşin Ödeme", acik, tutar, 0, KAYNAK_ALIMSATIM, id, True
            KasaYaz rs!tarih, rs!musteriID, Null, "Ödeme", "Nakit", acik, 0, tutar, KAYNAK_ALIMSATIM, id, True
        End If
    Else
        StokYaz rs!tarih, rs!urunID, rs!musteriID, Null, "Satım", acik, 0, rs!miktar, KAYNAK_ALIMSATIM, id, True
        CariYaz rs!tarih, rs!musteriID, Null, "Satım", acik, tutar, 0, KAYNAK_ALIMSATIM, id, True
        If Nz(rs!PesinOdendi, False) Then
            CariYaz rs!tarih, rs!musteriID, Null, "Peşin Tahsilat", acik, 0, tutar, KAYNAK_ALIMSATIM, id, True
            KasaYaz rs!tarih, rs!musteriID, Null, "Tahsilat", "Nakit", acik, tutar, 0, KAYNAK_ALIMSATIM, id, True
        End If
    End If
    rs.Close
End Sub

Public Function AlimSatimListedenSec()
    On Error GoTo Hata
    Dim frm As Form, rs As DAO.Recordset, id As Long
    Set frm = F()
    id = ListeID(frm!lstIslemler)
    If id = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_AlimSatim WHERE IslemID=" & id, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function
    frm!txtIslemID = rs!IslemID
    frm!txtTarih = rs!tarih
    frm!cmbIslemTuru = rs!IslemTuru
    frm!cmbMusteri = rs!musteriID
    frm!cmbUrun = rs!urunID
    frm!txtMiktar = rs!miktar
    frm!txtBirimFiyat = rs!BirimFiyat
    frm!txtToplam = rs!ToplamTutar
    frm!chkPesin = Nz(rs!PesinOdendi, False)
    frm!txtAciklama = rs!aciklama
    rs.Close
    Exit Function
Hata:
    Mesaj "İşlem açılamadı: " & Err.Description, vbExclamation
End Function

Public Function AlimSatimSil()
    On Error GoTo Hata
    Dim frm As Form, id As Long, islemde As Boolean
    Set frm = F()
    If Bos(frm!txtIslemID) Then
        Mesaj "Silmek için önce listeden bir işlem seçin.", vbExclamation
        Exit Function
    End If
    id = CLng(frm!txtIslemID)
    If Not Soru("Seçili işlem ve otomatik stok / cari / kasa kayıtları silinsin mi?") Then Exit Function
    DBEngine.BeginTrans: islemde = True
    OtomatikKayitlariSil KAYNAK_ALIMSATIM, id
    CurrentDb.Execute "DELETE FROM tbl_AlimSatim WHERE IslemID=" & id, dbFailOnError
    DBEngine.CommitTrans: islemde = False
    AlimSatimTemizle
    frm!lstIslemler.Requery
    Mesaj "İşlem silindi."
    Exit Function
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "Silme sırasında hata: " & Err.Description, vbExclamation
End Function