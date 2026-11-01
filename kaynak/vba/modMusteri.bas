Option Compare Database
Option Explicit

' =============================================================
' MÜŞTERİ YÖNETİMİ   (form: frm_CicerMusteri)
' -------------------------------------------------------------
' Üst bölüm : müşteri kartı (TCKN, ad, telefon, adres, banka, IBAN...)
' Alt bölüm : seçili müşterinin cari hareketleri (borç / alacak / kalan)
'             ve bidonları. Müşteri seçilmemişse tüm hareketler görünür.
' Elle cari hareket (açılış, düzeltme vb.) buradan girilir.
' Otomatik hareketler (Parti, Kasa, Alım/Satım kaynaklı) buradan
' değiştirilemez; kaynağından düzeltilir.
' =============================================================

Private Const FORM_ADI As String = "frm_CicerMusteri"
Private Const MUSTERI_KUTULARI As String = "txtMusteriID,txtTCKN,txtAdSoyad,txtTelefon,txtAdres,txtBanka,txtIBAN,cmbCariTipi,txtAcilisBakiyesi"
Private Const CARI_KUTULARI As String = "txtCariHareketID,cmbCariYon,txtCariTutar,txtCariAciklama"

Private Function F() As Form
    Set F = Forms(FORM_ADI)
End Function

' ---------------------------------------------------------------
' Müşteri kartı
' ---------------------------------------------------------------
Public Function MusteriTemizle()
    Dim frm As Form
    Set frm = F()
    KutulariTemizle frm, Split(MUSTERI_KUTULARI, ",")
    frm!cmbCariTipi = "Alım/Satım"
    frm!txtAcilisBakiyesi = 0
    frm!lstMusteriler = Null
    CariDuzeltmeTemizle
    MusteriCariYenile
End Function

Public Function MusteriYeni()
    MusteriTemizle
    F().Controls("txtTCKN").SetFocus
End Function

Public Function MusteriKaydet()
    MusteriYazIslem False
End Function

Public Function MusteriGuncelle()
    MusteriYazIslem True
End Function

Private Function IbanSade(ByVal v As Variant) As String
    IbanSade = UCase$(Replace(Nz(deger(v), ""), " ", ""))
End Function

Private Function MusteriKontrol(ByVal id As Long) As Boolean
    Dim frm As Form, tc As String, hesapNo As String, i As Long, hepsiRakam As Boolean
    Set frm = F()
    If Bos(frm!txtAdSoyad) Then
        Mesaj "Ad Soyad / Ünvan giriniz.", vbExclamation: Exit Function
    End If
    tc = Trim$(Nz(frm!txtTCKN, ""))
    If tc <> "" Then
        hepsiRakam = True
        For i = 1 To Len(tc)
            If Mid$(tc, i, 1) < "0" Or Mid$(tc, i, 1) > "9" Then hepsiRakam = False
        Next
        If Not hepsiRakam Or (Len(tc) <> 10 And Len(tc) <> 11) Then
            If Not Soru("TCKN 11, VKN 10 haneli olmalıdır. Yine de kaydedilsin mi?") Then Exit Function
        End If
        If DCount("*", "tbl_Musteri", "TCKN_VKN=" & SqlMetin(tc) & " AND MusteriID<>" & id) > 0 Then
            If Not Soru("Bu TCKN/VKN başka bir müşteride de kayıtlı. Yine de kaydedilsin mi?") Then Exit Function
        End If
    End If
    hesapNo = IbanSade(frm!txtIBAN)
    If hesapNo <> "" Then
        If Left$(hesapNo, 2) <> "TR" Or Len(hesapNo) <> 26 Then
            If Not Soru("IBAN 'TR' ile başlayan 26 karakter olmalıdır. Yine de kaydedilsin mi?") Then Exit Function
        End If
    End If
    If Not Bos(frm!txtAcilisBakiyesi) And Not SayiMi(frm!txtAcilisBakiyesi) Then
        Mesaj "Açılış bakiyesi sayı olmalıdır.", vbExclamation: Exit Function
    End If
    MusteriKontrol = True
End Function

Private Sub MusteriYazIslem(ByVal guncelle As Boolean)
    On Error GoTo Hata
    Dim frm As Form, id As Long, rs As DAO.Recordset
    Set frm = F()
    If guncelle Then
        If Bos(frm!txtMusteriID) Then
            Mesaj "Güncellemek için önce listeden bir müşteri seçin (çift tıklayın).", vbExclamation
            Exit Sub
        End If
        id = CLng(frm!txtMusteriID)
    ElseIf Not Bos(frm!txtMusteriID) Then
        Mesaj "Bu müşteri zaten kayıtlı. Değişiklik için 'Güncelle' düğmesini kullanın." & vbCrLf & _
              "Yeni müşteri için önce 'Yeni Kayıt' düğmesine basın.", vbExclamation
        Exit Sub
    End If
    If Not MusteriKontrol(id) Then Exit Sub

    Set rs = KayitAc("tbl_Musteri", "MusteriID", id)
    rs!TCKN_VKN = MetinVeyaNull(frm!txtTCKN)
    rs!adSoyad = MetinVeyaNull(frm!txtAdSoyad)
    rs!Telefon = MetinVeyaNull(frm!txtTelefon)
    rs!Adres = MetinVeyaNull(frm!txtAdres)
    rs!BankaAdi = MetinVeyaNull(frm!txtBanka)
    rs.Fields("IBAN") = MetinVeyaNull(IbanSade(frm!txtIBAN))
    rs!CariTipi = MetinVeyaNull(frm!cmbCariTipi)
    rs!AcilisBakiyesi = Sayi(frm!txtAcilisBakiyesi)
    rs.Update
    rs.Bookmark = rs.LastModified
    id = rs!musteriID
    rs.Close

    frm!txtMusteriID = id
    frm!txtIBAN = MetinVeyaNull(IbanSade(frm!txtIBAN))
    frm!lstMusteriler.Requery
    KombolariYenile
    MusteriCariYenile
    Mesaj IIf(guncelle, "Müşteri güncellendi.", "Müşteri kaydedildi.")
    Exit Sub
Hata:
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Sub

Public Function MusteriSil()
    On Error GoTo Hata
    Dim frm As Form, id As Long, n As Long, k As String
    Set frm = F()
    If Bos(frm!txtMusteriID) Then
        Mesaj "Silmek için önce listeden bir müşteri seçin.", vbExclamation
        Exit Function
    End If
    id = CLng(frm!txtMusteriID)
    k = "MusteriID=" & id
    n = DCount("*", "tbl_Parti", k) + DCount("*", "tbl_CariHareket", k) + DCount("*", "tbl_KasaHareket", k) + _
        DCount("*", "tbl_StokHareket", k) + DCount("*", "tbl_AlimSatim", k)
    If n > 0 Then
        Mesaj "Bu müşterinin parti veya hareket kayıtları var (" & n & " adet), silinemez.", vbExclamation
        Exit Function
    End If
    If Not Soru(Nz(frm!txtAdSoyad, "") & " silinsin mi?") Then Exit Function
    CurrentDb.Execute "DELETE FROM tbl_Musteri WHERE " & k, dbFailOnError
    MusteriTemizle
    frm!lstMusteriler.Requery
    KombolariYenile
    Mesaj "Müşteri silindi."
    Exit Function
Hata:
    Mesaj "Silme sırasında hata: " & Err.Description, vbExclamation
End Function

Private Sub MusteriYukle(ByVal id As Long)
    Dim frm As Form, rs As DAO.Recordset
    Set frm = F()
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Musteri WHERE MusteriID=" & id, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Sub
    frm!txtMusteriID = rs!musteriID
    frm!txtTCKN = rs!TCKN_VKN
    frm!txtAdSoyad = rs!adSoyad
    frm!txtTelefon = rs!Telefon
    frm!txtAdres = rs!Adres
    frm!txtBanka = rs!BankaAdi
    frm!txtIBAN = rs.Fields("IBAN")
    frm!cmbCariTipi = rs!CariTipi
    frm!txtAcilisBakiyesi = Nz(rs!AcilisBakiyesi, 0)
    rs.Close
    CariDuzeltmeTemizle
    MusteriCariYenile
End Sub

Public Function MusteriListedenSec()
    On Error GoTo Hata
    Dim id As Long
    id = ListeID(F().Controls("lstMusteriler"))
    If id > 0 Then MusteriYukle id
    Exit Function
Hata:
    Mesaj "Müşteri açılamadı: " & Err.Description, vbExclamation
End Function

' Alt bölümdeki cari hareket ve bidon listelerini seçili müşteriye göre süzer
Public Function MusteriCariYenile()
    On Error Resume Next
    Dim frm As Form, filtre As String
    Set frm = F()
    If Not Bos(frm!txtMusteriID) Then filtre = " WHERE MusteriID=" & frm!txtMusteriID
    frm!lstCari.RowSource = "SELECT * FROM qry_CariEkstre" & filtre & " ORDER BY TarihHam, CariHareketID;"
    frm!lstBidonlar.RowSource = "SELECT * FROM qry_BidonListesi" & filtre & ";"
    If Bos(frm!txtMusteriID) Then
        frm!txtBakiye = Null
    Else
        frm!txtBakiye = MusteriBakiye(CLng(frm!txtMusteriID))
    End If
End Function

' ---------------------------------------------------------------
' Elle cari hareket (açılış bakiyesi, düzeltme, mahsup...)
' ---------------------------------------------------------------
Public Function CariDuzeltmeTemizle()
    Dim frm As Form
    Set frm = F()
    KutulariTemizle frm, Split(CARI_KUTULARI, ",")
    frm!txtCariTarih = Date
    frm!lstCari = Null
End Function

Public Function CariDuzeltmeKaydet()
    CariYazIslem False
End Function

Public Function CariDuzeltmeGuncelle()
    CariYazIslem True
End Function

Private Function OtomatikCariMi(ByVal id As Long) As Boolean
    Dim tip As Variant
    If Nz(DLookup("OtomatikMi", "tbl_CariHareket", "CariHareketID=" & id), False) Then
        tip = DLookup("KaynakTipi", "tbl_CariHareket", "CariHareketID=" & id)
        Mesaj "Bu hareket '" & KaynakAdi(tip) & "' ekranından otomatik oluşturuldu." & vbCrLf & _
              "Değiştirmek veya silmek için o ekrandan düzeltin.", vbExclamation
        OtomatikCariMi = True
    End If
End Function

Private Sub CariYazIslem(ByVal guncelle As Boolean)
    On Error GoTo Hata
    Dim frm As Form, id As Long, borcMu As Boolean, tutar As Currency
    Set frm = F()
    If guncelle Then
        If Bos(frm!txtCariHareketID) Then
            Mesaj "Güncellemek için önce listeden bir cari hareket seçin (çift tıklayın).", vbExclamation
            Exit Sub
        End If
        id = CLng(frm!txtCariHareketID)
        If OtomatikCariMi(id) Then Exit Sub
    ElseIf Not Bos(frm!txtCariHareketID) Then
        Mesaj "Bu hareket zaten kayıtlı. Değişiklik için 'Hareket Güncelle' düğmesini kullanın." & vbCrLf & _
              "Yeni hareket için önce 'Hareket Temizle' düğmesine basın.", vbExclamation
        Exit Sub
    End If
    If Bos(frm!txtMusteriID) Then
        Mesaj "Önce listeden bir müşteri seçin.", vbExclamation: Exit Sub
    End If
    If Not TarihMi(frm!txtCariTarih) Then
        Mesaj "Hareket tarihini giriniz.", vbExclamation: Exit Sub
    End If
    If Bos(frm!cmbCariYon) Then
        Mesaj "Borç veya Alacak seçiniz.", vbExclamation: Exit Sub
    End If
    tutar = Sayi(frm!txtCariTutar)
    If tutar <= 0 Then
        Mesaj "Tutarı giriniz.", vbExclamation: Exit Sub
    End If
    borcMu = (InStr(Sade(frm!cmbCariYon), "borc") > 0)

    id = CariYaz(CDate(frm!txtCariTarih), CLng(frm!txtMusteriID), Null, _
        IIf(borcMu, "Manuel Borç", "Manuel Alacak"), Nz(frm!txtCariAciklama, ""), _
        IIf(borcMu, tutar, 0), IIf(borcMu, 0, tutar), KAYNAK_MANUEL, Null, False, id)

    frm!txtCariHareketID = id
    MusteriCariYenile
    frm!lstMusteriler.Requery
    Mesaj IIf(guncelle, "Cari hareket güncellendi.", "Cari hareket kaydedildi.")
    Exit Sub
Hata:
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Sub

Public Function CariListedenSec()
    On Error GoTo Hata
    Dim frm As Form, rs As DAO.Recordset, id As Long
    Set frm = F()
    id = ListeID(frm!lstCari)
    If id = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_CariHareket WHERE CariHareketID=" & id, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function
    If Nz(frm!txtMusteriID, 0) <> rs!musteriID Then MusteriYukle rs!musteriID
    frm!txtCariHareketID = rs!CariHareketID
    frm!txtCariTarih = rs!tarih
    If Nz(rs!borc, 0) > 0 Then
        frm!cmbCariYon = "Borç"
        frm!txtCariTutar = rs!borc
    Else
        frm!cmbCariYon = "Alacak"
        frm!txtCariTutar = rs!alacak
    End If
    frm!txtCariAciklama = rs!aciklama
    If Nz(rs!OtomatikMi, False) Then
        Mesaj "Bu hareket '" & KaynakAdi(rs!kaynakTipi) & "' ekranından otomatik oluşturuldu; sadece görüntüleme içindir."
    End If
    rs.Close
    Exit Function
Hata:
    Mesaj "Hareket açılamadı: " & Err.Description, vbExclamation
End Function

Public Function CariHareketSil()
    On Error GoTo Hata
    Dim frm As Form, id As Long
    Set frm = F()
    If Bos(frm!txtCariHareketID) Then
        Mesaj "Silmek için önce listeden bir cari hareket seçin.", vbExclamation
        Exit Function
    End If
    id = CLng(frm!txtCariHareketID)
    If OtomatikCariMi(id) Then Exit Function
    If Not Soru("Seçili cari hareket silinsin mi?") Then Exit Function
    CurrentDb.Execute "DELETE FROM tbl_CariHareket WHERE CariHareketID=" & id, dbFailOnError
    CariDuzeltmeTemizle
    MusteriCariYenile
    frm!lstMusteriler.Requery
    Mesaj "Cari hareket silindi."
    Exit Function
Hata:
    Mesaj "Silme sırasında hata: " & Err.Description, vbExclamation
End Function