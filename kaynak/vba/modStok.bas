Option Compare Database
Option Explicit

' =============================================================
' STOK HAREKETLERİ   (form: frm_CicerStok)
' -------------------------------------------------------------
' Elle stok giriş / çıkış. Liste: Giren, Çıkan ve yürüyen Kalan
' (ürünün açılış stoğu + o satıra kadarki hareketler).
' Ürün seçilip "Filtrele" ile sadece o ürünün hareketleri görülür.
' Parti / Alım-Satım'dan gelen otomatik kayıtlar buradan
' değiştirilemez; kaynağından düzeltilir.
' =============================================================

Private Const FORM_ADI As String = "frm_CicerStok"
Private Const KUTULAR As String = "txtStokHareketID,cmbUrun,cmbMusteri,cmbStokTur,txtMiktar,txtAciklama"

Private Function F() As Form
    Set F = Forms(FORM_ADI)
End Function

Public Function StokTemizle()
    Dim frm As Form
    Set frm = F()
    KutulariTemizle frm, Split(KUTULAR, ",")
    frm!txtTarih = Date
    frm!lstHareketler = Null
End Function

Public Function StokYeni()
    StokTemizle
    F().Controls("cmbUrun").SetFocus
End Function

Public Function StokHareketKaydet()
    StokYazIslem False
End Function

Public Function StokHareketGuncelle()
    StokYazIslem True
End Function

Private Function OtomatikStokMu(ByVal id As Long) As Boolean
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT OtomatikMi, KaynakTipi FROM tbl_StokHareket WHERE StokHareketID=" & id, dbOpenSnapshot)
    If Not rs.EOF Then
        If Nz(rs!OtomatikMi, False) Then
            Mesaj "Bu stok hareketi '" & KaynakAdi(rs!kaynakTipi) & "' ekranından otomatik oluşturuldu." & vbCrLf & _
                  "Değiştirmek veya silmek için o ekrandan düzeltin.", vbExclamation
            OtomatikStokMu = True
        End If
    End If
    rs.Close
End Function

Private Sub StokYazIslem(ByVal guncelle As Boolean)
    On Error GoTo Hata
    Dim frm As Form, id As Long, girisMi As Boolean, miktar As Double, kalan As Double
    Set frm = F()
    If guncelle Then
        If Bos(frm!txtStokHareketID) Then
            Mesaj "Güncellemek için önce listeden bir stok hareketi seçin (çift tıklayın).", vbExclamation
            Exit Sub
        End If
        id = CLng(frm!txtStokHareketID)
        If OtomatikStokMu(id) Then Exit Sub
    ElseIf Not Bos(frm!txtStokHareketID) Then
        Mesaj "Bu hareket zaten kayıtlı. Değişiklik için 'Güncelle' düğmesini kullanın." & vbCrLf & _
              "Yeni hareket için önce 'Yeni Kayıt' düğmesine basın.", vbExclamation
        Exit Sub
    End If
    If Not TarihMi(frm!txtTarih) Then
        Mesaj "Tarih giriniz.", vbExclamation: Exit Sub
    End If
    If Bos(frm!cmbUrun) Then
        Mesaj "Ürün seçiniz.", vbExclamation: Exit Sub
    End If
    If Bos(frm!cmbStokTur) Then
        Mesaj "Hareket türünü (Giriş / Çıkış) seçiniz.", vbExclamation: Exit Sub
    End If
    miktar = Sayi(frm!txtMiktar)
    If miktar <= 0 Then
        Mesaj "Miktarı giriniz.", vbExclamation: Exit Sub
    End If
    girisMi = (InStr(Sade(frm!cmbStokTur), "giris") > 0)
    If Not girisMi Then
        kalan = UrunStok(CLng(frm!cmbUrun))
        If id > 0 Then kalan = kalan + Nz(DLookup("Nz(Cikis,0)-Nz(Giris,0)", "tbl_StokHareket", "StokHareketID=" & id), 0)
        If miktar > kalan Then
            If Not Soru("Mevcut stok " & kalan & ". Çıkış sonrası stok eksiye düşecek. Devam edilsin mi?") Then Exit Sub
        End If
    End If

    id = StokYaz(CDate(frm!txtTarih), CLng(frm!cmbUrun), SayiVeyaNull(frm!cmbMusteri), Null, _
        IIf(girisMi, "Giriş", "Çıkış"), Nz(frm!txtAciklama, ""), _
        IIf(girisMi, miktar, 0), IIf(girisMi, 0, miktar), KAYNAK_MANUEL, Null, False, id)

    frm!txtStokHareketID = id
    frm!lstHareketler.Requery
    StokMevcutGoster
    Mesaj IIf(guncelle, "Stok hareketi güncellendi.", "Stok hareketi kaydedildi.")
    Exit Sub
Hata:
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Sub

Public Function StokListedenSec()
    On Error GoTo Hata
    Dim frm As Form, rs As DAO.Recordset, id As Long
    Set frm = F()
    id = ListeID(frm!lstHareketler)
    If id = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_StokHareket WHERE StokHareketID=" & id, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function
    frm!txtStokHareketID = rs!StokHareketID
    frm!txtTarih = rs!tarih
    frm!cmbUrun = rs!urunID
    frm!cmbMusteri = rs!musteriID
    If Nz(rs!giris, 0) > 0 Then
        frm!cmbStokTur = "Giriş"
        frm!txtMiktar = rs!giris
    Else
        frm!cmbStokTur = "Çıkış"
        frm!txtMiktar = rs!cikis
    End If
    frm!txtAciklama = rs!aciklama
    If Nz(rs!OtomatikMi, False) Then
        Mesaj "Bu hareket '" & KaynakAdi(rs!kaynakTipi) & "' ekranından otomatik oluşturuldu; sadece görüntüleme içindir."
    End If
    rs.Close
    Exit Function
Hata:
    Mesaj "Hareket açılamadı: " & Err.Description, vbExclamation
End Function

Public Function StokHareketSil()
    On Error GoTo Hata
    Dim frm As Form, id As Long
    Set frm = F()
    If Bos(frm!txtStokHareketID) Then
        Mesaj "Silmek için önce listeden bir stok hareketi seçin.", vbExclamation
        Exit Function
    End If
    id = CLng(frm!txtStokHareketID)
    If OtomatikStokMu(id) Then Exit Function
    If Not Soru("Seçili stok hareketi silinsin mi?") Then Exit Function
    CurrentDb.Execute "DELETE FROM tbl_StokHareket WHERE StokHareketID=" & id, dbFailOnError
    StokTemizle
    frm!lstHareketler.Requery
    StokMevcutGoster
    Mesaj "Stok hareketi silindi."
    Exit Function
Hata:
    Mesaj "Silme sırasında hata: " & Err.Description, vbExclamation
End Function

Public Function StokFiltrele()
    On Error GoTo Hata
    Dim frm As Form, filtre As String
    Set frm = F()
    If Not Bos(frm!cmbUrunFiltre) Then filtre = " WHERE UrunID=" & frm!cmbUrunFiltre
    frm!lstHareketler.RowSource = "SELECT * FROM qry_StokEkstre" & filtre & " ORDER BY TarihHam, StokHareketID;"
    StokMevcutGoster
    Exit Function
Hata:
    Mesaj "Filtreleme hatası: " & Err.Description, vbExclamation
End Function

Public Function StokTumu()
    F().Controls("cmbUrunFiltre").Value = Null
    StokFiltrele
End Function

Private Sub StokMevcutGoster()
    On Error Resume Next
    Dim frm As Form
    Set frm = F()
    If Bos(frm!cmbUrunFiltre) Then
        frm!txtMevcutStok = Null
    Else
        frm!txtMevcutStok = UrunStok(CLng(frm!cmbUrunFiltre))
    End If
End Sub