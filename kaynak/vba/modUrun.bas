Option Compare Database
Option Explicit

' =============================================================
' ÜRÜN / STOK KARTLARI   (form: frm_CicerUrun)
' -------------------------------------------------------------
' "HY" kodlu Hak Yağı ürünü hakyağı stoğu için kullanılır; silinmemeli.
' Hareketi olan ürün silinemez, "Aktif" işareti kaldırılarak
' listelerden gizlenebilir.
' =============================================================

Private Const FORM_ADI As String = "frm_CicerUrun"
Private Const KUTULAR As String = "txtUrunID,txtUrunKodu,txtUrunAdi,cmbBirim,txtAcilisStok,txtMinimumStok,txtAlisFiyati,txtSatisFiyati"

Private Function F() As Form
    Set F = Forms(FORM_ADI)
End Function

Public Function UrunTemizle()
    Dim frm As Form
    Set frm = F()
    KutulariTemizle frm, Split(KUTULAR, ",")
    frm!cmbBirim = "Kg"
    frm!txtAcilisStok = 0
    frm!chkAktif = True
    frm!lstUrunler = Null
End Function

Public Function UrunYeni()
    UrunTemizle
    F().Controls("txtUrunKodu").SetFocus
End Function

Public Function UrunKaydet()
    UrunYazIslem False
End Function

Public Function UrunGuncelle()
    UrunYazIslem True
End Function

Private Sub UrunYazIslem(ByVal guncelle As Boolean)
    On Error GoTo Hata
    Dim frm As Form, id As Long, rs As DAO.Recordset, ad As Variant
    Set frm = F()
    If guncelle Then
        If Bos(frm!txtUrunID) Then
            Mesaj "Güncellemek için önce listeden bir ürün seçin (çift tıklayın).", vbExclamation
            Exit Sub
        End If
        id = CLng(frm!txtUrunID)
    ElseIf Not Bos(frm!txtUrunID) Then
        Mesaj "Bu ürün zaten kayıtlı. Değişiklik için 'Güncelle' düğmesini kullanın." & vbCrLf & _
              "Yeni ürün için önce 'Yeni Kayıt' düğmesine basın.", vbExclamation
        Exit Sub
    End If
    If Bos(frm!txtUrunAdi) Then
        Mesaj "Ürün adını giriniz.", vbExclamation: Exit Sub
    End If
    If Not Bos(frm!txtUrunKodu) Then
        If DCount("*", "tbl_Urun", "UrunKodu=" & SqlMetin(Trim$(frm!txtUrunKodu)) & " AND UrunID<>" & id) > 0 Then
            Mesaj "Bu ürün kodu başka bir üründe kullanılıyor.", vbExclamation: Exit Sub
        End If
    End If
    For Each ad In Array("txtAcilisStok", "txtMinimumStok", "txtAlisFiyati", "txtSatisFiyati")
        If Not Bos(frm.Controls(CStr(ad)).Value) And Not SayiMi(frm.Controls(CStr(ad)).Value) Then
            Mesaj "Stok ve fiyat alanları sayı olmalıdır.", vbExclamation: Exit Sub
        End If
    Next

    Set rs = KayitAc("tbl_Urun", "UrunID", id)
    rs!UrunKodu = MetinVeyaNull(frm!txtUrunKodu)
    rs!UrunAdi = MetinVeyaNull(frm!txtUrunAdi)
    rs!Birim = MetinVeyaNull(frm!cmbBirim)
    rs!AcilisStok = Sayi(frm!txtAcilisStok)
    rs!MinimumStok = SayiVeyaNull(frm!txtMinimumStok)
    rs!AlisFiyati = SayiVeyaNull(frm!txtAlisFiyati)
    rs!SatisFiyati = SayiVeyaNull(frm!txtSatisFiyati)
    rs!Aktif = Nz(frm!chkAktif, True)
    rs.Update
    rs.Bookmark = rs.LastModified
    id = rs!urunID
    rs.Close

    frm!txtUrunID = id
    frm!lstUrunler.Requery
    KombolariYenile
    Mesaj IIf(guncelle, "Ürün güncellendi.", "Ürün kaydedildi.")
    Exit Sub
Hata:
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Sub

Public Function UrunListedenSec()
    On Error GoTo Hata
    Dim frm As Form, rs As DAO.Recordset, id As Long
    Set frm = F()
    id = ListeID(frm!lstUrunler)
    If id = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Urun WHERE UrunID=" & id, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function
    frm!txtUrunID = rs!urunID
    frm!txtUrunKodu = rs!UrunKodu
    frm!txtUrunAdi = rs!UrunAdi
    frm!cmbBirim = rs!Birim
    frm!txtAcilisStok = Nz(rs!AcilisStok, 0)
    frm!txtMinimumStok = rs!MinimumStok
    frm!txtAlisFiyati = rs!AlisFiyati
    frm!txtSatisFiyati = rs!SatisFiyati
    frm!chkAktif = Nz(rs!Aktif, True)
    rs.Close
    Exit Function
Hata:
    Mesaj "Ürün açılamadı: " & Err.Description, vbExclamation
End Function

Public Function UrunSil()
    On Error GoTo Hata
    Dim frm As Form, id As Long, n As Long
    Set frm = F()
    If Bos(frm!txtUrunID) Then
        Mesaj "Silmek için önce listeden bir ürün seçin.", vbExclamation
        Exit Function
    End If
    id = CLng(frm!txtUrunID)
    If id = HakYagiUrunID() Then
        Mesaj "Hak Yağı ürünü hakyağı stoğu için kullanılıyor, silinemez.", vbExclamation
        Exit Function
    End If
    n = DCount("*", "tbl_StokHareket", "UrunID=" & id) + DCount("*", "tbl_AlimSatim", "UrunID=" & id)
    If n > 0 Then
        Mesaj "Bu ürünün " & n & " hareketi var, silinemez." & vbCrLf & _
              "Kullanılmasın istiyorsanız 'Aktif' işaretini kaldırıp Güncelle'ye basın.", vbExclamation
        Exit Function
    End If
    If Not Soru(Nz(frm!txtUrunAdi, "") & " silinsin mi?") Then Exit Function
    CurrentDb.Execute "DELETE FROM tbl_Urun WHERE UrunID=" & id, dbFailOnError
    UrunTemizle
    frm!lstUrunler.Requery
    KombolariYenile
    Mesaj "Ürün silindi."
    Exit Function
Hata:
    Mesaj "Silme sırasında hata: " & Err.Description, vbExclamation
End Function