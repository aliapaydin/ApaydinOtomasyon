Option Compare Database
Option Explicit

' =============================================================
' PERSONEL   (form: frm_CicerPersonel)
' Parti ekranındaki Giriş / Teslim personeli listeleri buradan beslenir.
' =============================================================

Private Const FORM_ADI As String = "frm_CicerPersonel"
Private Const KUTULAR As String = "txtPersonelID,txtAdSoyad,txtGorev"

Private Function F() As Form
    Set F = Forms(FORM_ADI)
End Function

Public Function PersonelTemizle()
    Dim frm As Form
    Set frm = F()
    KutulariTemizle frm, Split(KUTULAR, ",")
    frm!lstPersonel = Null
End Function

Public Function PersonelYeni()
    PersonelTemizle
    F().Controls("txtAdSoyad").SetFocus
End Function

Public Function PersonelKaydet()
    PersonelYazIslem False
End Function

Public Function PersonelGuncelle()
    PersonelYazIslem True
End Function

Private Sub PersonelYazIslem(ByVal guncelle As Boolean)
    On Error GoTo Hata
    Dim frm As Form, id As Long, rs As DAO.Recordset
    Set frm = F()
    If guncelle Then
        If Bos(frm!txtPersonelID) Then
            Mesaj "Güncellemek için önce listeden bir personel seçin (çift tıklayın).", vbExclamation
            Exit Sub
        End If
        id = CLng(frm!txtPersonelID)
    ElseIf Not Bos(frm!txtPersonelID) Then
        Mesaj "Bu personel zaten kayıtlı. Değişiklik için 'Güncelle' düğmesini kullanın." & vbCrLf & _
              "Yeni personel için önce 'Yeni Kayıt' düğmesine basın.", vbExclamation
        Exit Sub
    End If
    If Bos(frm!txtAdSoyad) Then
        Mesaj "Ad Soyad giriniz.", vbExclamation: Exit Sub
    End If
    Set rs = KayitAc("tbl_Personel", "PersonelID", id)
    rs!adSoyad = MetinVeyaNull(frm!txtAdSoyad)
    rs!Gorev = MetinVeyaNull(frm!txtGorev)
    rs.Update
    rs.Bookmark = rs.LastModified
    id = rs!PersonelID
    rs.Close
    frm!txtPersonelID = id
    frm!lstPersonel.Requery
    KombolariYenile
    Mesaj IIf(guncelle, "Personel güncellendi.", "Personel kaydedildi.")
    Exit Sub
Hata:
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Sub

Public Function PersonelListedenSec()
    On Error GoTo Hata
    Dim frm As Form, rs As DAO.Recordset, id As Long
    Set frm = F()
    id = ListeID(frm!lstPersonel)
    If id = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Personel WHERE PersonelID=" & id, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function
    frm!txtPersonelID = rs!PersonelID
    frm!txtAdSoyad = rs!adSoyad
    frm!txtGorev = rs!Gorev
    rs.Close
    Exit Function
Hata:
    Mesaj "Personel açılamadı: " & Err.Description, vbExclamation
End Function

Public Function PersonelSil()
    On Error GoTo Hata
    Dim frm As Form, id As Long, n As Long
    Set frm = F()
    If Bos(frm!txtPersonelID) Then
        Mesaj "Silmek için önce listeden bir personel seçin.", vbExclamation
        Exit Function
    End If
    id = CLng(frm!txtPersonelID)
    n = DCount("*", "tbl_Parti", "GirisPersonelID=" & id & " OR TeslimPersonelID=" & id)
    If n > 0 Then
        Mesaj "Bu personel " & n & " partide kayıtlı, silinemez.", vbExclamation
        Exit Function
    End If
    If Not Soru(Nz(frm!txtAdSoyad, "") & " silinsin mi?") Then Exit Function
    CurrentDb.Execute "DELETE FROM tbl_Personel WHERE PersonelID=" & id, dbFailOnError
    PersonelTemizle
    frm!lstPersonel.Requery
    KombolariYenile
    Mesaj "Personel silindi."
    Exit Function
Hata:
    Mesaj "Silme sırasında hata: " & Err.Description, vbExclamation
End Function