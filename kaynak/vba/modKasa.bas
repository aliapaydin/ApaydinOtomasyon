Option Compare Database
Option Explicit

' =============================================================
' KASA / BANKA HAREKETLERİ   (form: frm_CicerKasa)
' -------------------------------------------------------------
' Üstte hesap seçilir: nakit kasa veya banka hesabı (banka hesapları
' Ayarlar ekranından eklenir). Liste ve bakiye seçili hesaba göredir.
' Tahsilat / Ödeme girişi. Müşteri seçilirse cariye de işlenir:
'   Tahsilat -> müşterinin carisine ALACAK
'   Ödeme    -> müşterinin carisine BORÇ
' Müşterinin borcu olan bir parti seçilirse (ör. sonradan gelen havale)
' tahsilat o partiye bağlanır ve partinin kalan borcundan düşer.
' Mal kabul, teslim ve alım-satımdan gelen otomatik kayıtlar buradan
' değiştirilemez; kaynağından düzeltilir.
' =============================================================

Private Const FORM_ADI As String = "frm_CicerKasa"
Private Const KUTULAR As String = "txtKasaHareketID,cmbMusteri,cmbParti,cmbHareketTuru,cmbOdemeYontemi,txtTutar,txtAciklama"

Private Function F() As Form
    Set F = Forms(FORM_ADI)
End Function

Private Function SeciliKasa() As Long
    SeciliKasa = Nz(F().Controls("cmbKasa").Value, 0)
    If SeciliKasa = 0 Then SeciliKasa = VarsayilanKasaID()
End Function

Public Function KasaFormYukle()
    On Error Resume Next
    Dim frm As Form
    Set frm = F()
    frm!cmbKasa.Requery
    frm!cmbKasa = VarsayilanKasaID()
    KasaHesapDegisti
End Function

Public Function KasaHesapDegisti()
    On Error Resume Next
    Dim frm As Form
    Set frm = F()
    frm!txtKasaAcilis = Nz(DLookup("AcilisBakiyesi", "tbl_Kasa", "KasaID=" & SeciliKasa()), 0)
    KasaTemizle
    KasaFiltrele
End Function

Private Sub KasaBakiyeGoster()
    On Error Resume Next
    F().Controls("txtKasaBakiye").Value = KasaBakiye(SeciliKasa())
End Sub

Public Function KasaTemizle()
    Dim frm As Form
    Set frm = F()
    KutulariTemizle frm, Split(KUTULAR, ",")
    frm!txtTarih = Date
    If Nz(DLookup("KasaTuru", "tbl_Kasa", "KasaID=" & SeciliKasa()), "Nakit") = "Banka" Then
        frm!cmbOdemeYontemi = "Havale"
    Else
        frm!cmbOdemeYontemi = "Nakit"
    End If
    frm!lstHareketler = Null
    KasaPartiListesi
    KasaBakiyeGoster
End Function

Public Function KasaYeni()
    KasaTemizle
    F().Controls("cmbHareketTuru").SetFocus
End Function

' Müşteri seçilince borcu kalan partileri listeler
Public Function KasaMusteriSecildi()
    Dim frm As Form
    Set frm = F()
    frm!cmbParti = Null
    KasaPartiListesi
End Function

Private Sub KasaPartiListesi()
    Dim frm As Form
    Set frm = F()
    If Bos(frm!cmbMusteri) Then
        frm!cmbParti.RowSource = "SELECT PartiNo, PartiKodu FROM tbl_Parti WHERE 1=0;"
    Else
        frm!cmbParti.RowSource = "SELECT PartiNo, [Parti No], Sezon, [Kalan Borç] FROM qry_PartiBorc WHERE MusteriID=" & CLng(frm!cmbMusteri) & _
            " AND (KalanHam>0.005" & IIf(Bos(frm!txtKasaHareketID), "", " OR PartiNo=" & Nz(DLookup("PartiNo", "tbl_KasaHareket", "KasaHareketID=" & Nz(frm!txtKasaHareketID, 0)), 0)) & _
            ") ORDER BY Sezon DESC, PartiSira;"
    End If
End Sub

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
            Mesaj "Bu hareket '" & KaynakAdi(rs!kaynakTipi) & "' ekranından otomatik oluşturuldu." & vbCrLf & _
                  "Değiştirmek veya silmek için o ekrandan düzeltin.", vbExclamation
            OtomatikKasaMi = True
        End If
    End If
    rs.Close
End Function

Private Sub KasaYazIslem(ByVal guncelle As Boolean)
    On Error GoTo Hata
    Dim frm As Form, id As Long, tahsilatMi As Boolean, tutar As Currency, islemde As Boolean
    Dim musteri As Variant, parti As Variant, acik As String, kasa As Long
    Set frm = F()

    If guncelle Then
        If Bos(frm!txtKasaHareketID) Then
            Mesaj "Güncellemek için önce listeden bir hareket seçin (çift tıklayın).", vbExclamation
            Exit Sub
        End If
        id = CLng(frm!txtKasaHareketID)
        If OtomatikKasaMi(id) Then Exit Sub
    ElseIf Not Bos(frm!txtKasaHareketID) Then
        Mesaj "Bu hareket zaten kayıtlı. Değişiklik için 'Güncelle' düğmesini kullanın." & vbCrLf & _
              "Yeni hareket için önce 'Yeni Kayıt' düğmesine basın.", vbExclamation
        Exit Sub
    End If
    kasa = SeciliKasa()
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
    parti = SayiVeyaNull(frm!cmbParti)
    If Not IsNull(parti) And IsNull(musteri) Then parti = Null
    acik = Nz(frm!txtAciklama, "")
    If acik = "" And Not IsNull(parti) Then acik = PartiAdi(CLng(parti)) & " sıkım ücreti " & LCase$(Nz(frm!cmbOdemeYontemi, "")) & " tahsilatı"
    If Not IsNull(parti) And tahsilatMi Then
        If tutar > PartiKalanBorc(CLng(parti)) + EskiPartiTutari(id, CLng(parti)) + 0.005 Then
            If Not Soru("Tutar, partinin kalan borcundan (" & TLYaz(PartiKalanBorc(CLng(parti)) + EskiPartiTutari(id, CLng(parti))) & ") fazla. Yine de kaydedilsin mi?") Then Exit Sub
        End If
    End If

    DBEngine.BeginTrans: islemde = True
    id = KasaYaz(CDate(frm!txtTarih), musteri, parti, IIf(tahsilatMi, "Tahsilat", "Ödeme"), _
        Nz(frm!cmbOdemeYontemi, ""), acik, IIf(tahsilatMi, tutar, 0), IIf(tahsilatMi, 0, tutar), _
        KAYNAK_MANUEL, Null, False, id, kasa)
    ' Bağlı cari kaydını yeniden yaz
    OtomatikKayitlariSil KAYNAK_KASA, id
    If Not IsNull(musteri) Then
        CariYaz CDate(frm!txtTarih), musteri, parti, IIf(tahsilatMi, "Tahsilat", "Ödeme"), _
            IIf(acik = "", IIf(tahsilatMi, "Tahsilat", "Ödeme") & " (" & Nz(frm!cmbOdemeYontemi, "") & ")", acik), _
            IIf(tahsilatMi, 0, tutar), IIf(tahsilatMi, tutar, 0), KAYNAK_KASA, id, True
    End If
    DBEngine.CommitTrans: islemde = False

    frm!txtKasaHareketID = id
    frm!lstHareketler.Requery
    KasaBakiyeGoster
    Mesaj IIf(guncelle, "Hareket güncellendi.", "Hareket kaydedildi.") & _
          IIf(IsNull(musteri), "", vbCrLf & "Müşterinin carisine de işlendi.") & _
          IIf(IsNull(parti), "", vbCrLf & PartiAdi(CLng(parti)) & " kalan borç: " & TLYaz(PartiKalanBorc(CLng(parti))))
    Exit Sub
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Sub

' Güncellenen hareketin daha önce aynı partiye yazdığı tutar
Private Function EskiPartiTutari(ByVal id As Long, ByVal parti As Long) As Currency
    If id = 0 Then Exit Function
    EskiPartiTutari = Nz(DLookup("Tahsilat", "tbl_KasaHareket", "KasaHareketID=" & id & " AND PartiNo=" & parti), 0)
End Function

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
    KasaPartiListesi
    frm!cmbParti = rs!partiNo
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
        Mesaj "Silmek için önce listeden bir hareket seçin.", vbExclamation
        Exit Function
    End If
    id = CLng(frm!txtKasaHareketID)
    If OtomatikKasaMi(id) Then Exit Function
    If Not Soru("Seçili hareket (ve bağlı cari kaydı) silinsin mi?") Then Exit Function
    DBEngine.BeginTrans: islemde = True
    OtomatikKayitlariSil KAYNAK_KASA, id
    CurrentDb.Execute "DELETE FROM tbl_KasaHareket WHERE KasaHareketID=" & id, dbFailOnError
    DBEngine.CommitTrans: islemde = False
    KasaTemizle
    frm!lstHareketler.Requery
    KasaBakiyeGoster
    Mesaj "Hareket silindi."
    Exit Function
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "Silme sırasında hata: " & Err.Description, vbExclamation
End Function

' Seçili hesabın hareketleri; iki tarih arası süzülebilir
Public Function KasaFiltrele()
    On Error GoTo Hata
    Dim frm As Form, filtre As String
    Set frm = F()
    filtre = " WHERE KasaID=" & SeciliKasa()
    If TarihMi(frm!txtBaslangic) Then filtre = filtre & " AND TarihHam>=" & SqlTarih(CDate(frm!txtBaslangic))
    If TarihMi(frm!txtBitis) Then filtre = filtre & " AND TarihHam<=" & SqlTarih(CDate(frm!txtBitis))
    frm!lstHareketler.RowSource = "SELECT * FROM qry_KasaEkstre" & filtre & " ORDER BY TarihHam, KasaHareketID;"
    KasaBakiyeGoster
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
    CurrentDb.Execute "UPDATE tbl_Kasa SET AcilisBakiyesi=" & SqlSayi(CDbl(frm!txtKasaAcilis)) & _
        " WHERE KasaID=" & SeciliKasa(), dbFailOnError
    frm!lstHareketler.Requery
    KasaBakiyeGoster
    Mesaj "Açılış bakiyesi kaydedildi."
    Exit Function
Hata:
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Function
