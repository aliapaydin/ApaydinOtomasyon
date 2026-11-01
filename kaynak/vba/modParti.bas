Option Compare Database
Option Explicit

' =============================================================
' PARTİ / ÜRETİM YÖNETİMİ   (form: frm_CicerParti)
' -------------------------------------------------------------
' Hesaplama kuralları:
'   Kızıl su / Posa (Atık) = Zeytin Kg - Çıkan Yağ Kg
'   Sıkım Ücreti seçilirse : TL Tutarı  = Zeytin Kg x Birim Sıkım Ücreti
'                            Müşteriye Kalan = Çıkan Yağ
'   Hakyağı seçilirse      : Hakyağı Kg = Çıkan Yağ Kg x Hakyağı Oranı / 100
'                            Müşteriye Kalan = Çıkan Yağ - Hakyağı
'
' Kaydet / Güncelle sonrası otomatik yazılan kayıtlar:
'   Sıkım Ücreti -> Cari BORÇ (TL Tutarı)
'                   "Peşin Tahsil Edildi" işaretliyse ayrıca
'                   Cari ALACAK + Kasa TAHSİLAT
'   Hakyağı      -> Stok GİRİŞ (Hak Yağı ürünü, Hakyağı Kg)
'
' Bir partiyi güncellemek (ör. müşteri hakyağından vazgeçip ücretliye
' geçti) o partinin eski otomatik kayıtlarını silip yenisini yazar.
' Bidonlar kutusuna "151, 152, 76" gibi yazılan numaralar
' tbl_PartiBidon tablosuna da işlenir (Müşteri/Cari ekranında görünür).
' =============================================================

Private Const FORM_ADI As String = "frm_CicerParti"
Private Const KUTULAR As String = "txtPartiNo,txtPartiKodu,cmbMusteri,txtTCKN,cmbCinsi,cmbOzellik,cmbGirisPersonel," & _
    "txtZeytinMiktari,txtCikanYag,txtAtikMiktari,cmbOdemeTipi,txtSikimUcreti,txtHakYagiOrani," & _
    "txtTLTutari,txtHakYagi,txtMusteriKalanYag,txtSikimTarihi,txtTeslimTarihi,cmbTeslimPersonel,txtBidonDetay"

Private Function F() As Form
    Set F = Forms(FORM_ADI)
End Function

' ---------------------------------------------------------------
' Hesapla (miktar, ücret, oran, ödeme tipi değişince de çalışır)
' ---------------------------------------------------------------
Public Function PartiHesapla()
    On Error GoTo Hata
    Dim frm As Form, zeytin As Double, yag As Double, hak As Double
    Set frm = F()
    zeytin = Sayi(frm!txtZeytinMiktari)
    yag = Sayi(frm!txtCikanYag)

    If zeytin > 0 And yag > 0 Then
        frm!txtAtikMiktari = zeytin - yag
    Else
        frm!txtAtikMiktari = Null
    End If

    If UcretliMi(frm!cmbOdemeTipi) Then
        frm!txtTLTutari = Round(zeytin * Sayi(frm!txtSikimUcreti), 2)
        frm!txtHakyagi = 0
        If yag > 0 Then frm!txtMusteriKalanYag = yag Else frm!txtMusteriKalanYag = Null
    ElseIf HakyagiMi(frm!cmbOdemeTipi) Then
        hak = Round(yag * Sayi(frm!txtHakYagiOrani) / 100, 2)
        frm!txtHakyagi = hak
        frm!txtTLTutari = 0
        If yag > 0 Then frm!txtMusteriKalanYag = yag - hak Else frm!txtMusteriKalanYag = Null
    End If
    PartiMusteriSecildi
    PartiAlanDurumu
    Exit Function
Hata:
    Mesaj "Hesaplama hatası: " & Err.Description, vbExclamation
End Function

' Ödeme tipine göre ilgili kutuları açar / kapatır
Private Sub PartiAlanDurumu()
    On Error Resume Next
    Dim frm As Form, ucretli As Boolean, hakyagi As Boolean
    Set frm = F()
    ucretli = UcretliMi(frm!cmbOdemeTipi)
    hakyagi = HakyagiMi(frm!cmbOdemeTipi)
    frm!txtSikimUcreti.Enabled = (ucretli Or Not hakyagi)
    frm!txtHakYagiOrani.Enabled = (hakyagi Or Not ucretli)
    frm!chkPesin.Enabled = ucretli
    If Not ucretli Then frm!chkPesin = False
End Sub

' Müşteri seçilince TCKN'yi gösterir
Public Function PartiMusteriSecildi()
    Dim frm As Form
    Set frm = F()
    If Bos(frm!cmbMusteri) Then
        frm!txtTCKN = Null
    Else
        frm!txtTCKN = DLookup("TCKN_VKN", "tbl_Musteri", "MusteriID=" & frm!cmbMusteri)
    End If
End Function

' ---------------------------------------------------------------
' Yeni / Temizle
' ---------------------------------------------------------------
Public Function PartiTemizle()
    Dim frm As Form
    Set frm = F()
    KutulariTemizle frm, Split(KUTULAR, ",")
    frm!chkPesin = False
    frm!txtTarih = Date
    frm!lstPartiler = Null
    PartiAlanDurumu
End Function

Public Function PartiYeni()
    Dim frm As Form
    PartiTemizle
    Set frm = F()
    frm!txtPartiKodu = Nz(DMax("PartiKodu", "tbl_Parti"), 0) + 1   ' öneri; değiştirilebilir
    ' Ayarlar ekranındaki varsayılanlar
    If SayiMi(Ayar("VarsayilanSikimUcreti")) Then frm!txtSikimUcreti = CDbl(Ayar("VarsayilanSikimUcreti"))
    If SayiMi(Ayar("VarsayilanHakyagiOrani")) Then frm!txtHakYagiOrani = CDbl(Ayar("VarsayilanHakyagiOrani"))
    frm!txtPartiKodu.SetFocus
End Function

' ---------------------------------------------------------------
' Kaydet / Güncelle
' ---------------------------------------------------------------
Public Function PartiKaydet()
    PartiYazIslem False
End Function

Public Function PartiGuncelle()
    PartiYazIslem True
End Function

Private Function PartiKontrol(ByVal partiNo As Long) As Boolean
    Dim frm As Form
    Set frm = F()
    If Not SayiMi(frm!txtPartiKodu) Then
        Mesaj "Parti No giriniz (sayı olmalı).", vbExclamation: Exit Function
    End If
    If DCount("*", "tbl_Parti", "PartiKodu=" & CLng(frm!txtPartiKodu) & " AND PartiNo<>" & partiNo) > 0 Then
        Mesaj "Bu Parti No (" & frm!txtPartiKodu & ") başka bir partide kullanılıyor.", vbExclamation: Exit Function
    End If
    If Not TarihMi(frm!txtTarih) Then
        Mesaj "Tarih giriniz.", vbExclamation: Exit Function
    End If
    If Bos(frm!cmbMusteri) Then
        Mesaj "Müşteri seçiniz.", vbExclamation: Exit Function
    End If
    If Sayi(frm!txtZeytinMiktari) <= 0 Then
        Mesaj "Zeytin miktarını (kg) giriniz.", vbExclamation: Exit Function
    End If
    If Sayi(frm!txtCikanYag) > Sayi(frm!txtZeytinMiktari) Then
        Mesaj "Çıkan yağ, zeytin miktarından büyük olamaz.", vbExclamation: Exit Function
    End If
    If UcretliMi(frm!cmbOdemeTipi) Then
        If Sayi(frm!txtSikimUcreti) <= 0 Then
            Mesaj "Birim sıkım ücretini giriniz.", vbExclamation: Exit Function
        End If
    ElseIf HakyagiMi(frm!cmbOdemeTipi) Then
        If Sayi(frm!txtHakYagiOrani) <= 0 Then
            Mesaj "Hakyağı oranını (%) giriniz.", vbExclamation: Exit Function
        End If
        If Sayi(frm!txtCikanYag) <= 0 Then
            Mesaj "Hakyağı hesabı için çıkan yağ miktarını giriniz.", vbExclamation: Exit Function
        End If
    Else
        Mesaj "Ödeme tipini seçiniz (Sıkım Ücreti veya Hakyağı).", vbExclamation: Exit Function
    End If
    If Not Bos(frm!txtTeslimTarihi) And Not TarihMi(frm!txtTeslimTarihi) Then
        Mesaj "Teslim tarihi geçerli değil.", vbExclamation: Exit Function
    End If
    PartiKontrol = True
End Function

Private Sub PartiYazIslem(ByVal guncelle As Boolean)
    On Error GoTo Hata
    Dim frm As Form, partiNo As Long, islemde As Boolean
    Set frm = F()

    If guncelle Then
        If Bos(frm!txtPartiNo) Then
            Mesaj "Güncellemek için önce listeden bir parti seçin (çift tıklayın).", vbExclamation
            Exit Sub
        End If
        partiNo = CLng(frm!txtPartiNo)
    ElseIf Not Bos(frm!txtPartiNo) Then
        Mesaj "Bu parti zaten kayıtlı. Değişiklik için 'Güncelle' düğmesini kullanın." & vbCrLf & _
              "Yeni parti için önce 'Yeni Kayıt' düğmesine basın.", vbExclamation
        Exit Sub
    End If

    PartiHesapla
    If Not PartiKontrol(partiNo) Then Exit Sub

    DBEngine.BeginTrans: islemde = True
    partiNo = PartiTabloyaYaz(partiNo)
    PartiMuhasebeYaz partiNo
    BidonlariEsle partiNo
    DBEngine.CommitTrans: islemde = False

    frm!txtPartiNo = partiNo
    frm!lstPartiler.Requery
    Mesaj IIf(guncelle, "Parti güncellendi.", "Parti kaydedildi.") & vbCrLf & _
          "Cari / kasa / stok kayıtları otomatik oluşturuldu."
    Exit Sub
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Sub

Private Function PartiTabloyaYaz(ByVal partiNo As Long) As Long
    Dim frm As Form, rs As DAO.Recordset
    Set frm = F()
    Set rs = KayitAc("tbl_Parti", "PartiNo", partiNo)
    rs!PartiKodu = CLng(frm!txtPartiKodu)
    rs!tarih = CDate(frm!txtTarih)
    rs!musteriID = CLng(frm!cmbMusteri)
    rs!Cinsi = MetinVeyaNull(frm!cmbCinsi)
    rs!Ozellik = MetinVeyaNull(frm!cmbOzellik)
    rs!GirisPersonelID = SayiVeyaNull(frm!cmbGirisPersonel)
    rs!ZeytinMiktari_Kg = SayiVeyaNull(frm!txtZeytinMiktari)
    rs!CikanYag_Kg = SayiVeyaNull(frm!txtCikanYag)
    rs!AtikMiktari_Kg = SayiVeyaNull(frm!txtAtikMiktari)
    rs!odemeTipi = MetinVeyaNull(frm!cmbOdemeTipi)
    rs!SimSikimUcreti = SayiVeyaNull(frm!txtSikimUcreti)
    rs!HakYagiOrani_Yuzde = SayiVeyaNull(frm!txtHakYagiOrani)
    rs!TLTutari = SayiVeyaNull(frm!txtTLTutari)
    rs!HakYagi_Kg = SayiVeyaNull(frm!txtHakyagi)
    rs!MusteriKalanYag_Kg = SayiVeyaNull(frm!txtMusteriKalanYag)
    rs!SikimTarihi = TarihVeyaNull(frm!txtSikimTarihi)
    rs!TeslimTarihi = TarihVeyaNull(frm!txtTeslimTarihi)
    rs!TeslimPersonelID = SayiVeyaNull(frm!cmbTeslimPersonel)
    rs!BidonDetay = MetinVeyaNull(frm!txtBidonDetay)
    rs!PesinTahsilEdildi = Nz(frm!chkPesin, False)
    If Not Nz(frm!chkPesin, False) Then rs!TahsilatTarihi = Null
    rs.Update
    rs.Bookmark = rs.LastModified
    PartiTabloyaYaz = rs!partiNo
    rs.Close
End Function

' ---------------------------------------------------------------
' Partinin cari / kasa / stok kayıtlarını (yeniden) yazar.
' Bilgiler tablodan okunur; form açık olmasa da çalışır.
' ---------------------------------------------------------------
Public Sub PartiMuhasebeYaz(ByVal partiNo As Long)
    Dim rs As DAO.Recordset, acik As String, tutar As Currency, tTarih As Date
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Parti WHERE PartiNo=" & partiNo, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Sub

    OtomatikKayitlariSil KAYNAK_PARTI, partiNo
    If IsNull(rs!musteriID) Or IsNull(rs!tarih) Then rs.Close: Exit Sub
    acik = "Parti " & rs!PartiKodu

    If UcretliMi(rs!odemeTipi) Then
        tutar = Nz(rs!TLTutari, 0)
        If tutar > 0 Then
            CariYaz rs!tarih, rs!musteriID, partiNo, "Sıkım Ücreti", _
                acik & " sıkım ücreti (" & Nz(rs!ZeytinMiktari_Kg, 0) & " kg x " & Nz(rs!SimSikimUcreti, 0) & " TL)", _
                tutar, 0, KAYNAK_PARTI, partiNo, True
            If Nz(rs!PesinTahsilEdildi, False) Then
                ' Teslimde tahsil edildiyse tahsilat tarihi teslim tarihidir
                tTarih = Nz(rs!TahsilatTarihi, rs!tarih)
                CariYaz tTarih, rs!musteriID, partiNo, "Peşin Tahsilat", acik & " sıkım ücreti tahsilatı", _
                    0, tutar, KAYNAK_PARTI, partiNo, True
                KasaYaz tTarih, rs!musteriID, partiNo, "Tahsilat", "Nakit", acik & " sıkım ücreti tahsilatı", _
                    tutar, 0, KAYNAK_PARTI, partiNo, True
            End If
        End If
    ElseIf HakyagiMi(rs!odemeTipi) Then
        If Nz(rs!HakYagi_Kg, 0) > 0 Then
            StokYaz rs!tarih, HakYagiUrunID(), rs!musteriID, partiNo, "Hakyağı Giriş", _
                acik & " %" & Nz(rs!HakYagiOrani_Yuzde, 0) & " hakyağı (" & Nz(rs!CikanYag_Kg, 0) & " kg yağdan)", _
                rs!HakYagi_Kg, 0, KAYNAK_PARTI, partiNo, True
        End If
    End If
    rs.Close
End Sub

' ---------------------------------------------------------------
' Bidonlar kutusundaki numaraları tbl_PartiBidon'a işler.
' Teslim tarihi girilmişse bidonlar "teslim edildi" olarak işaretlenir.
' ---------------------------------------------------------------
Public Sub BidonlariEsle(ByVal partiNo As Long)
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim metin As String, parcalar() As String, yeni As String, eski As String, b As String, i As Long
    Dim TESLIM As Variant

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT BidonDetay, TeslimTarihi FROM tbl_Parti WHERE PartiNo=" & partiNo, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Sub
    metin = Nz(rs!BidonDetay, "")
    TESLIM = rs!TeslimTarihi
    rs.Close

    metin = Replace(Replace(Replace(metin, ";", ","), "/", ","), " ", ",")
    parcalar = Split(metin, ",")
    For i = LBound(parcalar) To UBound(parcalar)
        b = Trim$(parcalar(i))
        If b <> "" Then yeni = yeni & "|" & Left$(b, 30)
    Next

    Set rs = db.OpenRecordset("SELECT BidonNo FROM tbl_PartiBidon WHERE PartiNo=" & partiNo & " ORDER BY PartiBidonID", dbOpenSnapshot)
    Do Until rs.EOF
        eski = eski & "|" & rs!BidonNo
        rs.MoveNext
    Loop
    rs.Close

    If eski <> yeni Then
        db.Execute "DELETE FROM tbl_PartiBidon WHERE PartiNo=" & partiNo, dbFailOnError
        If yeni <> "" Then
            parcalar = Split(Mid$(yeni, 2), "|")
            Set rs = db.OpenRecordset("tbl_PartiBidon", dbOpenDynaset)
            For i = 0 To UBound(parcalar)
                rs.AddNew
                rs!partiNo = partiNo
                rs!BidonNo = parcalar(i)
                rs.Update
            Next
            rs.Close
        End If
    End If

    If IsNull(TESLIM) Then
        db.Execute "UPDATE tbl_PartiBidon SET TeslimEdildi=False, TeslimTarihi=Null WHERE PartiNo=" & partiNo, dbFailOnError
    Else
        db.Execute "UPDATE tbl_PartiBidon SET TeslimEdildi=True, TeslimTarihi=" & SqlTarih(TESLIM) & _
            " WHERE PartiNo=" & partiNo, dbFailOnError
    End If
End Sub

' ---------------------------------------------------------------
' Listeden seç (çift tıklama)
' ---------------------------------------------------------------
Public Function PartiListedenSec()
    On Error GoTo Hata
    Dim frm As Form, rs As DAO.Recordset, id As Long
    Set frm = F()
    id = ListeID(frm!lstPartiler)
    If id = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Parti WHERE PartiNo=" & id, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function
    frm!txtPartiNo = rs!partiNo
    frm!txtPartiKodu = rs!PartiKodu
    frm!txtTarih = rs!tarih
    frm!cmbMusteri = rs!musteriID
    frm!cmbCinsi = rs!Cinsi
    frm!cmbOzellik = rs!Ozellik
    frm!cmbGirisPersonel = rs!GirisPersonelID
    frm!txtZeytinMiktari = rs!ZeytinMiktari_Kg
    frm!txtCikanYag = rs!CikanYag_Kg
    frm!txtAtikMiktari = rs!AtikMiktari_Kg
    frm!cmbOdemeTipi = rs!odemeTipi
    frm!txtSikimUcreti = rs!SimSikimUcreti
    frm!txtHakYagiOrani = rs!HakYagiOrani_Yuzde
    frm!txtTLTutari = rs!TLTutari
    frm!txtHakyagi = rs!HakYagi_Kg
    frm!txtMusteriKalanYag = rs!MusteriKalanYag_Kg
    frm!txtSikimTarihi = rs!SikimTarihi
    frm!txtTeslimTarihi = rs!TeslimTarihi
    frm!cmbTeslimPersonel = rs!TeslimPersonelID
    frm!txtBidonDetay = rs!BidonDetay
    frm!chkPesin = Nz(rs!PesinTahsilEdildi, False)
    rs.Close
    PartiMusteriSecildi
    PartiAlanDurumu
    Exit Function
Hata:
    Mesaj "Parti açılamadı: " & Err.Description, vbExclamation
End Function

' ---------------------------------------------------------------
' Sil
' ---------------------------------------------------------------
Public Function PartiSil()
    On Error GoTo Hata
    Dim frm As Form, id As Long, n As Long, islemde As Boolean, elle As String
    Set frm = F()
    If Bos(frm!txtPartiNo) Then
        Mesaj "Silmek için önce listeden bir parti seçin (çift tıklayın).", vbExclamation
        Exit Function
    End If
    id = CLng(frm!txtPartiNo)

    elle = "PartiNo=" & id & " AND (OtomatikMi=False OR OtomatikMi Is Null)"
    n = DCount("*", "tbl_CariHareket", elle) + DCount("*", "tbl_KasaHareket", elle) + DCount("*", "tbl_StokHareket", elle)
    If n > 0 Then
        Mesaj "Bu partiye bağlı elle girilmiş " & n & " cari/kasa/stok hareketi var." & vbCrLf & _
              "Önce o hareketleri silin.", vbExclamation
        Exit Function
    End If
    If Not Soru("Parti " & frm!txtPartiKodu & " ve otomatik oluşan cari / kasa / stok kayıtları silinsin mi?") Then Exit Function

    DBEngine.BeginTrans: islemde = True
    OtomatikKayitlariSil KAYNAK_PARTI, id
    CurrentDb.Execute "DELETE FROM tbl_PartiBidon WHERE PartiNo=" & id, dbFailOnError
    CurrentDb.Execute "DELETE FROM tbl_Parti WHERE PartiNo=" & id, dbFailOnError
    DBEngine.CommitTrans: islemde = False

    PartiTemizle
    frm!lstPartiler.Requery
    Mesaj "Parti silindi."
    Exit Function
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "Silme sırasında hata: " & Err.Description, vbExclamation
End Function

' ---------------------------------------------------------------
' BAKIM: Tüm partileri yeniden hesaplar ve otomatik kayıtlarını
' yeniden yazar. (Hesap kuralı değişirse veya veriler elle
' düzeltilirse VBA penceresinden çalıştırılabilir.)
' ---------------------------------------------------------------
Public Sub TumPartileriYenidenHesapla()
    Dim rs As DAO.Recordset, zeytin As Double, yag As Double, hak As Double
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Parti", dbOpenDynaset)
    Do Until rs.EOF
        zeytin = Nz(rs!ZeytinMiktari_Kg, 0)
        yag = Nz(rs!CikanYag_Kg, 0)
        rs.Edit
        If UcretliMi(rs!odemeTipi) Then rs!odemeTipi = "Sıkım Ücreti"
        If HakyagiMi(rs!odemeTipi) Then rs!odemeTipi = "Hakyağı"
        If InStr(Sade(rs!Cinsi), "yaglik") > 0 Then rs!Cinsi = "Yağlık Zeytin"
        If InStr(Sade(rs!Cinsi), "sofralik") > 0 Then rs!Cinsi = "Sofralık Zeytin"
        If zeytin > 0 And yag > 0 Then rs!AtikMiktari_Kg = zeytin - yag
        If UcretliMi(rs!odemeTipi) Then
            rs!TLTutari = Round(zeytin * Nz(rs!SimSikimUcreti, 0), 2)
            rs!HakYagi_Kg = 0
            If yag > 0 Then rs!MusteriKalanYag_Kg = yag
        ElseIf HakyagiMi(rs!odemeTipi) Then
            hak = Round(yag * Nz(rs!HakYagiOrani_Yuzde, 0) / 100, 2)
            rs!HakYagi_Kg = hak
            rs!TLTutari = 0
            rs!PesinTahsilEdildi = False
            If yag > 0 Then rs!MusteriKalanYag_Kg = yag - hak
        End If
        rs.Update
        PartiMuhasebeYaz rs!partiNo
        BidonlariEsle rs!partiNo
        rs.MoveNext
    Loop
    rs.Close
End Sub