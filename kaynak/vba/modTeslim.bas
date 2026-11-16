Option Compare Database
Option Explicit

' =============================================================
' YAĞ TESLİM   (form: frm_CicerTeslim)
' -------------------------------------------------------------
' Üretimi yapılmış partinin yağı müşteriye teslim edilir:
'   teslim tarihi, teslim alan, teslim eden (personel),
'   teslim edilen yağ ve açıklama kaydedilir.
' Sıkım ücretli partide tahsilat seçilir:
'   Tahsil Edilmedi : ücretin tamamı müşterinin borcunda kalır
'   Peşin           : kalan borcun tamamı tahsil edilir
'   Kısmi           : girilen tutar tahsil edilir, kalanı borçta kalır
'   Yöntem: Nakit (kasaya), Havale / Kart (seçilen banka hesabına)
' Tahsilat cariye ALACAK, kasaya / bankaya TAHSİLAT olarak yazılır
' ve partiye bağlanır. Sonradan gelen havaleler Kasa / Banka
' ekranından partiye bağlanarak girilir ve borçtan düşer.
' Teslim geri alınırken parti numarası aynı sezonda başka bir açık
' partide kullanılıyorsa eski parti 105/1, 105/2 ... olarak yeniden adlanır.
' Teslim fişi A5 olarak önizlenir, PDF kaydedilir veya yazdırılır.
' =============================================================

Private Const FORM_ADI As String = "frm_CicerTeslim"
Private Const BILGI As String = "txtMusteri,txtTelefon,txtPartiTarihi,txtCinsi,txtOdemeTipi,txtZeytin,txtYag,txtRandiman," & _
    "txtHakyagi,txtNetYag,txtBidonlar,txtUcret,txtOdenen,txtKalanBorc,txtCariBakiye"
Private Const TESLIM As String = "txtTeslimTarihi,txtTeslimAlan,cmbTeslimEden,txtTeslimYag,txtTeslimAciklama," & _
    "cmbTahsilat,txtTahsilatTutar,cmbYontem,cmbHesap"

Private OncekiTeslimTahsilati As Currency   ' seçili partinin kayıtlı teslim tahsilatı

Private Function F() As Form
    Set F = Forms(FORM_ADI)
End Function

Public Function TeslimFormYukle()
    On Error Resume Next
    F().Controls("cmbDurumFiltre").Value = "Teslim Bekleyenler"
    TeslimTemizle
End Function

Public Function TeslimTemizle()
    Dim frm As Form
    Set frm = F()
    KutulariTemizle frm, Split(BILGI & ",cmbParti," & TESLIM, ",")
    frm!lstTeslim = Null
    frm!lblDurum.Caption = ""
    frm!lblSezon.Caption = AktifSezonMetni()
    OncekiTeslimTahsilati = 0
    frm!cmbParti.RowSource = "SELECT PartiNo, [Parti No], Müşteri, Tarih, Durum FROM qry_MalKabulListe " & _
        "WHERE Sezon=" & SqlMetin(AktifSezon()) & " AND Durum<>'" & DURUM_KABUL & "' ORDER BY IIf(Durum='" & DURUM_TESLIM & "',1,0), PartiSira;"
    TahsilatAlanDurumu
    TeslimListeFiltrele
End Function

Public Function TeslimListeFiltrele()
    On Error Resume Next
    Dim frm As Form, w As String
    Set frm = F()
    w = " WHERE Sezon=" & SqlMetin(AktifSezon())
    Select Case Nz(frm!cmbDurumFiltre, "")
        Case "Teslim Bekleyenler": w = w & " AND TeslimTarihiHam Is Null"
        Case "Teslim Edilenler": w = w & " AND TeslimTarihiHam Is Not Null"
        Case "Borcu Kalanlar": w = w & " AND KalanHam>0.005"
    End Select
    frm!lstTeslim.RowSource = "SELECT * FROM qry_TeslimListe" & w & " ORDER BY PartiTarihiHam DESC, PartiSira DESC;"
End Function

Public Function TeslimListedenSec()
    Dim id As Long
    id = ListeID(F().Controls("lstTeslim"))
    If id = 0 Then Exit Function
    F().Controls("cmbParti").Value = id
    TeslimPartiSecildi
End Function

' Parti seçilince özet ve teslim bilgilerini doldurur
Public Function TeslimPartiSecildi()
    On Error GoTo Hata
    Dim frm As Form, rs As DAO.Recordset, zeytin As Double, yag As Double, net As Double, ucret As Currency
    Set frm = F()
    KutulariTemizle frm, Split(BILGI & "," & TESLIM, ",")
    frm!lblDurum.Caption = ""
    OncekiTeslimTahsilati = 0
    If Bos(frm!cmbParti) Then TahsilatAlanDurumu: Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT p.*, m.AdSoyad, m.Telefon FROM tbl_Parti AS p LEFT JOIN tbl_Musteri AS m ON p.MusteriID=m.MusteriID WHERE p.PartiNo=" & CLng(frm!cmbParti), dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function

    zeytin = Nz(rs!ZeytinMiktari_Kg, 0)
    yag = Nz(rs!CikanYag_Kg, 0)
    frm!txtMusteri = rs!adSoyad
    frm!txtTelefon = rs!Telefon
    frm!txtPartiTarihi = Format$(rs!tarih, "dd.mm.yyyy")
    frm!txtCinsi = Nz(rs!Cinsi, "") & IIf(IsNull(rs!Ozellik), "", " / " & rs!Ozellik)
    frm!txtOdemeTipi = rs!odemeTipi
    frm!txtZeytin = KgMetin(zeytin)
    frm!txtYag = KgMetin(yag)
    If zeytin > 0 And yag > 0 Then frm!txtRandiman = "%" & Format$(yag / zeytin * 100, "0.00")
    frm!txtBidonlar = rs!BidonDetay

    If HakyagiMi(rs!odemeTipi) Then
        net = Nz(rs!MusteriKalanYag_Kg, yag - Nz(rs!HakYagi_Kg, 0))
        frm!txtHakyagi = "%" & Nz(rs!HakYagiOrani_Yuzde, 0) & "  =  " & KgMetin(Nz(rs!HakYagi_Kg, 0))
        frm!txtUcret = "-"
    Else
        net = yag
        frm!txtHakyagi = "-"
        ucret = Nz(rs!TLTutari, 0)
        frm!txtUcret = TLYaz(ucret)
    End If
    frm!txtNetYag = KgMetin(net)

    ' Tahsilat bilgileri
    If Nz(rs!TahsilatTuru, "Yok") <> "Yok" Then OncekiTeslimTahsilati = Nz(rs!TahsilatTutari, 0)
    TeslimBorcGoster CLng(frm!cmbParti), rs!musteriID

    ' Teslim bilgileri: kayıtlıysa onları, değilse varsayılanları getir
    If Not IsNull(rs!TeslimTarihi) Then
        frm!txtTeslimTarihi = rs!TeslimTarihi
        frm!txtTeslimAlan = rs!TeslimAlan
        frm!cmbTeslimEden = rs!TeslimPersonelID
        frm!txtTeslimYag = Nz(rs!TeslimYag_Kg, net)
        frm!txtTeslimAciklama = rs!TeslimAciklama
        frm!lblDurum.Caption = "TESLİM EDİLDİ"
        frm!cmbTahsilat = TahsilatMetni(rs!TahsilatTuru)
        If OncekiTeslimTahsilati > 0 Then
            frm!txtTahsilatTutar = OncekiTeslimTahsilati
            frm!cmbYontem = Nz(rs!TahsilatYontemi, "Nakit")
            HesapListesi
            frm!cmbHesap = rs!TahsilatKasaID
        End If
    Else
        frm!txtTeslimTarihi = Date
        frm!txtTeslimAlan = rs!adSoyad
        frm!txtTeslimYag = net
        frm!lblDurum.Caption = IIf(IsNull(rs!SikimTarihi), "ÜRETİM BEKLİYOR", "TESLİM BEKLİYOR")
        ' Teslim geri alınırken korunan tahsilat varsa onu göster
        If OncekiTeslimTahsilati > 0 Then
            frm!cmbTahsilat = TahsilatMetni(rs!TahsilatTuru)
            frm!txtTahsilatTutar = OncekiTeslimTahsilati
            frm!cmbYontem = Nz(rs!TahsilatYontemi, "Nakit")
            HesapListesi
            frm!cmbHesap = rs!TahsilatKasaID
        Else
            frm!cmbTahsilat = "Tahsil Edilmedi"
        End If
    End If
    rs.Close
    TahsilatAlanDurumu
    Exit Function
Hata:
    Mesaj "Parti bilgileri okunamadı: " & Err.Description, vbExclamation
End Function

Private Function TahsilatMetni(ByVal tur As Variant) As String
    Select Case Nz(tur, "Yok")
        Case "Peşin": TahsilatMetni = "Peşin"
        Case "Kısmi": TahsilatMetni = "Kısmi"
        Case Else: TahsilatMetni = "Tahsil Edilmedi"
    End Select
End Function

Private Sub TeslimBorcGoster(ByVal partiNo As Long, ByVal musteriID As Variant)
    Dim frm As Form, kalan As Currency
    Set frm = F()
    kalan = PartiKalanBorc(partiNo)
    If UcretliMi(frm!txtOdemeTipi) Then
        frm!txtOdenen = TLYaz(Nz(DLookup("TLTutari", "tbl_Parti", "PartiNo=" & partiNo), 0) - kalan)
        frm!txtKalanBorc = TLYaz(kalan)
    Else
        frm!txtOdenen = "-"
        frm!txtKalanBorc = "-"
    End If
    If Not IsNull(musteriID) Then frm!txtCariBakiye = TLYaz(MusteriBakiye(CLng(musteriID)))
End Sub

' Tahsilat kutularını ödeme tipine ve seçime göre açar / kapatır
Public Function TahsilatAlanDurumu()
    On Error Resume Next
    Dim frm As Form, ucretli As Boolean, tur As String
    Set frm = F()
    ucretli = UcretliMi(frm!txtOdemeTipi)
    frm!cmbTahsilat.Enabled = ucretli
    If Not ucretli Then frm!cmbTahsilat = Null
    tur = Nz(frm!cmbTahsilat, "")
    frm!txtTahsilatTutar.Enabled = (tur = "Kısmi")
    frm!cmbYontem.Enabled = (tur = "Peşin" Or tur = "Kısmi")
    frm!cmbHesap.Enabled = frm!cmbYontem.Enabled
    If tur = "Peşin" Then frm!txtTahsilatTutar = PesinTutar()
    If tur = "" Or tur = "Tahsil Edilmedi" Then
        frm!txtTahsilatTutar = Null
        frm!cmbYontem = Null
        frm!cmbHesap = Null
    ElseIf Bos(frm!cmbYontem) Then
        frm!cmbYontem = "Nakit"
        TahsilatYontemDegisti
    End If
End Function

' Peşin tahsilatta alınacak tutar: partinin kalan borcu (bu teslimin eski tahsilatı hariç)
Private Function PesinTutar() As Currency
    Dim frm As Form
    Set frm = F()
    If Bos(frm!cmbParti) Then Exit Function
    PesinTutar = PartiKalanBorc(CLng(frm!cmbParti)) + OncekiTeslimTahsilati
    If PesinTutar < 0 Then PesinTutar = 0
End Function

Public Function TahsilatYontemDegisti()
    Dim frm As Form
    Set frm = F()
    HesapListesi
    If Nz(frm!cmbYontem, "") = "Nakit" Then
        frm!cmbHesap = VarsayilanKasaID()
    ElseIf frm!cmbHesap.ListCount = 1 Then
        frm!cmbHesap = frm!cmbHesap.ItemData(0)
    Else
        frm!cmbHesap = Null
    End If
End Function

' Yönteme göre kasa (nakit) veya banka hesapları
Private Sub HesapListesi()
    Dim frm As Form, tur As String
    Set frm = F()
    tur = IIf(Nz(frm!cmbYontem, "Nakit") = "Nakit", "Nakit", "Banka")
    frm!cmbHesap.RowSource = "SELECT KasaID, KasaAdi FROM tbl_Kasa WHERE Aktif=True AND Nz(KasaTuru,'Nakit')='" & tur & "' ORDER BY KasaAdi;"
End Sub

Public Function TeslimKaydet()
    On Error GoTo Hata
    Dim frm As Form, id As Long, rs As DAO.Recordset, islemde As Boolean
    Dim tur As String, tutar As Currency, borcOnce As Currency, net As Double, ucretli As Boolean
    Set frm = F()
    If Bos(frm!cmbParti) Then
        Mesaj "Önce parti seçin.", vbExclamation: Exit Function
    End If
    id = CLng(frm!cmbParti)
    If IsNull(DLookup("SikimTarihi", "tbl_Parti", "PartiNo=" & id)) Then
        Mesaj "Bu partinin üretimi yapılmamış. Önce Üretim ekranından çıkan yağı girin.", vbExclamation: Exit Function
    End If
    If Not TarihMi(frm!txtTeslimTarihi) Then
        Mesaj "Teslim tarihini giriniz.", vbExclamation: Exit Function
    End If
    If CDate(frm!txtTeslimTarihi) < CDate(DLookup("SikimTarihi", "tbl_Parti", "PartiNo=" & id)) Then
        Mesaj "Teslim tarihi sıkım tarihinden önce olamaz.", vbExclamation: Exit Function
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
    net = Nz(DLookup("MusteriKalanYag_Kg", "tbl_Parti", "PartiNo=" & id), 0)
    If Sayi(frm!txtTeslimYag) > net + 0.001 Then
        If Not Soru("Teslim edilen yağ (" & KgMetin(Sayi(frm!txtTeslimYag)) & ") müşteriye kalan yağdan (" & KgMetin(net) & ") fazla. Devam edilsin mi?") Then Exit Function
    End If

    ' Tahsilat
    ucretli = UcretliMi(frm!txtOdemeTipi)
    tur = "Yok"
    If ucretli Then
        Select Case Nz(frm!cmbTahsilat, "")
            Case "Peşin": tur = "Peşin"
            Case "Kısmi": tur = "Kısmi"
        End Select
    End If
    borcOnce = PartiKalanBorc(id) + OncekiTeslimTahsilati
    If tur = "Peşin" Then
        tutar = borcOnce
        If tutar <= 0 Then
            Mesaj "Bu partinin tahsil edilecek borcu kalmamış.", vbExclamation: Exit Function
        End If
    ElseIf tur = "Kısmi" Then
        tutar = Sayi(frm!txtTahsilatTutar)
        If tutar <= 0 Then
            Mesaj "Kısmi tahsilat tutarını giriniz.", vbExclamation: Exit Function
        End If
        If tutar >= borcOnce Then
            If Not Soru("Girilen tutar (" & TLYaz(tutar) & ") kalan borcun (" & TLYaz(borcOnce) & ") tamamını karşılıyor." & vbCrLf & _
                        "Peşin tahsilat olarak kaydedilsin mi?") Then Exit Function
            tur = "Peşin"
            If tutar > borcOnce Then tutar = borcOnce
        End If
    End If
    If tur <> "Yok" Then
        If Bos(frm!cmbYontem) Then
            Mesaj "Tahsilat yöntemini (Nakit / Havale / Kart) seçiniz.", vbExclamation: Exit Function
        End If
        If Bos(frm!cmbHesap) Then
            Mesaj IIf(frm!cmbYontem = "Nakit", "Kasayı seçiniz.", "Banka hesabını seçiniz." & vbCrLf & _
                "(Banka hesapları Ayarlar ekranından eklenir.)"), vbExclamation: Exit Function
        End If
    End If

    DBEngine.BeginTrans: islemde = True
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Parti WHERE PartiNo=" & id, dbOpenDynaset)
    rs.Edit
    rs!TeslimTarihi = CDate(frm!txtTeslimTarihi)
    rs!TeslimAlan = Trim$(frm!txtTeslimAlan)
    rs!TeslimPersonelID = CLng(frm!cmbTeslimEden)
    rs!TeslimYag_Kg = Sayi(frm!txtTeslimYag)
    rs!TeslimAciklama = MetinVeyaNull(frm!txtTeslimAciklama)
    rs!TahsilatTuru = tur
    If tur = "Yok" Then
        rs!TahsilatTutari = Null
        rs!TahsilatYontemi = Null
        rs!TahsilatKasaID = Null
        rs!TahsilatTarihi = Null
    Else
        rs!TahsilatTutari = tutar
        rs!TahsilatYontemi = frm!cmbYontem
        rs!TahsilatKasaID = CLng(frm!cmbHesap)
        rs!TahsilatTarihi = CDate(frm!txtTeslimTarihi)
    End If
    rs!PesinTahsilEdildi = (tur = "Peşin")
    rs.Update
    rs.Close
    TeslimMuhasebeYaz id
    BidonlariEsle id
    DBEngine.CommitTrans: islemde = False

    frm!lstTeslim.Requery
    TeslimPartiSecildi
    If tur = "Yok" And ucretli Then
        Mesaj "Teslim kaydedildi." & vbCrLf & "Sıkım ücreti müşterinin borcunda kalıyor (" & frm!txtKalanBorc & ")."
    ElseIf tur = "Kısmi" Then
        Mesaj "Teslim kaydedildi." & vbCrLf & TLYaz(tutar) & " tahsil edildi (" & frm!cmbYontem & "), kalan borç: " & frm!txtKalanBorc & "."
    ElseIf tur = "Peşin" Then
        Mesaj "Teslim kaydedildi." & vbCrLf & TLYaz(tutar) & " tahsil edildi (" & frm!cmbYontem & ")."
    Else
        Mesaj "Teslim kaydedildi."
    End If
    Exit Function
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "Teslim kaydedilemedi: " & Err.Description, vbExclamation
End Function

Public Function TeslimGeriAl()
    On Error GoTo Hata
    Dim frm As Form, id As Long, islemde As Boolean, eskiKod As String, yeniKod As String, tahsilSil As Boolean
    Set frm = F()
    If Bos(frm!cmbParti) Then
        Mesaj "Önce parti seçin.", vbExclamation: Exit Function
    End If
    id = CLng(frm!cmbParti)
    If IsNull(DLookup("TeslimTarihi", "tbl_Parti", "PartiNo=" & id)) Then
        Mesaj "Bu partinin teslim kaydı yok.", vbExclamation: Exit Function
    End If
    If Not Soru("Bu partinin teslim kaydı geri alınsın mı?") Then Exit Function
    If Nz(DLookup("TahsilatTutari", "tbl_Parti", "PartiNo=" & id), 0) > 0 Then
        tahsilSil = Soru("Teslimde alınan " & TLYaz(DLookup("TahsilatTutari", "tbl_Parti", "PartiNo=" & id)) & _
            " tahsilat da silinsin mi?" & vbCrLf & "(Hayır derseniz tahsilat kayıtlı kalır.)")
    End If
    eskiKod = Nz(DLookup("PartiKodu", "tbl_Parti", "PartiNo=" & id), "")

    DBEngine.BeginTrans: islemde = True
    CurrentDb.Execute "UPDATE tbl_Parti SET TeslimTarihi=Null, TeslimAlan=Null, TeslimYag_Kg=Null, TeslimAciklama=Null, TeslimPersonelID=Null WHERE PartiNo=" & id, dbFailOnError
    If tahsilSil Then
        CurrentDb.Execute "UPDATE tbl_Parti SET TahsilatTuru='Yok', TahsilatTutari=Null, TahsilatYontemi=Null, TahsilatKasaID=Null, " & _
            "TahsilatTarihi=Null, PesinTahsilEdildi=False WHERE PartiNo=" & id, dbFailOnError
        TeslimMuhasebeYaz id
    End If
    yeniKod = PartiNoCakismaCoz(id)
    BidonlariEsle id
    DBEngine.CommitTrans: islemde = False

    frm!lstTeslim.Requery
    TeslimTemizle
    frm!cmbParti = id
    TeslimPartiSecildi
    If yeniKod <> eskiKod Then
        Mesaj "Teslim kaydı geri alındı." & vbCrLf & vbCrLf & "Parti No " & eskiKod & " bu sezonda başka bir açık partide kullanıldığı için" & _
            " bu partinin numarası " & yeniKod & " olarak değiştirildi."
    Else
        Mesaj "Teslim kaydı geri alındı."
    End If
    Exit Function
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "İşlem yapılamadı: " & Err.Description, vbExclamation
End Function

' Teslim fişi: "onizle" / "pdf" / "yazdir"
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
