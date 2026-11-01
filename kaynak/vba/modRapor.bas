Option Compare Database
Option Explicit

' =============================================================
' RAPORLAR   (form: frm_Raporlar, raporlar: rpt_...)
' -------------------------------------------------------------
' Tüm raporlar A5 boyutundadır. Rapor açılırken (OnOpen =RaporHazirla())
' kayıt kaynağı burada hazırlanan SQL ile değiştirilir; başlık ve
' tarih aralığı RaporBasligi() / RaporAraligi() ile yazılır.
'
' Çıktı türleri:
'   "onizle" : ekranda önizleme
'   "pdf"    : program klasöründeki \Raporlar altına PDF kaydeder ve açar
'   "yazdir" : önizlemeyi açar ve yazıcı seçme penceresini gösterir
'
' Yeni rapor eklemek için: RaporListesi'ne satır ekleyin, RaporSQL'e
' sorgusunu yazın ve rpt_ raporunu tasarlayın (OnOpen =RaporHazirla()).
' =============================================================

Private Const FORM_ADI As String = "frm_Raporlar"

Public RaporAktif As String       ' açılan raporun adı
Public RaporKaynak As String      ' raporun SQL kaynağı
Public RaporBaslikMetni As String
Public RaporAralikMetni As String

Private Function F() As Form
    Set F = Forms(FORM_ADI)
End Function

' Rapor kodu ; rapor adı (listede görünen)
Public Function RaporListesi() As String
    RaporListesi = "URETIM;Üretim Raporu (Gelen Zeytin / Çıkan Yağ);" & _
        "UCRET;Sıkım Ücretleri (Tahsil Edilen / Edilmeyen);" & _
        "BIDON;Müşteri Bidonları (Teslim Edilen / Bekleyen);" & _
        "TESLIMLER;Yağ Teslimleri;" & _
        "ALACAK;Alacaklar ve Borçlar;" & _
        "CARI;Cari Hesap Ekstresi (Müşteri Bazında);" & _
        "STOK;Stok Hareketleri (Ürün Bazında Giriş-Çıkış);" & _
        "KASA;Kasa Raporu (Tahsilat ve Ödemeler)"
End Function

Private Function RaporNesnesi(ByVal kod As String) As String
    Select Case kod
        Case "URETIM": RaporNesnesi = "rpt_Uretim"
        Case "UCRET": RaporNesnesi = "rpt_Ucretler"
        Case "BIDON": RaporNesnesi = "rpt_Bidonlar"
        Case "TESLIMLER": RaporNesnesi = "rpt_Teslimler"
        Case "ALACAK": RaporNesnesi = "rpt_AlacakBorc"
        Case "CARI": RaporNesnesi = "rpt_CariEkstre"
        Case "STOK": RaporNesnesi = "rpt_StokHareket"
        Case "KASA": RaporNesnesi = "rpt_Kasa"
        Case "FIS": RaporNesnesi = "rpt_TeslimFisi"
    End Select
End Function

Private Function RaporAciklamasi(ByVal kod As String) As String
    Select Case kod
        Case "URETIM": RaporAciklamasi = "Seçilen tarihler arasında gelen partiler: gelen zeytin, çıkan yağ, randıman, hakyağı ve sıkım ücretleri. İsterseniz müşteriye göre süzün."
        Case "UCRET": RaporAciklamasi = "Sıkım ücretli partilerin ücretleri; tahsil edilenler ve edilmeyenler ayrı gruplarda toplanır."
        Case "BIDON": RaporAciklamasi = "Müşteri bazında bidonlar: yağı teslim edilenler ve teslim bekleyenler."
        Case "TESLIMLER": RaporAciklamasi = "Seçilen tarihler arasında yapılan yağ teslimleri: teslim alan, teslim eden ve miktarlar."
        Case "ALACAK": RaporAciklamasi = "Bitiş tarihi itibarıyla her carinin bakiyesi: alacaklarımız ve borçlarımız. Başlangıç tarihi kullanılmaz."
        Case "CARI": RaporAciklamasi = "Müşteri bazında cari hareket dökümü: devir, borç, alacak ve yürüyen bakiye."
        Case "STOK": RaporAciklamasi = "Ürün bazında stok giriş-çıkış dökümü: devir, giren, çıkan ve kalan."
        Case "KASA": RaporAciklamasi = "Kasa tahsilat ve ödemeleri: devir, tahsilat, ödeme ve yürüyen bakiye."
    End Select
End Function

' ---------------------------------------------------------------
' Rapor içinden çağrılan fonksiyonlar
' ---------------------------------------------------------------
Public Function RaporBasligi() As String
    RaporBasligi = RaporBaslikMetni
End Function

Public Function RaporAraligi() As String
    RaporAraligi = RaporAralikMetni
End Function

Public Function FirmaBilgiSatiri() As String
    Dim s As String
    s = Ayar("FirmaAdres")
    If Ayar("FirmaTelefon") <> "" Then s = s & IIf(s = "", "", "   •   ") & "Tel: " & Ayar("FirmaTelefon")
    FirmaBilgiSatiri = s
End Function

Public Function RaporAltBilgi() As String
    RaporAltBilgi = URUN_ADI & " v" & URUN_SURUM
End Function

' Raporun OnOpen olayı
Public Function RaporHazirla()
    On Error Resume Next
    Dim r As Report
    Set r = Reports(RaporAktif)
    If RaporKaynak <> "" Then r.RecordSource = RaporKaynak
    ' A5 kağıt (yazıcı değişmişse yeniden uygula)
    If r.Printer.PaperSize <> acPRPSA5 Then
        r.Printer.PaperSize = acPRPSA5
        If r.Tag = "yatay" Then r.Printer.Orientation = acPRORLandscape Else r.Printer.Orientation = acPRORPortrait
    End If
End Function

' ---------------------------------------------------------------
' Rapor SQL'leri
' ---------------------------------------------------------------
Private Function SqlSayi(ByVal d As Double) As String
    SqlSayi = Trim$(Str$(d))
End Function

Private Function RaporSQL(ByVal kod As String, ByVal bas As Date, ByVal bit As Date, _
        ByVal musteri As Variant, ByVal urun As Variant, ByVal durum As String) As String
    Dim w As String, aralik As String, devir As String
    aralik = " Between " & SqlTarih(bas) & " And " & SqlTarih(bit)
    Select Case kod
    Case "URETIM"
        w = "Tarih" & aralik
        If Not IsNull(musteri) Then w = w & " AND MusteriID=" & musteri
        RaporSQL = "SELECT * FROM qry_R_Uretim WHERE " & w

    Case "UCRET"
        w = "Tarih" & aralik
        If Not IsNull(musteri) Then w = w & " AND MusteriID=" & musteri
        If durum = "Tahsil Edilen" Or durum = "Tahsil Edilmeyen" Then w = w & " AND Durum=" & SqlMetin(durum)
        RaporSQL = "SELECT * FROM qry_R_Ucret WHERE " & w

    Case "BIDON"
        w = "Tarih" & aralik
        If Not IsNull(musteri) Then w = w & " AND MusteriID=" & musteri
        If durum = "Teslim Edilen" Then w = w & " AND Durum='Teslim Edildi'"
        If durum = "Bekleyen" Then w = w & " AND Durum='Bekliyor'"
        RaporSQL = "SELECT * FROM qry_R_Bidon WHERE " & w

    Case "TESLIMLER"
        w = "TeslimTarihi" & aralik
        If Not IsNull(musteri) Then w = w & " AND MusteriID=" & musteri
        RaporSQL = "SELECT * FROM qry_R_Teslim WHERE " & w

    Case "ALACAK"
        RaporSQL = "SELECT x.*, IIf(x.Bakiye>0,x.Bakiye,0) AS Alacagimiz, IIf(x.Bakiye<0,-x.Bakiye,0) AS Borcumuz FROM (" & _
            "SELECT m.MusteriID, m.AdSoyad AS Musteri, m.Telefon, Round(Nz(m.AcilisBakiyesi,0)+Nz(DSum('Nz(Borc,0)-Nz(Alacak,0)','tbl_CariHareket'," & _
            "'Nz(IptalMi,False)=False AND MusteriID=' & m.MusteriID & ' AND Tarih<=" & SqlSayi(CDbl(bit)) & "'),0),2) AS Bakiye FROM tbl_Musteri AS m" & _
            IIf(IsNull(musteri), "", " WHERE m.MusteriID=" & musteri) & ") AS x" & _
            IIf(durum = "Tümü", "", " WHERE Abs(x.Bakiye)>0.005")

    Case "CARI"
        devir = "Nz(DLookup('AcilisBakiyesi','tbl_Musteri','MusteriID=' & q.MusteriID),0)+Nz(DSum('Nz(Borc,0)-Nz(Alacak,0)','tbl_CariHareket'," & _
            "'Nz(IptalMi,False)=False AND MusteriID=' & q.MusteriID & ' AND Tarih<" & SqlSayi(CDbl(bas)) & "'),0)"
        w = "q.Tarih" & aralik
        If Not IsNull(musteri) Then w = w & " AND q.MusteriID=" & musteri
        RaporSQL = "SELECT q.*, " & devir & " AS Devir FROM qry_R_Cari AS q WHERE " & w

    Case "STOK"
        devir = "Nz(DLookup('AcilisStok','tbl_Urun','UrunID=' & q.UrunID),0)+Nz(DSum('Nz(Giris,0)-Nz(Cikis,0)','tbl_StokHareket'," & _
            "'Nz(IptalMi,False)=False AND UrunID=' & q.UrunID & ' AND Tarih<" & SqlSayi(CDbl(bas)) & "'),0)"
        w = "q.Tarih" & aralik
        If Not IsNull(urun) Then w = w & " AND q.UrunID=" & urun
        RaporSQL = "SELECT q.*, " & devir & " AS Devir FROM qry_R_Stok AS q WHERE " & w

    Case "KASA"
        devir = SqlSayi(Nz(DSum("Nz(AcilisBakiyesi,0)", "tbl_Kasa"), 0) + _
            Nz(DSum("Nz(Tahsilat,0)-Nz(Odeme,0)", "tbl_KasaHareket", "Nz(IptalMi,False)=False AND Tarih<" & SqlTarih(bas)), 0))
        RaporSQL = "SELECT q.*, " & devir & " AS Devir FROM qry_R_Kasa AS q WHERE q.Tarih" & aralik
    End Select
End Function

' ---------------------------------------------------------------
' Çıktı
' ---------------------------------------------------------------
Private Function DosyaAdiTemizle(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|", "(", ")")
        s = Replace(s, c, "")
    Next
    s = Replace(Trim$(s), " ", "_")
    DosyaAdiTemizle = s
End Function

Private Function PdfYolu(ByVal ad As String) As String
    Dim klasor As String
    klasor = CurrentProject.Path & "\Raporlar"
    If Dir(klasor, vbDirectory) = "" Then MkDir klasor
    PdfYolu = klasor & "\" & DosyaAdiTemizle(ad) & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
End Function

Private Function KayitVarMi(ByVal sql As String) As Boolean
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    KayitVarMi = Not rs.EOF
    rs.Close
End Function

' Raporu hazırlar ve istenen çıktıyı verir. PDF'te dosya yolunu döndürür.
Public Function RaporCiktisi(ByVal rapor As String, ByVal sql As String, ByVal baslik As String, _
        ByVal aralik As String, ByVal cikti As String) As String
    On Error GoTo Hata
    Dim yol As String
    If Not KayitVarMi(sql) Then
        Mesaj "Seçilen ölçütlere uygun kayıt bulunamadı.", vbInformation
        Exit Function
    End If
    RaporAktif = rapor
    RaporKaynak = sql
    RaporBaslikMetni = baslik
    RaporAralikMetni = aralik
    If CurrentProject.AllReports(rapor).IsLoaded Then DoCmd.Close acReport, rapor

    Select Case cikti
    Case "pdf"
        yol = PdfYolu(baslik)
        DoCmd.OutputTo acOutputReport, rapor, acFormatPDF, yol, Not TestModu
        RaporCiktisi = yol
        If Not TestModu Then SonMesaj = yol
    Case "yazdir"
        DoCmd.OpenReport rapor, acViewPreview
        If Not TestModu Then
            On Error Resume Next
            DoCmd.RunCommand acCmdPrint     ' yazıcı seçme penceresi (A5)
            On Error GoTo Hata
        End If
    Case Else
        DoCmd.OpenReport rapor, acViewPreview
    End Select
    Exit Function
Hata:
    If Err.Number = 2501 Then Exit Function      ' kullanıcı iptal etti
    Mesaj "Rapor hazırlanamadı: " & Err.Description, vbExclamation
End Function

' ---------------------------------------------------------------
' RAPOR MERKEZİ FORMU
' ---------------------------------------------------------------
Public Function RaporlarYukle()
    On Error Resume Next
    Dim frm As Form
    Set frm = F()
    frm!lstRaporlar.RowSource = RaporListesi()
    frm!txtBas = DateSerial(Year(Date), Month(Date), 1)
    frm!txtBit = Date
    frm!lstRaporlar = "URETIM"
    RaporSecildi
End Function

Public Function RaporSecildi()
    On Error Resume Next
    Dim frm As Form, kod As String, i As Long
    Set frm = F()
    kod = Nz(frm!lstRaporlar, "")
    frm!lblSecilen.Caption = ""
    For i = 0 To frm!lstRaporlar.ListCount - 1
        If frm!lstRaporlar.Column(0, i) = kod Then frm!lblSecilen.Caption = frm!lstRaporlar.Column(1, i)
    Next
    frm!lblAciklama.Caption = RaporAciklamasi(kod)
    frm!cmbRMusteri.Enabled = (kod = "URETIM" Or kod = "UCRET" Or kod = "BIDON" Or kod = "TESLIMLER" Or kod = "ALACAK" Or kod = "CARI")
    frm!cmbRUrun.Enabled = (kod = "STOK")
    frm!txtBas.Enabled = (kod <> "ALACAK")
    Select Case kod
        Case "UCRET": frm!cmbRDurum.RowSource = "Tümü;Tahsil Edilen;Tahsil Edilmeyen"
        Case "BIDON": frm!cmbRDurum.RowSource = "Tümü;Teslim Edilen;Bekleyen"
        Case "ALACAK": frm!cmbRDurum.RowSource = "Bakiyesi Olanlar;Tümü"
        Case Else: frm!cmbRDurum.RowSource = ""
    End Select
    frm!cmbRDurum.Enabled = (frm!cmbRDurum.RowSource <> "")
    If frm!cmbRDurum.Enabled Then frm!cmbRDurum = frm!cmbRDurum.ItemData(0) Else frm!cmbRDurum = Null
    If Not frm!cmbRMusteri.Enabled Then frm!cmbRMusteri = Null
    If Not frm!cmbRUrun.Enabled Then frm!cmbRUrun = Null
End Function

Public Function RaporHizliTarih(ByVal tur As String)
    Dim frm As Form
    Set frm = F()
    Select Case tur
        Case "gun": frm!txtBas = Date: frm!txtBit = Date
        Case "hafta": frm!txtBas = Date - Weekday(Date, vbMonday) + 1: frm!txtBit = Date
        Case "ay": frm!txtBas = DateSerial(Year(Date), Month(Date), 1): frm!txtBit = Date
        Case "yil": frm!txtBas = DateSerial(Year(Date), 1, 1): frm!txtBit = Date
    End Select
End Function

Public Function RaporCalistir(ByVal cikti As String)
    On Error GoTo Hata
    Dim frm As Form, kod As String, bas As Date, bit As Date, sql As String, aralik As String
    Set frm = F()
    kod = Nz(frm!lstRaporlar, "")
    If kod = "" Then Mesaj "Listeden bir rapor seçin.", vbExclamation: Exit Function
    If Not TarihMi(frm!txtBit) Or (kod <> "ALACAK" And Not TarihMi(frm!txtBas)) Then
        Mesaj "Başlangıç ve bitiş tarihlerini giriniz.", vbExclamation: Exit Function
    End If
    bit = CDate(frm!txtBit)
    If kod = "ALACAK" Then bas = bit Else bas = CDate(frm!txtBas)
    If bas > bit Then Mesaj "Başlangıç tarihi bitişten sonra olamaz.", vbExclamation: Exit Function

    sql = RaporSQL(kod, bas, bit, SayiVeyaNull(frm!cmbRMusteri), SayiVeyaNull(frm!cmbRUrun), Nz(frm!cmbRDurum, "Tümü"))
    If kod = "ALACAK" Then
        aralik = Format$(bit, "dd.mm.yyyy") & " tarihi itibarıyla"
    Else
        aralik = Format$(bas, "dd.mm.yyyy") & " - " & Format$(bit, "dd.mm.yyyy")
    End If
    If Not IsNull(frm!cmbRMusteri) Then aralik = aralik & "   •   Müşteri: " & frm!cmbRMusteri.Column(1)
    If Not IsNull(frm!cmbRUrun) Then aralik = aralik & "   •   Ürün: " & frm!cmbRUrun.Column(1)
    If Not IsNull(frm!cmbRDurum) And Nz(frm!cmbRDurum, "") <> "Tümü" And kod <> "ALACAK" Then aralik = aralik & "   •   " & frm!cmbRDurum

    RaporCalistir = RaporCiktisi(RaporNesnesi(kod), sql, frm!lblSecilen.Caption, aralik, cikti)
    Exit Function
Hata:
    Mesaj "Rapor hazırlanamadı: " & Err.Description, vbExclamation
End Function

' Teslim modülünden çağrılır
Public Function TeslimFisiAc(ByVal partiNo As Long, ByVal cikti As String) As String
    TeslimFisiAc = RaporCiktisi("rpt_TeslimFisi", "SELECT * FROM qry_R_TeslimFisi WHERE PartiNo=" & partiNo, _
        "Yağ Teslim Fişi - Parti " & Nz(DLookup("PartiKodu", "tbl_Parti", "PartiNo=" & partiNo), ""), _
        "Fiş tarihi: " & Format$(Date, "dd.mm.yyyy"), cikti)
End Function

' Test / otomasyon için: raporu doğrudan kodla PDF'e çıkarır
Public Function RaporTestPdf(ByVal kod As String) As String
    Dim bas As Date, bit As Date
    bas = DateSerial(2000, 1, 1): bit = DateSerial(2100, 12, 31)
    RaporTestPdf = RaporCiktisi(RaporNesnesi(kod), RaporSQL(kod, bas, bit, Null, Null, "Tümü"), "Test " & kod, "Tüm kayıtlar", "pdf")
End Function