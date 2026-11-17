Option Compare Database
Option Explicit

' =============================================================
' CARİ İŞLEM YÖNETİMİ   (form: frm_CariYonetim)
' -------------------------------------------------------------
' Ayarlar > Yönetici Araçları'ndan açılır, yalnızca yönetici.
' Soldaki listeden bir cari seçilince kart bilgileri, kayıt sayıları
' ve seçilen türdeki kayıtları (cari, parti, kasa/banka, alım/satım,
' stok) gösterilir.
'
' "Cariyi Tümüyle Sil" carinin kendisini ve ona bağlı her kaydı siler:
'   partiler ve bidonları, cari hareketleri, kasa / banka hareketleri,
'   alım / satım kayıtları, stok hareketleri (hakyağı girişleri dahil).
' Kasa / banka bakiyeleri ve stok miktarları buna göre değişir; ekranda
' önceden gösterilir. Onay için carinin adı ve yönetici şifresi istenir.
' Silmeden önce bu bilgisayardaki Yedek klasörüne otomatik yedek alınır.
' Silme bitince veri dosyasının sıkıştırılmış kopyası hazırlanır ve
' program yeniden başlar; açılışta bu kopya eski dosyanın yerine geçer,
' böylece silinen kayıtların izi veri dosyasında kalmaz.
' =============================================================

Private Const FORM_ADI As String = "frm_CariYonetim"
Private gSeciliID As Long

Private Function F() As Form
    Set F = Forms(FORM_ADI)
End Function

Public Function CariYonetimiAc()
    If Not OturumAdmin And Not TestModu Then
        Mesaj "Bu bölüm yalnızca yönetici (admin) içindir.", vbExclamation
        Exit Function
    End If
    DoCmd.OpenForm FORM_ADI
End Function

Public Function CariYonYukle()
    On Error Resume Next
    gSeciliID = 0
    F()!cmbGorunum = "Cari Hareketler"
    CariListesiYenile ""
    SecimGoster
End Function

' ---------------------------------------------------------------
' Cari listesi ve arama
' ---------------------------------------------------------------
Private Sub CariListesiYenile(ByVal ara As String)
    Dim k As String
    ara = Trim$(ara)
    If ara <> "" Then
        k = " WHERE m.AdSoyad Like " & SqlMetin("*" & ara & "*") & _
            " OR m.Telefon Like " & SqlMetin("*" & ara & "*") & _
            " OR m.TCKN_VKN Like " & SqlMetin("*" & ara & "*")
    End If
    F()!lstCariler.RowSource = "SELECT m.MusteriID, m.AdSoyad AS [Ad Soyad / Ünvan], m.Telefon, " & _
        "Format(Nz(m.AcilisBakiyesi,0)+Nz(b.Toplam,0),'#,##0.00') AS Bakiye " & _
        "FROM tbl_Musteri AS m LEFT JOIN (SELECT MusteriID, Sum(Nz(Borc,0)-Nz(Alacak,0)) AS Toplam " & _
        "FROM tbl_CariHareket WHERE Nz(IptalMi,False)=False GROUP BY MusteriID) AS b ON m.MusteriID=b.MusteriID" & _
        k & " ORDER BY m.AdSoyad;"
End Sub

Public Function CariYonAra()
    On Error Resume Next
    CariListesiYenile Nz(F()!txtAra.Text, "")
End Function

Public Function CariYonSecildi()
    On Error Resume Next
    gSeciliID = Nz(F()!lstCariler.Value, 0)
    SecimGoster
End Function

' ---------------------------------------------------------------
' Bir cariye bağlı kayıtları bulan koşullar
' ---------------------------------------------------------------
Private Function PartiSql(ByVal id As Long) As String
    PartiSql = "SELECT PartiNo FROM tbl_Parti WHERE MusteriID=" & id
End Function

' Cari / kasa / stok hareketleri için: cariye ya da partilerine ya da
' alım / satım kayıtlarına bağlı olanlar
Private Function HareketKosulu(ByVal id As Long) As String
    Dim p As String
    p = PartiSql(id)
    HareketKosulu = "(MusteriID=" & id & " OR PartiNo IN (" & p & ")" & _
        " OR (KaynakTipi IN ('" & KAYNAK_PARTI & "','" & KAYNAK_TESLIM & "') AND KaynakID IN (" & p & "))" & _
        " OR (KaynakTipi='" & KAYNAK_ALIMSATIM & "' AND KaynakID IN (SELECT IslemID FROM tbl_AlimSatim WHERE MusteriID=" & id & ")))"
End Function

Private Function KasaKosulu(ByVal id As Long) As String
    KasaKosulu = "(" & HareketKosulu(id) & " OR KasaHareketID IN (SELECT KaynakID FROM tbl_CariHareket WHERE MusteriID=" & _
        id & " AND KaynakTipi='" & KAYNAK_KASA & "'))"
End Function

Private Function Sayac(ByVal tablo As String, ByVal kosul As String) As Long
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) AS n FROM " & tablo & " WHERE " & kosul, dbOpenSnapshot)
    Sayac = Nz(rs!n, 0)
    rs.Close
End Function

' ---------------------------------------------------------------
' Seçili carinin bilgileri
' ---------------------------------------------------------------
Private Sub SecimGoster()
    On Error GoTo Hata
    Dim frm As Form, rs As DAO.Recordset, d As String
    Set frm = F()
    KutulariTemizle frm, Array("txtOnayAd", "txtSifre")
    frm!lblEtki.Caption = ""
    If gSeciliID = 0 Then
        frm!lblCari.Caption = "Soldaki listeden bir cari seçin"
        frm!lblCariDetay.Caption = ""
        frm!lblOzet.Caption = ""
        frm!lstKayitlar.RowSource = ""
        On Error Resume Next
        frm!txtAra.SetFocus
        frm!btnSil.Enabled = False
        Exit Sub
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Musteri WHERE MusteriID=" & gSeciliID, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        gSeciliID = 0
        SecimGoster
        Exit Sub
    End If
    frm!lblCari.Caption = Nz(rs!adSoyad, "(adsız)")
    d = "Bakiye: " & TLYaz(MusteriBakiye(gSeciliID))
    If Not Bos(rs!Telefon) Then d = d & "   •   Tel: " & rs!Telefon
    If Not Bos(rs!TCKN_VKN) Then d = d & "   •   TCKN/VKN: " & rs!TCKN_VKN
    If Not Bos(rs.Fields("IBAN")) Then d = d & "   •   IBAN: " & rs.Fields("IBAN")
    rs.Close
    frm!lblCariDetay.Caption = d
    frm!lblOzet.Caption = OzetMetni(gSeciliID)
    frm!lblEtki.Caption = EtkiMetni(gSeciliID)
    frm!btnSil.Enabled = True
    CariYonGorunum
    Exit Sub
Hata:
    Mesaj "Cari bilgileri gösterilemedi: " & Err.Description, vbExclamation
End Sub

Private Function OzetMetni(ByVal id As Long) As String
    Dim nP As Long, nBek As Long, s As String
    nP = Sayac("tbl_Parti", "MusteriID=" & id)
    nBek = Sayac("tbl_Parti", "MusteriID=" & id & " AND TeslimTarihi Is Null")
    s = nP & " parti" & IIf(nBek > 0, " (" & nBek & " teslim bekliyor)", "")
    s = s & "   •   " & Sayac("tbl_CariHareket", HareketKosulu(id)) & " cari hareket"
    s = s & "   •   " & Sayac("tbl_KasaHareket", KasaKosulu(id)) & " kasa / banka"
    s = s & "   •   " & Sayac("tbl_AlimSatim", "MusteriID=" & id) & " alım / satım"
    s = s & "   •   " & Sayac("tbl_StokHareket", HareketKosulu(id)) & " stok hareketi"
    OzetMetni = s
End Function

' Silmenin kasa / banka bakiyelerine ve stoklara etkisi
Private Function EtkiMetni(ByVal id As Long) As String
    Dim rs As DAO.Recordset, s As String, kasa As String, stok As String, simdi As Double, net As Double
    Dim nBek As Long, nBidon As Long, birim As String
    Set rs = CurrentDb.OpenRecordset("SELECT h.KasaID, k.KasaAdi, Sum(Nz(h.Tahsilat,0)-Nz(h.Odeme,0)) AS Net " & _
        "FROM tbl_KasaHareket AS h LEFT JOIN tbl_Kasa AS k ON h.KasaID=k.KasaID " & _
        "WHERE Nz(h.IptalMi,False)=False AND h.KasaHareketID IN (SELECT KasaHareketID FROM tbl_KasaHareket WHERE " & _
        KasaKosulu(id) & ") GROUP BY h.KasaID, k.KasaAdi;", dbOpenSnapshot)
    Do While Not rs.EOF
        net = Nz(rs!net, 0)
        If net <> 0 Then
            simdi = KasaBakiye(Nz(rs!kasaID, 0))
            kasa = kasa & IIf(kasa = "", "", "   |   ") & Nz(rs!KasaAdi, "Hesap") & ": " & _
                Format$(simdi, "#,##0.00") & " → " & Format$(simdi - net, "#,##0.00") & " TL"
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = CurrentDb.OpenRecordset("SELECT s.UrunID, u.UrunAdi, u.Birim, Sum(Nz(s.Giris,0)-Nz(s.Cikis,0)) AS Net " & _
        "FROM tbl_StokHareket AS s INNER JOIN tbl_Urun AS u ON s.UrunID=u.UrunID " & _
        "WHERE Nz(s.IptalMi,False)=False AND s.StokHareketID IN (SELECT StokHareketID FROM tbl_StokHareket WHERE " & _
        HareketKosulu(id) & ") GROUP BY s.UrunID, u.UrunAdi, u.Birim;", dbOpenSnapshot)
    Do While Not rs.EOF
        net = Nz(rs!net, 0)
        If net <> 0 Then
            simdi = UrunStok(rs!urunID)
            birim = Nz(rs!birim, "")
            stok = stok & IIf(stok = "", "", "   |   ") & Nz(rs!UrunAdi, "Ürün") & ": " & _
                Miktar(simdi) & " → " & Miktar(simdi - net) & IIf(birim = "", "", " " & birim) & _
                IIf(simdi - net < 0, "  (EKSİYE DÜŞER)", "")
        End If
        rs.MoveNext
    Loop
    rs.Close
    nBidon = Sayac("tbl_PartiBidon", "PartiNo IN (" & PartiSql(id) & ")")
    nBek = Sayac("tbl_Parti", "MusteriID=" & id & " AND TeslimTarihi Is Null")
    s = "Silinince değişecekler:" & vbCrLf
    s = s & "Kasa / Banka:  " & IIf(kasa = "", "değişmez", kasa) & vbCrLf
    s = s & "Stok:  " & IIf(stok = "", "değişmez", stok)
    If nBidon > 0 Then s = s & vbCrLf & nBidon & " bidon kaydı silinir, numaraları boşa çıkar."
    If nBek > 0 Then s = s & IIf(nBidon > 0, "  ", vbCrLf) & "DİKKAT: " & nBek & " partinin yağı henüz teslim edilmedi."
    EtkiMetni = s
End Function

' 110 -> "110", 87,5 -> "87,5"
Private Function Miktar(ByVal x As Double) As String
    If x = Int(x) Then Miktar = Format$(x, "#,##0") Else Miktar = Format$(x, "#,##0.0#")
End Function

' ---------------------------------------------------------------
' Kayıt listesi
' ---------------------------------------------------------------
Public Function CariYonGorunum()
    On Error GoTo Hata
    Dim frm As Form, sql As String, gen As String, n As Integer, id As Long
    Set frm = F()
    id = gSeciliID
    If id = 0 Then
        frm!lstKayitlar.RowSource = ""
        Exit Function
    End If
    Select Case Nz(frm!cmbGorunum, "Cari Hareketler")
        Case "Partiler"
            sql = "SELECT PartiNo, [Parti No], Kabul, [Net Zeytin], [Sıkım], [Çıkan Yağ], [Randıman %], [Ödeme], " & _
                  "[Ücret TL], [Kalan Yağ], Durum FROM qry_UretimListe WHERE PartiNo IN (" & PartiSql(id) & _
                  ") ORDER BY TarihHam DESC, PartiNo DESC;"
            n = 11: gen = "0;800;950;950;950;950;900;1000;1000;950;1300"
        Case "Kasa / Banka"
            sql = "SELECT KasaHareketID, Tarih, Parti, Hareket, [Yöntem], [Açıklama], Tahsilat, [Ödeme], Kaynak " & _
                  "FROM qry_KasaEkstre WHERE KasaHareketID IN (SELECT KasaHareketID FROM tbl_KasaHareket WHERE " & _
                  KasaKosulu(id) & ") ORDER BY TarihHam, KasaHareketID;"
            n = 9: gen = "0;950;650;1000;900;3050;1100;1100;1000"
        Case "Alım / Satım"
            sql = "SELECT IslemID, Tarih, [İşlem], [Ürün], Miktar, [Birim Fiyat], Toplam, [Peşin], [Açıklama] " & _
                  "FROM qry_AlimSatimDetay WHERE MusteriID=" & id & ";"
            n = 9: gen = "0;1000;800;1800;900;1100;1200;700;2200"
        Case "Stok Hareketleri"
            sql = "SELECT StokHareketID, Tarih, [Ürün], Parti, Hareket, [Açıklama], Giren, [Çıkan], Kaynak " & _
                  "FROM qry_StokEkstre WHERE StokHareketID IN (SELECT StokHareketID FROM tbl_StokHareket WHERE " & _
                  HareketKosulu(id) & ") ORDER BY TarihHam, StokHareketID;"
            n = 9: gen = "0;950;1600;650;1300;2700;850;850;950"
        Case Else
            sql = "SELECT CariHareketID, Tarih, Parti, Hareket, [Açıklama], [Borç], Alacak, Kalan, Kaynak " & _
                  "FROM qry_CariEkstre WHERE CariHareketID IN (SELECT CariHareketID FROM tbl_CariHareket WHERE " & _
                  HareketKosulu(id) & ") ORDER BY TarihHam, CariHareketID;"
            n = 9: gen = "0;950;650;1400;2700;1050;1050;1050;950"
    End Select
    frm!lstKayitlar.ColumnCount = n
    frm!lstKayitlar.ColumnWidths = gen
    frm!lstKayitlar.RowSource = sql
    Exit Function
Hata:
    Mesaj "Kayıtlar listelenemedi: " & Err.Description, vbExclamation
End Function

' ---------------------------------------------------------------
' Tümüyle silme
' ---------------------------------------------------------------
Public Function CariTumuyleSil()
    On Error GoTo Hata
    Dim frm As Form, id As Long, ad As String, saklanan As String, metin As String, yedek As String
    Dim ws As DAO.Workspace, db As DAO.Database, islemde As Boolean
    Set frm = F()
    id = gSeciliID
    If id = 0 Then
        Mesaj "Önce soldaki listeden silinecek cariyi seçin.", vbExclamation: Exit Function
    End If
    If Not OturumAdmin Or OturumID = 0 Then
        Mesaj "Bu işlemi yalnızca giriş yapmış bir yönetici yapabilir.", vbExclamation: Exit Function
    End If
    ad = Nz(DLookup("AdSoyad", "tbl_Musteri", "MusteriID=" & id), "")
    If BuyukHarf(Trim$(Nz(frm!txtOnayAd, ""))) <> BuyukHarf(Trim$(ad)) Then
        Mesaj "Onay için carinin adını listede göründüğü gibi yazın:" & vbCrLf & ad, vbExclamation
        frm!txtOnayAd.SetFocus
        Exit Function
    End If
    saklanan = Nz(DLookup("Sifre", "tbl_Kullanici", "KullaniciID=" & OturumID), "")
    If Not SifreDogruMu(Nz(frm!txtSifre, ""), saklanan) Then
        frm!txtSifre = Null
        Mesaj "Yönetici şifresi hatalı.", vbExclamation
        frm!txtSifre.SetFocus
        Exit Function
    End If

    metin = ad & " ve ona ait TÜM kayıtlar silinecek:" & vbCrLf & vbCrLf & _
            Replace(OzetMetni(id), "   •   ", vbCrLf) & vbCrLf & vbCrLf & _
            Replace(EtkiMetni(id), "   |   ", vbCrLf & "      ") & vbCrLf & vbCrLf & _
            "Bu işlem geri alınamaz. Silmeden önce bu bilgisayardaki Yedek klasörüne otomatik yedek alınır; " & _
            "silme bitince program yeniden başlar." & vbCrLf & vbCrLf & "Devam edilsin mi?"
    If Mesaj(metin, vbYesNo + vbExclamation + vbDefaultButton2) <> vbYes Then Exit Function

    DoCmd.Hourglass True
    ' 1. Otomatik yedek
    yedek = CariSilmeYedegiAl()

    ' 2. Silme (tek işlem: biri olmazsa hiçbiri olmaz)
    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb
    ws.BeginTrans
    islemde = True
    db.Execute "DELETE FROM tbl_StokHareket WHERE " & HareketKosulu(id), dbFailOnError
    db.Execute "DELETE FROM tbl_KasaHareket WHERE " & KasaKosulu(id), dbFailOnError
    db.Execute "DELETE FROM tbl_CariHareket WHERE " & HareketKosulu(id), dbFailOnError
    db.Execute "DELETE FROM tbl_PartiBidon WHERE PartiNo IN (" & PartiSql(id) & ")", dbFailOnError
    db.Execute "DELETE FROM tbl_AlimSatim WHERE MusteriID=" & id, dbFailOnError
    db.Execute "DELETE FROM tbl_Parti WHERE MusteriID=" & id, dbFailOnError
    db.Execute "DELETE FROM tbl_Musteri WHERE MusteriID=" & id, dbFailOnError
    ws.CommitTrans
    islemde = False
    DBEngine.Idle dbRefreshCache

    ' 3. Silinen kayıtların izini temizlemek için sıkıştırılmış kopyayı hazırla, programı yeniden başlat
    VeriSikistirmaHazirla
    DoCmd.Hourglass False
    Mesaj ad & " ve tüm kayıtları silindi." & vbCrLf & vbCrLf & _
          "Silmeden önce alınan yedek:" & vbCrLf & yedek & vbCrLf & vbCrLf & _
          "Program şimdi yeniden başlatılacak ve veri dosyası temizlenecek."
    ProgramiYenidenBaslat
    Exit Function
Hata:
    Dim hataMetni As String
    hataMetni = Err.Description
    On Error Resume Next
    If islemde Then ws.Rollback
    DoCmd.Hourglass False
    If yedek = "" Then
        Mesaj "Yedek alınamadığı için silme yapılmadı: " & hataMetni, vbExclamation
    ElseIf islemde Then
        Mesaj "Silme sırasında hata oluştu, hiçbir kayıt silinmedi: " & hataMetni, vbExclamation
    Else
        Mesaj "Cari silindi ancak veri dosyası temizlenemedi: " & hataMetni & vbCrLf & _
              "Program bir sonraki açılışta normal çalışır.", vbExclamation
        CariListesiYenile ""
        gSeciliID = 0
        SecimGoster
    End If
End Function
