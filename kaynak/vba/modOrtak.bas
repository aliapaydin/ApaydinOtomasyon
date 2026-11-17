Option Compare Database
Option Explicit

' =============================================================
' ZEYTİNYAĞI ÜRETİM OTOMASYONU (Apaydın Otomasyon) - ORTAK FONKSİYONLAR
' -------------------------------------------------------------
' Formlardaki düğmeler "=FonksiyonAdi()" şeklinde bu ve diğer
' mod... modüllerindeki Public Function'ları çağırır.
'
' Bu modülde:
'   Mesaj / Soru            : mesaj kutuları
'   Bos / Sayi / ...        : form kutularındaki değerleri okuma
'   CariYaz/KasaYaz/StokYaz : hareket tablolarına kayıt yazma
'   OtomatikKayitlariSil    : bir kaynağın (parti, alım-satım, kasa)
'                             otomatik oluşturduğu kayıtları siler
'
' Hareket tablolarındaki alanlar:
'   KaynakTipi : PARTI / ALIMSATIM / KASA / MANUEL
'   KaynakID   : kaynağın numarası (PartiNo, IslemID, KasaHareketID)
'   OtomatikMi : True ise kayıt programca üretilmiştir; elle
'                değiştirilmez, kaynağı güncellenince yeniden yazılır.
' =============================================================

Public Const KAYNAK_PARTI As String = "PARTI"
Public Const KAYNAK_ALIMSATIM As String = "ALIMSATIM"
Public Const KAYNAK_KASA As String = "KASA"
Public Const KAYNAK_MANUEL As String = "MANUEL"
Public Const KAYNAK_TESLIM As String = "TESLIM"

' Ürün ve geliştirici bilgileri (Hakkında ekranında gösterilir, değiştirilemez)
Public Const URUN_ADI As String = "Zeytinyağı Üretim Otomasyonu"
Public Const URUN_SURUM As String = "3.0"
Public Const GELISTIRICI As String = "Ali Apaydın"
Public Const GELISTIRICI_EPOSTA As String = "apaydin.a@gmail.com"
Public Const GELISTIRICI_TELEFON As String = "0553 605 35 00"
Public Const GELISTIRICI_GITHUB As String = "github.com/aliapaydin"

' Unicode mesaj kutusu: VBA'nın MsgBox'ı Türkçe olmayan Windows'ta
' ı, ş, ğ harflerini "?" gösterir; MessageBoxW her dilde doğru gösterir.
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, _
    ByVal lpCaption As LongPtr, ByVal uType As Long) As Long

Public Const ANA_MENU As String = "frm_CicerAnaMenu"
Public Const GIRIS_FORMU As String = "frm_Giris"

' Oturum: giriş yapan kullanıcı
Public OturumID As Long
Public OturumAd As String
Public OturumAdmin As Boolean

' Test modunda mesaj kutusu açılmaz, sorulara "Evet" denir.
' (Otomatik deneme içindir; normal kullanımda False kalır.)
Public TestModu As Boolean
Public SonMesaj As String

Public Sub TestModuAyarla(ByVal acik As Boolean)
    TestModu = acik
    SonMesaj = ""
End Sub

Public Function TestModuAc()
    TestModuAyarla True
End Function

Public Function SonMesajiAl() As String
    SonMesajiAl = SonMesaj
End Function

' ---------------------------------------------------------------
' Mesaj kutuları
' ---------------------------------------------------------------
Public Function Mesaj(ByVal metin As String, Optional ByVal stil As VbMsgBoxStyle = vbInformation) As VbMsgBoxResult
    SonMesaj = metin
    If TestModu Then
        If (stil And 7) = vbYesNo Then
            Mesaj = vbYes
        Else
            Mesaj = vbOK
        End If
    Else
        Dim baslik As String
        baslik = ProgramBasligi()
        Mesaj = MessageBoxW(Application.hWndAccessApp, StrPtr(metin), StrPtr(baslik), stil Or &H2000&)  ' MB_TASKMODAL
    End If
End Function

Public Function Soru(ByVal metin As String) As Boolean
    Soru = (Mesaj(metin, vbYesNo + vbQuestion) = vbYes)
End Function

' ---------------------------------------------------------------
' Değer okuma yardımcıları
' ---------------------------------------------------------------
' Form kutusu fonksiyona nesne olarak gelirse değerini alır
Public Function deger(ByVal v As Variant) As Variant
    If IsObject(v) Then
        deger = v.Value
    Else
        deger = v
    End If
End Function

Public Function Bos(ByVal v As Variant) As Boolean
    v = deger(v)
    If IsNull(v) Then
        Bos = True
    ElseIf IsEmpty(v) Then
        Bos = True
    Else
        Bos = (Trim$(CStr(v)) = "")
    End If
End Function

' Kutudaki değeri sayıya çevirir; boşsa veya sayı değilse 0 döner.
Public Function Sayi(ByVal v As Variant) As Double
    On Error GoTo Hata
    v = deger(v)
    If Bos(v) Then Exit Function
    Sayi = CDbl(v)
    Exit Function
Hata:
    Sayi = 0
End Function

Public Function SayiMi(ByVal v As Variant) As Boolean
    v = deger(v)
    If Bos(v) Then Exit Function
    SayiMi = IsNumeric(v)
End Function

Public Function TarihMi(ByVal v As Variant) As Boolean
    v = deger(v)
    If Bos(v) Then Exit Function
    TarihMi = IsDate(v)
End Function

Public Function MetinVeyaNull(ByVal v As Variant) As Variant
    v = deger(v)
    If Bos(v) Then
        MetinVeyaNull = Null
    Else
        MetinVeyaNull = Trim$(CStr(v))
    End If
End Function

Public Function SayiVeyaNull(ByVal v As Variant) As Variant
    v = deger(v)
    If SayiMi(v) Then
        SayiVeyaNull = CDbl(v)
    Else
        SayiVeyaNull = Null
    End If
End Function

Public Function TarihVeyaNull(ByVal v As Variant) As Variant
    v = deger(v)
    If TarihMi(v) Then
        TarihVeyaNull = CDate(v)
    Else
        TarihVeyaNull = Null
    End If
End Function

' Türkçe karakterleri sadeleştirip küçük harfe çevirir.
' "Sıkım Ücreti", "Sikim Ucreti", "SIKIM ÜCRETİ" -> "sikim ucreti"
Public Function Sade(ByVal v As Variant) As String
    Dim s As String
    v = deger(v)
    If Bos(v) Then Exit Function
    s = Trim$(CStr(v))
    s = Replace(s, "ı", "i"): s = Replace(s, "İ", "i"): s = Replace(s, "I", "i")
    s = Replace(s, "ş", "s"): s = Replace(s, "Ş", "s")
    s = Replace(s, "ğ", "g"): s = Replace(s, "Ğ", "g")
    s = Replace(s, "ü", "u"): s = Replace(s, "Ü", "u")
    s = Replace(s, "ö", "o"): s = Replace(s, "Ö", "o")
    s = Replace(s, "ç", "c"): s = Replace(s, "Ç", "c")
    Sade = LCase$(s)
End Function

' Ödeme tipi "Sıkım Ücreti" mi?
Public Function UcretliMi(ByVal odemeTipi As Variant) As Boolean
    Dim s As String
    s = Sade(odemeTipi)
    UcretliMi = (InStr(s, "sikim") > 0 Or InStr(s, "ucret") > 0)
End Function

' Ödeme tipi "Hakyağı" mı?
Public Function HakyagiMi(ByVal odemeTipi As Variant) As Boolean
    HakyagiMi = (InStr(Sade(odemeTipi), "hak") > 0)
End Function

' SQL içinde kullanmak için tarih: #aa/gg/yyyy#
Public Function SqlTarih(ByVal d As Date) As String
    SqlTarih = "#" & Format$(d, "mm\/dd\/yyyy") & "#"
End Function

' SQL içinde kullanmak için metin: 'O''Brien'
Public Function SqlMetin(ByVal s As String) As String
    SqlMetin = "'" & Replace(s, "'", "''") & "'"
End Function

' Kilogram yazısı: 200 -> "200 kg", 87,1 -> "87,1 kg"
Public Function KgMetin(ByVal v As Variant) As String
    Dim x As Double
    x = Sayi(v)
    If x = Int(x) Then
        KgMetin = Format$(x, "#,##0") & " kg"
    Else
        KgMetin = Format$(x, "#,##0.0#") & " kg"
    End If
End Function

Public Function TLYaz(ByVal tutar As Variant) As String
    TLYaz = Format$(Nz(tutar, 0), "#,##0.00") & " TL"
End Function

' ---------------------------------------------------------------
' Form yardımcıları
' ---------------------------------------------------------------
Public Sub KutulariTemizle(ByVal frm As Form, ByVal adlar As Variant)
    Dim ad As Variant
    For Each ad In adlar
        frm.Controls(CStr(ad)).Value = Null
    Next
End Sub

' Listede seçili satırın ilk (gizli) sütunundaki ID
Public Function ListeID(ByVal lst As Control) As Long
    If lst.ListIndex < 0 Then Exit Function
    If IsNull(lst.Column(0)) Then Exit Function
    ListeID = CLng(lst.Column(0))
End Function

' Açık formlardaki tablo/sorgu bağlı açılır kutuları yeniler
' (yeni müşteri / ürün / personel eklenince diğer formlarda görünsün)
Public Sub KombolariYenile()
    On Error Resume Next
    Dim frm As Form, ctl As Control
    For Each frm In Forms
        For Each ctl In frm.Controls
            If ctl.ControlType = acComboBox Then
                If ctl.RowSourceType = "Table/Query" Then ctl.Requery
            End If
        Next
    Next
End Sub

Public Function FormAcikMi(ByVal ad As String) As Boolean
    On Error Resume Next
    FormAcikMi = CurrentProject.AllForms(ad).IsLoaded
End Function

' ---------------------------------------------------------------
' Menü
' ---------------------------------------------------------------
Public Function MenuAc(ByVal formAdi As String)
    On Error GoTo Hata
    If (formAdi = "frm_Kullanicilar" Or formAdi = "frm_Ayarlar") And Not OturumAdmin Then
        Mesaj "Bu bölüm yalnızca yönetici (admin) içindir.", vbExclamation
        Exit Function
    End If
    DoCmd.OpenForm formAdi
    Exit Function
Hata:
    Mesaj "Form açılamadı: " & formAdi & vbCrLf & Err.Description, vbExclamation
End Function

Public Function AktifFormuKapat()
    On Error Resume Next
    Dim ad As String
    ad = Screen.ActiveForm.Name
    If ad = "" Then Exit Function
    On Error GoTo 0
    If ad = ANA_MENU Or ad = GIRIS_FORMU Then
        If Soru("Programdan çıkılsın mı?") Then DoCmd.Quit acQuitSaveAll
    Else
        DoCmd.Close acForm, ad
        If Not FormAcikMi(ANA_MENU) Then DoCmd.OpenForm ANA_MENU
    End If
End Function

' Eski düğmeler için: aktif forma göre doğru "Yeni" işlemini çağırır.
Public Function YeniKayit()
    On Error Resume Next
    Select Case Screen.ActiveForm.Name
        Case "frm_MalKabul": MalKabulYeni
        Case "frm_CicerMusteri": MusteriYeni
        Case "frm_CicerUrun": UrunYeni
        Case "frm_CicerKasa": KasaYeni
        Case "frm_CicerStok": StokYeni
        Case "frm_CicerAlimSatim": AlimSatimYeni
        Case "frm_CicerPersonel": PersonelYeni
    End Select
End Function

' ---------------------------------------------------------------
' Sabit kayıtlar
' ---------------------------------------------------------------
' Hakyağı stoğunun tutulduğu ürün (UrunKodu = HY). Yoksa oluşturur.
Public Function HakYagiUrunID() As Long
    Dim rs As DAO.Recordset
    HakYagiUrunID = Nz(DLookup("UrunID", "tbl_Urun", "UrunKodu='HY'"), 0)
    If HakYagiUrunID = 0 Then
        Set rs = CurrentDb.OpenRecordset("tbl_Urun", dbOpenDynaset)
        rs.AddNew
        rs!UrunKodu = "HY"
        rs!UrunAdi = "Hak Yağı"
        rs!Birim = "Kg"
        rs!Aktif = True
        rs.Update
        rs.Bookmark = rs.LastModified
        HakYagiUrunID = rs!urunID
        rs.Close
    End If
End Function

' Kasa hareketlerinin yazılacağı kasa (ilk aktif kasa). Yoksa "Ana Kasa" oluşturur.
Public Function VarsayilanKasaID() As Long
    Dim rs As DAO.Recordset
    VarsayilanKasaID = Nz(DMin("KasaID", "tbl_Kasa", "Aktif=True AND Nz(KasaTuru,'Nakit')='Nakit'"), 0)
    If VarsayilanKasaID = 0 Then VarsayilanKasaID = Nz(DMin("KasaID", "tbl_Kasa", "Aktif=True"), 0)
    If VarsayilanKasaID = 0 Then VarsayilanKasaID = Nz(DMin("KasaID", "tbl_Kasa"), 0)
    If VarsayilanKasaID = 0 Then
        Set rs = CurrentDb.OpenRecordset("tbl_Kasa", dbOpenDynaset)
        rs.AddNew
        rs!KasaAdi = "Ana Kasa"
        rs!KasaTuru = "Nakit"
        rs!AcilisBakiyesi = 0
        rs!Aktif = True
        rs.Update
        rs.Bookmark = rs.LastModified
        VarsayilanKasaID = rs!KasaID
        rs.Close
    End If
End Function

' ---------------------------------------------------------------
' Bakiyeler
' ---------------------------------------------------------------
' Müşterinin güncel bakiyesi (+ ise müşteri bize borçlu)
Public Function MusteriBakiye(ByVal musteriID As Long) As Currency
    MusteriBakiye = Nz(DLookup("AcilisBakiyesi", "tbl_Musteri", "MusteriID=" & musteriID), 0) + _
        Nz(DSum("Nz(Borc,0)-Nz(Alacak,0)", "tbl_CariHareket", _
            "MusteriID=" & musteriID & " AND (IptalMi=False OR IptalMi Is Null)"), 0)
End Function

' kasaID = 0 ise tüm kasa ve banka hesaplarının toplamı
Public Function KasaBakiye(Optional ByVal kasaID As Long = 0) As Currency
    Dim k As String
    If kasaID <> 0 Then k = " AND KasaID=" & kasaID
    KasaBakiye = Nz(DSum("Nz(AcilisBakiyesi,0)", "tbl_Kasa", "True" & k), 0) + _
        Nz(DSum("Nz(Tahsilat,0)-Nz(Odeme,0)", "tbl_KasaHareket", "Nz(IptalMi,False)=False" & k), 0)
End Function

Public Function UrunStok(ByVal urunID As Long) As Double
    UrunStok = Nz(DLookup("AcilisStok", "tbl_Urun", "UrunID=" & urunID), 0) + _
        Nz(DSum("Nz(Giris,0)-Nz(Cikis,0)", "tbl_StokHareket", _
            "UrunID=" & urunID & " AND (IptalMi=False OR IptalMi Is Null)"), 0)
End Function

' ---------------------------------------------------------------
' Hareket tablolarına yazma
' id = 0 ise yeni kayıt ekler, değilse o kaydı günceller.
' Geriye kaydın numarasını döndürür.
' ---------------------------------------------------------------
Public Function KayitAc(ByVal tablo As String, ByVal anahtar As String, ByVal id As Long) As DAO.Recordset
    Dim rs As DAO.Recordset
    If id = 0 Then
        Set rs = CurrentDb.OpenRecordset(tablo, dbOpenDynaset)
        rs.AddNew
    Else
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & tablo & " WHERE " & anahtar & "=" & id, dbOpenDynaset)
        If rs.EOF Then Err.Raise vbObjectError + 513, , "Kayıt bulunamadı: " & tablo & " / " & id
        rs.Edit
    End If
    Set KayitAc = rs
End Function

Public Function CariYaz(ByVal tarih As Date, ByVal musteriID As Variant, ByVal partiNo As Variant, _
        ByVal hareketTuru As String, ByVal aciklama As String, _
        ByVal borc As Currency, ByVal alacak As Currency, _
        ByVal kaynakTipi As String, ByVal kaynakID As Variant, ByVal otomatik As Boolean, _
        Optional ByVal id As Long = 0) As Long
    Dim rs As DAO.Recordset
    Set rs = KayitAc("tbl_CariHareket", "CariHareketID", id)
    rs!tarih = tarih
    rs!musteriID = musteriID
    rs!partiNo = partiNo
    rs!hareketTuru = Left$(hareketTuru, 40)
    rs!aciklama = MetinVeyaNull(Left$(aciklama, 255))
    rs!borc = borc
    rs!alacak = alacak
    rs!kaynakTipi = kaynakTipi
    rs!kaynakID = kaynakID
    rs!OtomatikMi = otomatik
    rs!IptalMi = False
    rs.Update
    rs.Bookmark = rs.LastModified
    CariYaz = rs!CariHareketID
    rs.Close
End Function

Public Function KasaYaz(ByVal tarih As Date, ByVal musteriID As Variant, ByVal partiNo As Variant, _
        ByVal hareketTuru As String, ByVal odemeYontemi As String, ByVal aciklama As String, _
        ByVal tahsilat As Currency, ByVal odeme As Currency, _
        ByVal kaynakTipi As String, ByVal kaynakID As Variant, ByVal otomatik As Boolean, _
        Optional ByVal id As Long = 0, Optional ByVal kasaID As Long = 0) As Long
    Dim rs As DAO.Recordset
    Set rs = KayitAc("tbl_KasaHareket", "KasaHareketID", id)
    rs!tarih = tarih
    If kasaID = 0 Then kasaID = VarsayilanKasaID()
    rs!kasaID = kasaID
    rs!musteriID = musteriID
    rs!partiNo = partiNo
    rs!hareketTuru = Left$(hareketTuru, 20)
    rs!odemeYontemi = MetinVeyaNull(Left$(odemeYontemi, 30))
    rs!aciklama = MetinVeyaNull(Left$(aciklama, 255))
    rs!tahsilat = tahsilat
    rs!odeme = odeme
    rs!kaynakTipi = kaynakTipi
    rs!kaynakID = kaynakID
    rs!OtomatikMi = otomatik
    rs!IptalMi = False
    rs.Update
    rs.Bookmark = rs.LastModified
    KasaYaz = rs!KasaHareketID
    rs.Close
End Function

Public Function StokYaz(ByVal tarih As Date, ByVal urunID As Long, ByVal musteriID As Variant, ByVal partiNo As Variant, _
        ByVal hareketTuru As String, ByVal aciklama As String, _
        ByVal giris As Double, ByVal cikis As Double, _
        ByVal kaynakTipi As String, ByVal kaynakID As Variant, ByVal otomatik As Boolean, _
        Optional ByVal id As Long = 0) As Long
    Dim rs As DAO.Recordset
    Set rs = KayitAc("tbl_StokHareket", "StokHareketID", id)
    rs!tarih = tarih
    rs!urunID = urunID
    rs!musteriID = musteriID
    rs!partiNo = partiNo
    rs!hareketTuru = Left$(hareketTuru, 30)
    rs!aciklama = MetinVeyaNull(Left$(aciklama, 255))
    rs!giris = giris
    rs!cikis = cikis
    rs!kaynakTipi = kaynakTipi
    rs!kaynakID = kaynakID
    rs!OtomatikMi = otomatik
    rs!IptalMi = False
    rs.Update
    rs.Bookmark = rs.LastModified
    StokYaz = rs!StokHareketID
    rs.Close
End Function

' Bir kaynağın otomatik oluşturduğu cari / kasa / stok kayıtlarını siler.
Public Sub OtomatikKayitlariSil(ByVal kaynakTipi As String, ByVal kaynakID As Long)
    Dim db As DAO.Database, kosul As String
    Set db = CurrentDb
    kosul = " WHERE KaynakTipi=" & SqlMetin(kaynakTipi) & " AND KaynakID=" & kaynakID & " AND OtomatikMi=True"
    db.Execute "DELETE FROM tbl_CariHareket" & kosul, dbFailOnError
    db.Execute "DELETE FROM tbl_KasaHareket" & kosul, dbFailOnError
    db.Execute "DELETE FROM tbl_StokHareket" & kosul, dbFailOnError
End Sub

' Otomatik kayıt kullanıcıya açıklanırken kaynağın adı
Public Function KaynakAdi(ByVal kaynakTipi As Variant) As String
    Select Case Nz(kaynakTipi, "")
        Case KAYNAK_PARTI: KaynakAdi = "Üretim"
        Case KAYNAK_ALIMSATIM: KaynakAdi = "Alım / Satım"
        Case KAYNAK_KASA: KaynakAdi = "Kasa / Banka"
        Case KAYNAK_TESLIM: KaynakAdi = "Yağ Teslim"
        Case Else: KaynakAdi = "Manuel"
    End Select
End Function

' ---------------------------------------------------------------
' SEZON
' Zeytin sezonu 1 Eylül'de başlar, 31 Ağustos'ta biter.
' Sezon adı "2026/2027" biçimindedir. Ayarlar'da aktif sezon seçilir;
' yeni parti numaraları ve listeler aktif sezona göre çalışır.
' ---------------------------------------------------------------
Public Function SezonBul(ByVal tarih As Variant) As String
    Dim y As Integer
    If Not IsDate(tarih) Then tarih = Date
    y = Year(CDate(tarih))
    If Month(CDate(tarih)) < 9 Then y = y - 1
    SezonBul = y & "/" & (y + 1)
End Function

Public Function AktifSezon() As String
    AktifSezon = Ayar("AktifSezon")
    If Len(AktifSezon) <> 9 Then AktifSezon = SezonBul(Date)
End Function

Public Function SezonBaslangic(ByVal sezon As String) As Date
    SezonBaslangic = DateSerial(CInt(Left$(sezon, 4)), 9, 1)
End Function

Public Function SezonBitis(ByVal sezon As String) As Date
    SezonBitis = DateSerial(CInt(Left$(sezon, 4)) + 1, 8, 31)
End Function

' Açılır kutular için sezon listesi (en eski partiden gelecek sezona)
Public Function SezonListesi() As String
    Dim ilk As Integer, son As Integer, y As Integer, s As String, enEski As Variant
    son = CInt(Left$(SezonBul(Date), 4)) + 1
    ilk = son - 2
    enEski = DMin("Tarih", "tbl_Parti")
    If IsDate(enEski) Then
        y = CInt(Left$(SezonBul(enEski), 4))
        If y < ilk Then ilk = y
    End If
    For y = son To ilk Step -1
        s = s & IIf(s = "", "", ";") & y & "/" & (y + 1)
    Next
    SezonListesi = s
End Function

Public Function AktifSezonMetni() As String
    AktifSezonMetni = "Sezon: " & AktifSezon()
End Function

' Sayıyı SQL'e nokta ondalıkla yazar
Public Function SqlSayi(ByVal d As Double) As String
    SqlSayi = Trim$(Str$(d))
End Function

' Ondalık ayarı: "12" / "12,5" / "12.5" -> sayı (geçersizse varsayılan)
Public Function AyarSayi(ByVal anahtar As String, ByVal varsayilan As Double) As Double
    Dim s As String
    s = Trim$(Ayar(anahtar))
    If s = "" Then AyarSayi = varsayilan: Exit Function
    If IsNumeric(s) Then
        AyarSayi = CDbl(s)
    ElseIf IsNumeric(Replace(s, ".", ",")) Then
        AyarSayi = CDbl(Replace(s, ".", ","))
    Else
        AyarSayi = varsayilan
    End If
End Function

' ---------------------------------------------------------------
' Ayarlar (tbl_Ayarlar: Anahtar / Deger)
' Firma, geliştirici ve iletişim bilgileri burada tutulur.
' ---------------------------------------------------------------
Public Function Ayar(ByVal anahtar As String) As String
    ' Veri dosyası henüz bağlanmamışken (ilk açılış, başka bilgisayar) hata vermesin
    On Error Resume Next
    Ayar = Nz(DLookup("Deger", "tbl_Ayarlar", "Anahtar=" & SqlMetin(anahtar)), "")
End Function

Public Sub AyarYaz(ByVal anahtar As String, ByVal deger As Variant)
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Ayarlar WHERE Anahtar=" & SqlMetin(anahtar), dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!anahtar = anahtar
    Else
        rs.Edit
    End If
    rs!deger = MetinVeyaNull(deger)
    rs.Update
    rs.Close
End Sub

' ---------------------------------------------------------------
' Firma adı ve başlıklar
' Firma adı Ayarlar ekranından girilir; başlıklar ve raporlar onu kullanır.
' ---------------------------------------------------------------
Public Function FirmaAdi() As String
    FirmaAdi = Ayar("FirmaAdi")
    If FirmaAdi = "" Then FirmaAdi = "Firma Adı"
End Function

' Türkçe büyük harf: i -> İ, ı -> I
Public Function BuyukHarf(ByVal s As String) As String
    s = Replace(s, "i", "İ", , , vbBinaryCompare)
    s = Replace(s, "ı", "I", , , vbBinaryCompare)
    BuyukHarf = UCase$(s)
End Function

Public Function FirmaAdiBuyuk() As String
    FirmaAdiBuyuk = BuyukHarf(FirmaAdi())
End Function

Public Function ProgramBasligi() As String
    ProgramBasligi = FirmaAdi() & " - " & URUN_ADI
End Function

' Pencere başlığını firma adına göre yeniler
Public Sub BaslikYenile()
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb
    db.Properties("AppTitle") = ProgramBasligi()
    If Err.Number <> 0 Then
        Err.Clear
        db.Properties.Append db.CreateProperty("AppTitle", dbText, ProgramBasligi())
    End If
    Application.RefreshTitleBar
End Sub

Public Function OturumBilgisi() As String
    If OturumID = 0 Then
        OturumBilgisi = ""
    Else
        OturumBilgisi = "Kullanıcı: " & OturumAd & IIf(OturumAdmin, "  (Yönetici)", "")
    End If
End Function

' ---------------------------------------------------------------
' Şifre özeti (tuzlu, 128 bit, tek yönlü)
' Şifreler tabloda açık yazılmaz; "tuz$özet" olarak saklanır.
' ---------------------------------------------------------------
Private Function Mod32(ByVal x As Double) As Double
    Mod32 = x - Int(x / 4294967296#) * 4294967296#
End Function

Private Function Carp32(ByVal a As Double, ByVal b As Double) As Double
    Dim aH As Double, aL As Double
    aH = Int(a / 65536#): aL = a - aH * 65536#
    Carp32 = Mod32(Mod32(aH * b) * 65536# + aL * b)
End Function

Private Function Xor32(ByVal a As Double, ByVal b As Long) As Double
    Dim hi As Long, lo As Long
    hi = CLng(Int(a / 65536#)): lo = CLng(a - hi * 65536#)
    Xor32 = hi * 65536# + (lo Xor (b And &HFFFF&))
End Function

Private Function Hex32(ByVal a As Double) As String
    Dim hi As Long, lo As Long
    hi = CLng(Int(a / 65536#)): lo = CLng(a - hi * 65536#)
    Hex32 = Right$("0000" & Hex$(hi), 4) & Right$("0000" & Hex$(lo), 4)
End Function

Private Function Serit(ByVal metin As String, ByVal baslangic As Double) As Double
    Dim h As Double, i As Long, c As Long
    h = baslangic
    For i = 1 To Len(metin)
        c = AscW(Mid$(metin, i, 1)) And &HFFFF&
        h = Carp32(Xor32(h, c And &HFF&), 16777619#)
        h = Carp32(Xor32(h, (c \ 256) And &HFF&), 16777619#)
    Next
    Serit = h
End Function

Public Function SifreOzeti(ByVal sifre As String, ByVal tuz As String) As String
    Dim t As String, tur As Long, k As Long, sonuc As String
    Dim baslar As Variant
    baslar = Array(2166136261#, 3735928559#, 2654435769#, 1779033703#)
    t = tuz & "|" & sifre
    For tur = 1 To 150
        sonuc = ""
        For k = 0 To 3
            sonuc = sonuc & Hex32(Serit(t & "#" & k & "#" & tur, baslar(k)))
        Next
        t = sonuc & "|" & tuz
    Next
    SifreOzeti = sonuc
End Function

Public Function YeniTuz() As String
    Randomize
    YeniTuz = Right$("0000" & Hex$(Int(Rnd * 65536)), 4) & Right$("0000" & Hex$(Int(Rnd * 65536)), 4)
End Function

Public Function SifreSakla(ByVal sifre As String) As String
    Dim tuz As String
    tuz = YeniTuz()
    SifreSakla = tuz & "$" & SifreOzeti(sifre, tuz)
End Function

Public Function SifreDogruMu(ByVal sifre As String, ByVal saklanan As String) As Boolean
    Dim p As Long
    p = InStr(saklanan, "$")
    If p = 0 Then Exit Function
    SifreDogruMu = (SifreOzeti(sifre, Left$(saklanan, p - 1)) = Mid$(saklanan, p + 1))
End Function