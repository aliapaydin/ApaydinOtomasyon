Option Compare Database
Option Explicit

' =============================================================
' YEDEKLEME, GERİ YÜKLEME VE ÇIKIŞ
' -------------------------------------------------------------
' Formlar:
'   frm_Yedek      yedek alma (Ayarlar'dan ya da çıkışta "Yedek Alıp Kapat")
'   frm_GeriYukle  yedekten geri yükleme (Ayarlar, yalnızca yönetici)
'   frm_Cikis      çıkış sorusu: Evet / Yedek Alıp Kapat / Hayır
'   frm_Kapanis    gizli form; Access penceresinin X düğmesiyle
'                  kapatılırken de çıkış sorusunu gösterir
'
' Yedekler seçilen sürücüde ya da klasörde "ApaydinOtomasyon_Yedek"
' klasörüne ApaydinVeri_yyyy-aa-gg_ss-dd-nn.accdb adıyla yazılır.
' Eski yedekler silinmez. Veri dosyası açıkken de güvenle kopyalanır:
' önce geçici klasöre kopyalanır, sıkıştırılarak doğrulanır, sonra
' hedefe yazılır.
'
' Geri yükleme: seçilen yedek doğrulanır ve veri dosyasının yanına
' "ApaydinVeri_geriyukle.accdb" olarak kopyalanır, program yeniden
' başlatılır. Açılışta (veri tablolarına dokunulmadan önce) mevcut
' veri bu bilgisayardaki Yedek klasörüne taşınır ve yedek yerine konur.
' =============================================================

Public Const YEDEK_KLASOR_ADI As String = "ApaydinOtomasyon_Yedek"
Private Const BEKLEYEN_AD As String = "ApaydinVeri_geriyukle.accdb"
Private Const SIKISTIRMA_AD As String = "ApaydinVeri_sikistirilmis.accdb"

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)

Public CikisOnay As Boolean          ' True ise program soru sormadan kapanır
Private CikisSecim As Long           ' 0 hayır, 1 evet, 2 yedek alıp kapat
Private YedekKapanis As Boolean      ' frm_Yedek çıkış için açıldı
Private YedekAlindi As Boolean
Private SeciliKlasor As String       ' yedek ekranında elle seçilen klasör
Private EkKlasorler As Collection    ' geri yüklemede elle seçilen klasörler
Private OneriYol As String
Private OneriTarih As Date
Private BulunanYollar() As String
Private BulunanTarihler() As Date
Private BulunanKonumlar() As String
Private BulunanSayi As Long

' ---------------------------------------------------------------
' Yardımcılar
' ---------------------------------------------------------------
Private Function Fso() As Object
    Set Fso = CreateObject("Scripting.FileSystemObject")
End Function

Private Function YolBirlestir(ByVal a As String, ByVal b As String) As String
    If Right$(a, 1) = "\" Then
        YolBirlestir = a & b
    Else
        YolBirlestir = a & "\" & b
    End If
End Function

Private Function SondakiTersBolu(ByVal s As String) As String
    If Len(s) > 3 And Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
    SondakiTersBolu = s
End Function

' Seçilen kök (sürücü ya da klasör) için yedek klasörü
Public Function YedekKlasoru(ByVal kok As String) As String
    kok = SondakiTersBolu(kok)
    If LCase$(Fso().GetFileName(kok)) = LCase$(YEDEK_KLASOR_ADI) Then
        YedekKlasoru = kok
    Else
        YedekKlasoru = YolBirlestir(kok, YEDEK_KLASOR_ADI)
    End If
End Function

' Bu bilgisayardaki otomatik yedek klasörü: ...\Veri\ApaydinVeri.accdb -> ...\Yedek
Public Function YerelYedekKlasoru() As String
    Dim veri As String
    veri = VeriYolu()
    If veri = "" Then veri = VarsayilanVeriYolu()
    YerelYedekKlasoru = Fso().GetParentFolderName(Fso().GetParentFolderName(veri)) & "\Yedek"
End Function

Private Function VL(ByVal s As String) As String
    VL = """" & Replace(s, """", "'") & """"
End Function

Private Function SurucuTuru(ByVal t As Long) As String
    Select Case t
        Case 1: SurucuTuru = "USB / Flaş"
        Case 2: SurucuTuru = "Yerel disk"
        Case 3: SurucuTuru = "Ağ sürücüsü"
        Case Else: SurucuTuru = "Diğer"
    End Select
End Function

Private Function BoyutMetni(ByVal b As Double) As String
    If b >= 1073741824# Then
        BoyutMetni = Format$(b / 1073741824#, "0.0") & " GB"
    ElseIf b >= 1048576# Then
        BoyutMetni = Format$(b / 1048576#, "0.0") & " MB"
    Else
        BoyutMetni = Format$(b / 1024#, "0") & " KB"
    End If
End Function

Private Function TarihMetni(ByVal d As Date) As String
    TarihMetni = Format$(d, "dd\.mm\.yyyy hh:nn")
End Function

Private Function TarihSaniye(ByVal d As Date) As String
    TarihSaniye = Format$(d, "dd\.mm\.yyyy hh:nn:ss")
End Function

Private Function YedekDosyasiMi(ByVal ad As String) As Boolean
    ad = LCase$(ad)
    YedekDosyasiMi = (ad Like "apaydinveri*.accdb") And ad <> LCase$(BEKLEYEN_AD)
End Function

Private Function YedekSayisi(ByVal klasor As String) As Long
    On Error Resume Next
    Dim f As Object, n As Long
    If Not Fso().FolderExists(klasor) Then Exit Function
    For Each f In Fso().GetFolder(klasor).Files
        If YedekDosyasiMi(f.Name) Then n = n + 1
    Next
    YedekSayisi = n
End Function

' ApaydinVeri_2026-11-17_13-10-05.accdb -> tarih; ad bu biçimde değilse dosya tarihi
Private Function YedekTarihi(ByVal ad As String, ByVal varsayilan As Date) As Date
    Dim s As String
    s = Mid$(ad, 13, 19)
    If s Like "####-##-##_##-##-##" Then
        YedekTarihi = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2))) + _
                      TimeSerial(CInt(Mid$(s, 12, 2)), CInt(Mid$(s, 15, 2)), CInt(Mid$(s, 18, 2)))
    ElseIf Left$(s, 15) Like "########_######" Then
        ' Kurulum yedekleri: ApaydinVeri_20261117_131506.accdb
        YedekTarihi = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 5, 2)), CInt(Mid$(s, 7, 2))) + _
                      TimeSerial(CInt(Mid$(s, 10, 2)), CInt(Mid$(s, 12, 2)), CInt(Mid$(s, 14, 2)))
    Else
        YedekTarihi = varsayilan
    End If
End Function

' "yyyy-mm-dd hh:nn:ss" -> tarih
Private Function IsoTarih(ByVal s As String) As Date
    On Error Resume Next
    IsoTarih = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2))) + _
               TimeSerial(CInt(Mid$(s, 12, 2)), CInt(Mid$(s, 15, 2)), CInt(Mid$(s, 18, 2)))
End Function

Private Sub SilVarsa(ByVal yol As String)
    On Error Resume Next
    If Fso().FileExists(yol) Then Fso().DeleteFile yol, True
End Sub

Private Function KlasorSecPenceresi(ByVal baslik As String) As String
    On Error GoTo Hata
    Dim fd As Object
    Set fd = Application.FileDialog(4)   ' msoFileDialogFolderPicker
    fd.Title = baslik
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then KlasorSecPenceresi = fd.SelectedItems(1)
    Exit Function
Hata:
    Mesaj "Klasör seçme penceresi açılamadı: " & Err.Description, vbExclamation
End Function

' Açık (kullanımdaki) dosyayı paylaşımlı okuyarak kopyalar
Private Sub PaylasimliKopyala(ByVal kaynak As String, ByVal hedef As String)
    Dim f1 As Integer, f2 As Integer, kalan As Double, parca As Long, buf() As Byte
    f1 = FreeFile
    Open kaynak For Binary Access Read Shared As #f1
    f2 = FreeFile
    Open hedef For Binary Access Write Lock Read Write As #f2
    kalan = LOF(f1)
    Do While kalan > 0
        If kalan > 1048576 Then parca = 1048576 Else parca = CLng(kalan)
        ReDim buf(0 To parca - 1)
        Get #f1, , buf
        Put #f2, , buf
        kalan = kalan - parca
    Loop
    Close #f2
    Close #f1
End Sub

' Dosya geçerli bir veri dosyası mı? Geçerliyse veri sürümünü döndürür.
Private Function YedekBilgisi(ByVal yol As String, ByRef surum As Long) As Boolean
    On Error GoTo Hata
    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = DBEngine.OpenDatabase(yol, False, True)
    Set rs = db.OpenRecordset("SELECT Deger FROM tbl_Ayarlar WHERE Anahtar='VeriSurum'", dbOpenSnapshot)
    If rs.EOF Then
        surum = 1
    Else
        surum = CLng(Val(Nz(rs!deger, "1")))
    End If
    rs.Close
    Set rs = db.OpenRecordset("SELECT TOP 1 KullaniciID FROM tbl_Kullanici", dbOpenSnapshot)
    rs.Close
    db.Close
    YedekBilgisi = True
    Exit Function
Hata:
    On Error Resume Next
    If Not db Is Nothing Then db.Close
End Function

Public Function SonYedekMetni() As String
    On Error Resume Next
    Dim t As String, d As Date, gun As Long, ek As String
    t = Ayar("SonYedekTarihi")
    If t = "" Then
        SonYedekMetni = "Henüz yedek alınmadı."
        Exit Function
    End If
    d = IsoTarih(t)
    gun = DateDiff("d", d, Date)
    If gun <= 0 Then
        ek = "bugün"
    ElseIf gun = 1 Then
        ek = "dün"
    Else
        ek = gun & " gün önce"
    End If
    SonYedekMetni = "Son yedek: " & TarihMetni(d) & " (" & ek & ")   " & Ayar("SonYedekYolu")
End Function

' ---------------------------------------------------------------
' Çıkış
' ---------------------------------------------------------------
Public Function CikisOnaylandi() As Boolean
    CikisOnaylandi = CikisOnay
End Function

' Ana menü / giriş ekranındaki Çıkış düğmesi ve Access'in X düğmesi buraya gelir
Public Function CikisSor() As Boolean
    On Error GoTo Hata
    If TestModu Then
        ProgramiKapat
        Exit Function
    End If
    CikisSecim = 0
    DoCmd.OpenForm "frm_Cikis", WindowMode:=acDialog
    Select Case CikisSecim
        Case 1
            ProgramiKapat
        Case 2
            YedekAlindi = False
            DoCmd.OpenForm "frm_Yedek", WindowMode:=acDialog, OpenArgs:="kapanis"
            If YedekAlindi Then
                ProgramiKapat
            Else
                EkranlariGeriAc
            End If
        Case Else
            EkranlariGeriAc
    End Select
    Exit Function
Hata:
    Mesaj "Çıkış sırasında hata: " & Err.Description, vbExclamation
End Function

Private Sub ProgramiKapat()
    CikisOnay = True
    DoCmd.Quit acQuitSaveAll
End Sub

' Access penceresi kapatılmak istenip vazgeçildiyse, kapanmış ekranları geri açar
Private Sub EkranlariGeriAc()
    On Error Resume Next
    Dim i As Long, gorunen As Boolean
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> "frm_Kapanis" And Forms(i).Visible Then gorunen = True
    Next
    If gorunen Then Exit Sub
    If OturumID > 0 Then
        DoCmd.OpenForm ANA_MENU
    Else
        DoCmd.OpenForm GIRIS_FORMU
    End If
End Sub

' Giriş ekranı açılırken çağrılır: X ile kapatmayı yakalayan gizli form
Public Sub KapanisFormunuAc()
    On Error Resume Next
    If TestModu Then Exit Sub
    CikisOnay = False
    If Not FormAcikMi("frm_Kapanis") Then DoCmd.OpenForm "frm_Kapanis", WindowMode:=acHidden
End Sub

Public Function CikisFormYukle()
    On Error Resume Next
    Forms("frm_Cikis")!lblSonYedek.Caption = SonYedekMetni()
End Function

Public Function CikisEvet()
    CikisSecim = 1
    DoCmd.Close acForm, "frm_Cikis"
End Function

Public Function CikisYedekle()
    CikisSecim = 2
    DoCmd.Close acForm, "frm_Cikis"
End Function

Public Function CikisHayir()
    CikisSecim = 0
    DoCmd.Close acForm, "frm_Cikis"
End Function

' ---------------------------------------------------------------
' YEDEK ALMA   (form: frm_Yedek)
' ---------------------------------------------------------------
Private Function Y() As Form
    Set Y = Forms("frm_Yedek")
End Function

Public Function YedekEkraniAc()
    On Error Resume Next
    DoCmd.OpenForm "frm_Yedek", WindowMode:=acDialog
    If FormAcikMi("frm_Ayarlar") Then Forms("frm_Ayarlar")!lblSonYedek.Caption = SonYedekMetni()
End Function

Public Function YedekFormYukle()
    On Error Resume Next
    Dim frm As Form
    Set frm = Y()
    YedekKapanis = (Nz(frm.OpenArgs, "") = "kapanis")
    YedekAlindi = False
    If YedekKapanis Then
        frm!btnYedekle.Caption = "Yedeği Al ve Kapat"
        frm.Caption = "Yedek Alıp Kapat"
    End If
    frm!lblSonYedek.Caption = SonYedekMetni()
    SuruculeriListele
End Function

Public Function SuruculeriListele()
    On Error GoTo Hata
    Dim frm As Form, d As Object, satir As String, usb As String, diger As String, ilkUsb As String
    Dim ad As String, kok As String, secili As String, son As String, ekler As String, sonSecim As String
    DoCmd.Hourglass True
    Set frm = Y()
    secili = Nz(frm!lstSuruculer, "")
    For Each d In Fso().Drives
        If d.DriveType <> 4 And d.DriveType <> 0 Then
            If d.IsReady Then
                kok = d.DriveLetter & ":\"
                ad = ""
                ad = d.VolumeName
                satir = VL(kok) & ";" & VL(d.DriveLetter & ":") & ";" & VL(ad) & ";" & _
                        VL(SurucuTuru(d.DriveType)) & ";" & VL(BoyutMetni(d.FreeSpace)) & ";" & _
                        VL(CStr(YedekSayisi(YedekKlasoru(kok))))
                If d.DriveType = 1 Then
                    usb = usb & ";" & satir
                    If ilkUsb = "" Then ilkUsb = kok
                Else
                    diger = diger & ";" & satir
                End If
            End If
        End If
    Next
    ' Bu sefer seçilen klasör ve daha önce yedek alınan klasör (sürücü kökü değilse)
    If SeciliKlasor <> "" Then ekler = ";" & KlasorSatiri(SeciliKlasor)
    son = Ayar("SonYedekYolu")
    If son <> "" Then
        son = Fso().GetParentFolderName(son)
        If LCase$(Fso().GetFileName(son)) = LCase$(YEDEK_KLASOR_ADI) Then son = Fso().GetParentFolderName(son)
        son = SondakiTersBolu(son)
        If Len(son) <= 3 Then sonSecim = Left$(son, 2) & "\" Else sonSecim = son
        If Not Fso().FolderExists(sonSecim) Then sonSecim = ""
    End If
    If Len(son) > 3 And LCase$(son) <> LCase$(SeciliKlasor) Then
        If Fso().FolderExists(son) Then ekler = ekler & ";" & KlasorSatiri(son)
    End If
    frm!lstSuruculer.RowSource = VL("Yol") & ";" & VL("Sürücü") & ";" & VL("Ad") & ";" & VL("Tür") & ";" & _
                                 VL("Boş Alan") & ";" & VL("Yedek") & usb & diger & ekler
    frm!lstSuruculer.Requery
    If SeciliKlasor <> "" Then
        frm!lstSuruculer = SeciliKlasor
    ElseIf secili <> "" Then
        frm!lstSuruculer = secili
    ElseIf ilkUsb <> "" Then
        frm!lstSuruculer = ilkUsb
    ElseIf sonSecim <> "" Then
        frm!lstSuruculer = sonSecim    ' flaş takılı değilse en son kullanılan yer
    End If
    If ilkUsb = "" Then
        frm!lblUyari.Caption = "Takılı flaş bellek bulunamadı. Flaşı takıp 'Sürücüleri Yenile'ye basın ya da bir klasör seçin."
    Else
        frm!lblUyari.Caption = ""
    End If
    HedefGoster
    DoCmd.Hourglass False
    Exit Function
Hata:
    DoCmd.Hourglass False
    Mesaj "Sürücüler listelenemedi: " & Err.Description, vbExclamation
End Function

Private Function KlasorSatiri(ByVal k As String) As String
    Dim ad As String
    ad = Fso().GetFileName(SondakiTersBolu(k))
    KlasorSatiri = VL(k) & ";" & VL(Left$(k, 2)) & ";" & VL(ad) & ";" & VL("Klasör") & ";" & VL("") & ";" & _
                   VL(CStr(YedekSayisi(YedekKlasoru(k))))
End Function

Public Function HedefGoster()
    On Error Resume Next
    Dim frm As Form, kok As String, metin As String
    Set frm = Y()
    kok = Nz(frm!lstSuruculer, "")
    If kok = "" Then
        frm!lblHedef.Caption = "Listeden bir sürücü ya da klasör seçin."
        Exit Function
    End If
    metin = "Yedek şuraya yazılacak:  " & YolBirlestir(YedekKlasoru(kok), "ApaydinVeri_" & Format$(Now, "yyyy-mm-dd_hh-nn-ss") & ".accdb")
    If UCase$(Left$(kok, 1)) = UCase$(Left$(VeriYolu(), 1)) Then
        metin = metin & vbCrLf & "Not: Bu sürücü veri dosyasıyla aynı diskte. Bilgisayar bozulursa yedek de kaybolabilir; flaş bellek önerilir."
    End If
    frm!lblHedef.Caption = metin
End Function

Public Function YedekKlasorSec()
    Dim k As String
    k = KlasorSecPenceresi("Yedeğin alınacağı klasörü seçin")
    If k = "" Then Exit Function
    SeciliKlasor = SondakiTersBolu(k)
    SuruculeriListele
End Function

Public Function YedekAl()
    On Error GoTo Hata
    Dim frm As Form, kok As String, hedefKlasor As String, hedef As String, kaynak As String
    Dim tmp1 As String, tmp2 As String, surum As Long
    Set frm = Y()
    kok = Nz(frm!lstSuruculer, "")
    If kok = "" Then
        Mesaj "Yedeğin alınacağı sürücüyü ya da klasörü seçin.", vbExclamation
        Exit Function
    End If
    kaynak = VeriYolu()
    If Not Fso().FileExists(kaynak) Then
        Mesaj "Veri dosyası bulunamadı:" & vbCrLf & kaynak, vbExclamation
        Exit Function
    End If
    hedefKlasor = YedekKlasoru(kok)
    hedef = YolBirlestir(hedefKlasor, "ApaydinVeri_" & Format$(Now, "yyyy-mm-dd_hh-nn-ss") & ".accdb")
    DoCmd.Hourglass True
    tmp1 = Environ$("TEMP") & "\apaydin_yedek_1.accdb"
    tmp2 = Environ$("TEMP") & "\apaydin_yedek_2.accdb"
    SilVarsa tmp1
    SilVarsa tmp2
    PaylasimliKopyala kaynak, tmp1
    DBEngine.CompactDatabase tmp1, tmp2
    If Not YedekBilgisi(tmp2, surum) Then Err.Raise vbObjectError + 513, , "Yedek dosyası doğrulanamadı."
    If Not Fso().FolderExists(hedefKlasor) Then Fso().CreateFolder hedefKlasor
    Fso().CopyFile tmp2, hedef, False
    SilVarsa tmp1
    SilVarsa tmp2
    DoCmd.Hourglass False
    AyarYaz "SonYedekTarihi", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    AyarYaz "SonYedekYolu", hedef
    YedekAlindi = True
    Mesaj "Yedek alındı:" & vbCrLf & hedef & vbCrLf & "(" & BoyutMetni(FileLen(hedef)) & ")" & _
          IIf(YedekKapanis, vbCrLf & vbCrLf & "Program kapatılacak.", "")
    DoCmd.Close acForm, "frm_Yedek"
    Exit Function
Hata:
    Close
    DoCmd.Hourglass False
    SilVarsa tmp1
    SilVarsa tmp2
    Mesaj "Yedek alınamadı: " & Err.Description, vbExclamation
End Function

Public Function YedekVazgec()
    YedekAlindi = False
    DoCmd.Close acForm, "frm_Yedek"
End Function

' ---------------------------------------------------------------
' GERİ YÜKLEME   (form: frm_GeriYukle)
' ---------------------------------------------------------------
Private Function G() As Form
    Set G = Forms("frm_GeriYukle")
End Function

Public Function GeriYukleEkraniAc()
    If Not OturumAdmin And Not TestModu Then
        Mesaj "Yedekten geri yüklemeyi yalnızca yönetici yapabilir.", vbExclamation
        Exit Function
    End If
    DoCmd.OpenForm "frm_GeriYukle", WindowMode:=acDialog
End Function

Public Function GeriYukleFormYukle()
    Set EkKlasorler = New Collection
    YedekleriTara
End Function

Public Function GeriYukleKlasorSec()
    Dim k As String
    k = KlasorSecPenceresi("Yedeklerin bulunduğu klasörü seçin")
    If k = "" Then Exit Function
    If EkKlasorler Is Nothing Then Set EkKlasorler = New Collection
    EkKlasorler.Add SondakiTersBolu(k)
    YedekleriTara
End Function

Public Function YedekleriTara()
    On Error GoTo Hata
    Dim frm As Form, d As Object, f As Object, k As Variant, klasorler As Collection
    Dim gorulen As String, n As Long, i As Long, j As Long, son As String
    Dim yollar() As String, tarihler() As Date, konumlar() As String, otomatik() As Boolean, boyutlar() As String
    Dim sira() As Long, tmp As Long, oneri As Long, liste As String, ad As String

    DoCmd.Hourglass True
    Set frm = G()
    Set klasorler = New Collection
    For Each d In Fso().Drives
        If d.DriveType <> 4 And d.DriveType <> 0 Then
            If d.IsReady Then
                ad = ""
                ad = d.VolumeName
                klasorler.Add Array(YedekKlasoru(d.DriveLetter & ":\"), d.DriveLetter & ": " & SurucuTuru(d.DriveType) & _
                                    IIf(ad = "", "", " (" & ad & ")"), False)
            End If
        End If
    Next
    son = Ayar("SonYedekYolu")
    If son <> "" Then klasorler.Add Array(Fso().GetParentFolderName(son), "Son yedek klasörü", False)
    If Not EkKlasorler Is Nothing Then
        For Each k In EkKlasorler
            klasorler.Add Array(CStr(k), "Seçilen klasör", False)
            klasorler.Add Array(YedekKlasoru(CStr(k)), "Seçilen klasör", False)
        Next
    End If
    klasorler.Add Array(YerelYedekKlasoru(), "Bu bilgisayar (otomatik)", True)

    n = 0
    gorulen = "|"
    For Each k In klasorler
        If Fso().FolderExists(k(0)) Then
            For Each f In Fso().GetFolder(k(0)).Files
                If YedekDosyasiMi(f.Name) And InStr(1, gorulen, "|" & LCase$(f.Path) & "|") = 0 Then
                    gorulen = gorulen & LCase$(f.Path) & "|"
                    ReDim Preserve yollar(n): ReDim Preserve tarihler(n): ReDim Preserve konumlar(n)
                    ReDim Preserve otomatik(n): ReDim Preserve boyutlar(n)
                    yollar(n) = f.Path
                    tarihler(n) = YedekTarihi(f.Name, f.DateLastModified)
                    konumlar(n) = k(1)
                    otomatik(n) = k(2)
                    boyutlar(n) = BoyutMetni(f.Size)
                    n = n + 1
                End If
            Next
        End If
    Next

    ' Tarihe göre yeniden eskiye sırala
    BulunanSayi = n
    OneriYol = ""
    If n > 0 Then
        ReDim sira(n - 1)
        For i = 0 To n - 1: sira(i) = i: Next
        For i = 1 To n - 1
            tmp = sira(i)
            j = i - 1
            Do While j >= 0
                If tarihler(sira(j)) >= tarihler(tmp) Then Exit Do
                sira(j + 1) = sira(j)
                j = j - 1
            Loop
            sira(j + 1) = tmp
        Next
        ReDim BulunanYollar(n - 1): ReDim BulunanTarihler(n - 1): ReDim BulunanKonumlar(n - 1)
        For i = 0 To n - 1
            BulunanYollar(i) = yollar(sira(i))
            BulunanTarihler(i) = tarihler(sira(i))
            BulunanKonumlar(i) = konumlar(sira(i))
        Next
        ' Öneri: flaş / klasördeki en güncel yedek; yoksa en güncel otomatik kopya
        oneri = -1
        For i = 0 To n - 1
            If Not otomatik(sira(i)) Then oneri = sira(i): Exit For
        Next
        If oneri = -1 Then oneri = sira(0)
        OneriYol = yollar(oneri)
        OneriTarih = tarihler(oneri)
    End If

    liste = VL("Yol") & ";" & VL("Tarih") & ";" & VL("Konum") & ";" & VL("Dosya") & ";" & VL("Boyut")
    For i = 0 To n - 1
        If i > 150 Then Exit For   ' liste kutusu sınırı; en eski yedekler gösterilmez
        If yollar(sira(i)) <> OneriYol Then
            liste = liste & ";" & VL(yollar(sira(i))) & ";" & VL(TarihMetni(tarihler(sira(i)))) & ";" & _
                    VL(konumlar(sira(i))) & ";" & VL(Fso().GetFileName(yollar(sira(i)))) & ";" & VL(boyutlar(sira(i)))
        End If
    Next
    frm!lstYedekler.RowSource = liste
    frm!lstYedekler.Requery
    frm!lstYedekler = Null

    If OneriYol = "" Then
        frm!lblOneri.Caption = "Yedek bulunamadı."
        frm!lblOneriYol.Caption = "Flaşı takıp 'Yeniden Tara'ya basın ya da 'Klasör Seç' ile yedeklerin bulunduğu klasörü gösterin."
        frm!btnOneriYukle.Enabled = False
    Else
        frm!lblOneri.Caption = TarihMetni(OneriTarih) & "   •   " & konumlar(oneri)
        frm!lblOneriYol.Caption = OneriYol & "   (" & boyutlar(oneri) & ")"
        frm!btnOneriYukle.Enabled = True
    End If
    frm!lblSayi.Caption = "DİĞER YEDEKLER  (" & IIf(n > 1, n - 1, 0) & ")"
    DoCmd.Hourglass False
    Exit Function
Hata:
    DoCmd.Hourglass False
    Mesaj "Yedekler taranamadı: " & Err.Description, vbExclamation
End Function

Private Function BulunanTarih(ByVal yol As String) As Date
    Dim i As Long
    For i = 0 To BulunanSayi - 1
        If BulunanYollar(i) = yol Then BulunanTarih = BulunanTarihler(i): Exit Function
    Next
    BulunanTarih = FileDateTime(yol)
End Function

Private Function BulunanKonum(ByVal yol As String) As String
    Dim i As Long
    For i = 0 To BulunanSayi - 1
        If BulunanYollar(i) = yol Then BulunanKonum = BulunanKonumlar(i): Exit Function
    Next
End Function

Public Function OneriYukle()
    If OneriYol = "" Then Exit Function
    YedegiYukle OneriYol
End Function

Public Function SeciliYukle()
    Dim yol As String
    yol = Nz(G()!lstYedekler, "")
    If yol = "" Then
        Mesaj "Diğer yedekler listesinden bir yedek seçin.", vbExclamation
        Exit Function
    End If
    YedegiYukle yol
End Function

Private Sub YedegiYukle(ByVal yol As String)
    On Error GoTo Hata
    Dim t As Date, surum As Long, mevcut As String, mevcutT As Date, metin As String, hedef As String, eski As Boolean

    If Not YedekBilgisi(yol, surum) Then
        Mesaj "Seçilen dosya geçerli bir veri dosyası değil ya da bozuk:" & vbCrLf & yol, vbCritical
        Exit Sub
    End If
    If surum > VERI_SURUM Then
        Mesaj "Bu yedek programın daha yeni bir sürümüyle alınmış (veri sürümü " & surum & ")." & vbCrLf & _
              "Önce programı güncelleyin.", vbExclamation
        Exit Sub
    End If

    t = BulunanTarih(yol)
    mevcut = VeriYolu()
    eski = (yol <> OneriYol And t < OneriTarih)
    If eski Then
        metin = "DİKKAT: ESKİ BİR YEDEK YÜKLÜYORSUNUZ." & vbCrLf & vbCrLf & _
                "Seçilen yedek:   " & TarihSaniye(t) & "   " & BulunanKonum(yol) & vbCrLf & _
                "En güncel yedek: " & TarihSaniye(OneriTarih) & "   " & BulunanKonum(OneriYol) & vbCrLf & vbCrLf & _
                "Seçilen yedekten sonra girilen tüm kayıtlar kaybolur."
    Else
        metin = "Seçilen yedek yüklenecek:" & vbCrLf & TarihSaniye(t) & "   " & BulunanKonum(yol)
    End If
    If Fso().FileExists(mevcut) Then
        mevcutT = FileDateTime(mevcut)
        If mevcutT > DateAdd("n", 2, t) Then
            metin = metin & vbCrLf & vbCrLf & "Bu bilgisayardaki veriler en son " & TarihMetni(mevcutT) & _
                    " tarihinde değişmiş, yani yedekten daha yeni."
        End If
        metin = metin & vbCrLf & vbCrLf & "Mevcut veri silinmez; bu bilgisayardaki Yedek klasörüne taşınır."
    End If
    metin = metin & vbCrLf & vbCrLf & "Devam edilsin mi? (Program yeniden başlatılacak.)"
    If Mesaj(metin, vbYesNo + IIf(eski, vbExclamation, vbQuestion) + vbDefaultButton2) <> vbYes Then Exit Sub

    hedef = YolBirlestir(Fso().GetParentFolderName(mevcut), BEKLEYEN_AD)
    SilVarsa hedef
    DoCmd.Hourglass True
    Fso().CopyFile yol, hedef, True
    DoCmd.Hourglass False
    YenidenBaslat
    Exit Sub
Hata:
    DoCmd.Hourglass False
    Mesaj "Yedek yüklenemedi: " & Err.Description, vbExclamation
End Sub

Private Sub YenidenBaslat()
    ProgramiYenidenBaslat "Program kapanacak. Lütfen programı yeniden açın; yedek açılışta yüklenecek."
End Sub

Public Sub ProgramiYenidenBaslat(Optional ByVal hataMesaji As String = "")
    On Error Resume Next
    If hataMesaji = "" Then hataMesaji = "Program kapanacak. Lütfen programı yeniden açın."
    Dim exe As String
    exe = SysCmd(acSysCmdAccessDir) & "MSACCESS.EXE"
    Shell """" & exe & """ """ & CurrentProject.FullName & """", vbNormalFocus
    If Err.Number <> 0 Then Mesaj hataMesaji
    CikisOnay = True
    DoCmd.Quit acQuitSaveNone
End Sub

' ---------------------------------------------------------------
' Cari silme (modCariYonetim) için yedek ve sıkıştırma
' ---------------------------------------------------------------
' Veri dosyasının sıkıştırılmış ve doğrulanmış kopyasını hedefe yazar (dosya açıkken de güvenli)
Private Sub VeriKopyasiYaz(ByVal hedef As String)
    Dim kaynak As String, tmp1 As String, tmp2 As String, surum As Long
    kaynak = VeriYolu()
    If Not Fso().FileExists(kaynak) Then Err.Raise vbObjectError + 514, , "Veri dosyası bulunamadı: " & kaynak
    tmp1 = Environ$("TEMP") & "\apaydin_kopya_1.accdb"
    tmp2 = Environ$("TEMP") & "\apaydin_kopya_2.accdb"
    SilVarsa tmp1
    SilVarsa tmp2
    PaylasimliKopyala kaynak, tmp1
    DBEngine.CompactDatabase tmp1, tmp2
    SilVarsa tmp1
    If Not YedekBilgisi(tmp2, surum) Then
        SilVarsa tmp2
        Err.Raise vbObjectError + 513, , "Veri kopyası doğrulanamadı."
    End If
    SilVarsa hedef
    Fso().CopyFile tmp2, hedef, True
    SilVarsa tmp2
End Sub

' Silmeden önce bu bilgisayardaki Yedek klasörüne yedek alır, yolunu döndürür
Public Function CariSilmeYedegiAl() As String
    Dim hedef As String
    If Not Fso().FolderExists(YerelYedekKlasoru()) Then Fso().CreateFolder YerelYedekKlasoru()
    hedef = YolBirlestir(YerelYedekKlasoru(), "ApaydinVeri_cari_silme_oncesi_" & Format$(Now, "yyyy-mm-dd_hh-nn-ss") & ".accdb")
    VeriKopyasiYaz hedef
    CariSilmeYedegiAl = hedef
End Function

' Silinen kayıtların izi dosyada kalmasın diye sıkıştırılmış kopya hazırlanır;
' program yeniden açılırken (veri tablolarına dokunulmadan önce) eski dosyanın yerine geçer.
Public Sub VeriSikistirmaHazirla()
    VeriKopyasiYaz YolBirlestir(Fso().GetParentFolderName(VeriYolu()), SIKISTIRMA_AD)
End Sub

' Açılışta BekleyenGeriYukleme'den hemen sonra çağrılır
Public Sub BekleyenSikistirma()
    On Error GoTo Hata
    Dim veri As String, bekleyen As String, eski As String, i As Long, tasindi As Boolean
    veri = VeriYolu()
    If veri = "" Then Exit Sub
    bekleyen = YolBirlestir(Fso().GetParentFolderName(veri), SIKISTIRMA_AD)
    If Not Fso().FileExists(bekleyen) Then Exit Sub
    If Not Fso().FileExists(veri) Then SilVarsa bekleyen: Exit Sub
    ' Kopya hazırlandıktan sonra veri değiştiyse (program yeniden başlatılmadan kullanıldıysa) kopyayı kullanma
    If FileDateTime(veri) > DateAdd("n", 3, FileDateTime(bekleyen)) Then SilVarsa bekleyen: Exit Sub
    eski = YolBirlestir(Fso().GetParentFolderName(veri), "ApaydinVeri_sikistirma_oncesi.accdb")
    SilVarsa eski
    ' Eski program kapanırken dosya bir süre kullanımda kalabilir: 20 saniyeye kadar dene
    For i = 1 To 40
        On Error Resume Next
        Err.Clear
        Name veri As eski
        If Err.Number = 0 Then tasindi = True
        On Error GoTo Hata
        If tasindi Then Exit For
        Sleep 500
    Next
    If Not tasindi Then Exit Sub        ' bir sonraki açılışta tekrar denenir
    Name bekleyen As veri
    SilVarsa eski
    Exit Sub
Hata:
    On Error Resume Next
    If tasindi And Not Fso().FileExists(veri) Then Name eski As veri
End Sub

Public Function GeriYukleKapat()
    DoCmd.Close acForm, "frm_GeriYukle"
End Function

' Açılışta, veri tablolarına dokunulmadan önce çağrılır.
' Bekleyen bir geri yükleme varsa mevcut veriyi Yedek klasörüne taşır, yedeği yerine koyar.
Public Sub BekleyenGeriYukleme()
    On Error GoTo Hata
    Dim veri As String, bekleyen As String, eski As String, i As Long, tamam As Boolean, tasindi As Boolean
    veri = VeriYolu()
    If veri = "" Then veri = VarsayilanVeriYolu()
    bekleyen = YolBirlestir(Fso().GetParentFolderName(veri), BEKLEYEN_AD)
    If Not Fso().FileExists(bekleyen) Then Exit Sub
    If Not Fso().FolderExists(YerelYedekKlasoru()) Then Fso().CreateFolder YerelYedekKlasoru()
    eski = YolBirlestir(YerelYedekKlasoru(), "ApaydinVeri_geriyukleme_oncesi_" & Format$(Now, "yyyy-mm-dd_hh-nn-ss") & ".accdb")
    ' Eski program kapanırken dosya bir süre kullanımda kalabilir: 20 saniyeye kadar dene
    For i = 1 To 40
        If Not Fso().FileExists(veri) Then
            tamam = True
            Exit For
        End If
        On Error Resume Next
        Err.Clear
        Name veri As eski
        If Err.Number = 0 Then
            tamam = True
            tasindi = True
            Exit For
        End If
        On Error GoTo Hata
        Sleep 500
    Next
    On Error GoTo Hata
    If Not tamam Then
        Mesaj "Veri dosyası kullanımda olduğu için yedek yüklenemedi." & vbCrLf & _
              "Programı kapatıp yeniden açın; yükleme tekrar denenecek.", vbExclamation
        Exit Sub
    End If
    Name bekleyen As veri
    Mesaj "Yedek yüklendi." & IIf(tasindi, vbCrLf & vbCrLf & "Önceki veri şuraya taşındı:" & vbCrLf & eski, "")
    Exit Sub
Hata:
    Dim hataMetni As String
    hataMetni = Err.Description
    On Error Resume Next
    If tasindi And Not Fso().FileExists(veri) Then Name eski As veri
    Mesaj "Yedek yüklenirken hata: " & hataMetni, vbExclamation
End Sub
