Option Compare Database
Option Explicit

' =============================================================
' OTOMATİK GÜNCELLEME   (form: frm_GuncellemeBilgi)
' -------------------------------------------------------------
' Program her açılışta gizli app_licenses deposundaki
' apaydin_otomasyon/surum.txt bildirimini okur (lisans dosyasıyla
' aynı okuma anahtarı):
'
'   SURUM / ADRES / SHA256 / ZORUNLU / NOTLAR          herkese açık sürüm
'   ADRES_ACCDB / SHA256_ACCDB                          32 bit Access için .accdb
'   DENEME_SURUM / DENEME_ADRES / DENEME_SHA256 / ...  deneme sürümü
'   DENEME_KODLAR = AO-..., AO-...                     deneme sürümünü alacak kurulumlar
'
' Deneme sürümünü varsayılan olarak yalnızca geliştirici anahtarı
' olan kurulumlar alır; DENEME_KODLAR'a yazılan kurulumlar da alır.
' Herkese açmak için değerler DENEME_ öneki olmadan yazılır.
'
' Dosyalar herkese açık ApaydinOtomasyon deposunun Releases bölümünden
' zip olarak indirilir (GitHub .accde/.accdb yüklemeye izin vermiyor),
' zip SHA-256 ile doğrulanır ve içinden program dosyası çıkarılır. Zorunlu sürümler sorulmadan,
' diğerleri onayla kurulur. Açık program kendini değiştiremediği için
' program klasörüne Guncelle.vbs yazılır ve çalıştırılır: program
' kapanınca eski sürümü Yedek klasörüne alır, yenisini koyar, masaüstü
' kısayollarını yeni dosyaya çevirir ve programı yeniden açar.
' Sonuç guncelleme.sonuc dosyasıyla bir sonraki açılışta gösterilir.
'
' Kaynaktan kurulmuş geliştirme kopyasında (.accdb + kaynak klasörü +
' geliştirici anahtarı) otomatik güncelleme yapılmaz.
' =============================================================

Private Const BILDIRIM_DOSYASI As String = "apaydin_otomasyon/surum.txt"
Private Const BETIK_ADI As String = "Guncelle.vbs"
Private Const AYAR_DOSYASI As String = "guncelleme.ayar"
Private Const SONUC_DOSYASI As String = "guncelleme.sonuc"
Private Const INDIRME_SURESI As Single = 300

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private gSurum As String
Private gAdres As String
Private gOzet As String
Private gZorunlu As Boolean
Private gNotlar As String
Private gDeneme As Boolean
Private gUzanti As String

' ---------------------------------------------------------------
' Yardımcılar
' ---------------------------------------------------------------
Private Function Fso() As Object
    Set Fso = CreateObject("Scripting.FileSystemObject")
End Function

Private Sub SilVarsa(ByVal yol As String)
    On Error Resume Next
    If Fso().FileExists(yol) Then Fso().DeleteFile yol, True
End Sub

Public Function KorumaliSurum() As Boolean
    KorumaliSurum = (LCase$(Right$(CurrentProject.Name, 6)) = ".accde")
End Function

' Kaynaktan kurulmuş geliştirme kopyası mı? (orada otomatik güncelleme yapılmaz)
Public Function GelistirmeKopyasi() As Boolean
    If KorumaliSurum() Then Exit Function
    If Not GelistiriciAnahtariVar() Then Exit Function
    GelistirmeKopyasi = Fso().FileExists(CurrentProject.Path & "\kaynak\vba\modOrtak.bas")
End Function

' Sürüm karşılaştırma: "3.10" > "3.9". a büyükse 1, eşitse 0, küçükse -1
Public Function SurumKarsilastir(ByVal a As String, ByVal b As String) As Long
    Dim x() As String, y() As String, i As Long, n As Long, sa As Long, sb As Long
    x = Split(Trim$(a), ".")
    y = Split(Trim$(b), ".")
    n = UBound(x)
    If UBound(y) > n Then n = UBound(y)
    For i = 0 To n
        sa = 0: sb = 0
        If i <= UBound(x) Then sa = Val(x(i))
        If i <= UBound(y) Then sb = Val(y(i))
        If sa <> sb Then
            SurumKarsilastir = Sgn(sa - sb)
            Exit Function
        End If
    Next
End Function

Private Function NotuAt(ByVal s As String) As String
    Dim p As Long
    s = Trim$(s)
    If Left$(s, 1) = "#" Then s = ""
    p = InStr(1, s, " #")
    If p > 0 Then s = Left$(s, p - 1)
    NotuAt = Trim$(s)
End Function

Private Function BildirimCoz(ByVal metin As String) As Object
    Dim d As Object, satirlar() As String, i As Long, s As String, p As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    metin = Replace(metin, vbCr, "")
    satirlar = Split(metin, vbLf)
    For i = 0 To UBound(satirlar)
        s = Trim$(satirlar(i))
        If s <> "" And Left$(s, 1) <> "#" Then
            p = InStr(1, s, "=")
            If p > 1 Then d(UCase$(Trim$(Left$(s, p - 1)))) = NotuAt(Mid$(s, p + 1))
        End If
    Next
    Set BildirimCoz = d
End Function

Private Function BDeger(ByVal d As Object, ByVal anahtar As String) As String
    If d.Exists(anahtar) Then BDeger = d(anahtar)
End Function

Private Function KodListede(ByVal kod As String, ByVal liste As String) As Boolean
    Dim p As Variant
    For Each p In Split(Replace(liste, ";", ","), ",")
        If Trim$(p) <> "" And UCase$(Trim$(p)) = UCase$(kod) Then
            KodListede = True
            Exit Function
        End If
    Next
End Function

Private Function NotMetni(ByVal notlar As String) As String
    If Trim$(notlar) = "" Then Exit Function
    NotMetni = vbCrLf & vbCrLf & "Yenilikler:" & vbCrLf & "• " & Replace(Replace(notlar, " | ", "|"), "|", vbCrLf & "• ")
End Function

' ---------------------------------------------------------------
' Bildirimden bu kurulum için yeni sürümü seçer
' ---------------------------------------------------------------
Private Function YeniSurumBul(ByVal metin As String) As Boolean
    Dim d As Object, ek As String, onEk As String, denemeHakki As Boolean, adayDeneme As Boolean
    Set d = BildirimCoz(metin)
    #If Win64 Then
        ek = ""
        gUzanti = "accde"
    #Else
        ek = "_ACCDB"
        gUzanti = "accdb"
    #End If
    denemeHakki = GelistiriciAnahtariVar() Or KodListede(LisansKodu(), BDeger(d, "DENEME_KODLAR"))
    If denemeHakki And BDeger(d, "DENEME_SURUM") <> "" And BDeger(d, "DENEME_ADRES" & ek) <> "" Then
        If SurumKarsilastir(BDeger(d, "DENEME_SURUM"), URUN_SURUM) > 0 And _
           SurumKarsilastir(BDeger(d, "DENEME_SURUM"), BDeger(d, "SURUM")) > 0 Then adayDeneme = True
    End If
    If adayDeneme Then onEk = "DENEME_" Else onEk = ""
    gSurum = BDeger(d, onEk & "SURUM")
    gAdres = BDeger(d, onEk & "ADRES" & ek)
    gOzet = LCase$(BDeger(d, onEk & "SHA256" & ek))
    gZorunlu = (LCase$(BDeger(d, onEk & "ZORUNLU")) = "evet")
    gNotlar = BDeger(d, onEk & "NOTLAR")
    gDeneme = adayDeneme
    If gSurum = "" Or gAdres = "" Or Len(gOzet) <> 64 Then Exit Function
    YeniSurumBul = (SurumKarsilastir(gSurum, URUN_SURUM) > 0)
End Function

' ---------------------------------------------------------------
' Denetim
' ---------------------------------------------------------------
' Açılışta (elle=False) ve Ayarlar'daki düğmeyle (elle=True) çağrılır.
' Güncelleme kuruluyorsa program kapanır ve True döner.
Public Function GuncellemeDenetle(Optional ByVal elle As Boolean = False) As Boolean
    On Error GoTo Hata
    Dim metin As String, soruMetni As String, gecici As String
    If TestModu Then Exit Function
    If GelistirmeKopyasi() Then
        If elle Then Mesaj "Bu kopya kaynaktan kurulmuş geliştirme kopyası; otomatik güncelleme yapılmaz."
        Exit Function
    End If
    If Not GizliDosyaOku(BILDIRIM_DOSYASI, metin) Then
        If elle Then Mesaj "Güncelleme bilgisine ulaşılamadı. İnternet bağlantısını kontrol edin.", vbExclamation
        Exit Function
    End If
    AyarYaz "GuncellemeSonDenetim", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    If Not YeniSurumBul(metin) Then
        If elle Then Mesaj "Program güncel (sürüm " & URUN_SURUM & ")."
        Exit Function
    End If

    If gZorunlu Then
        Mesaj "Zorunlu güncelleme: sürüm " & gSurum & " kuruluyor." & NotMetni(gNotlar) & vbCrLf & vbCrLf & _
              "Program kapanıp yeniden açılacak; verileriniz değişmez."
    Else
        soruMetni = "Yeni sürüm var: " & gSurum & IIf(gDeneme, "  (deneme sürümü)", "") & _
               "    Kullandığınız: " & URUN_SURUM & NotMetni(gNotlar) & vbCrLf & vbCrLf & _
               "Şimdi güncellensin mi? Program kapanıp yeniden açılacak; verileriniz değişmez."
        If Not Soru(soruMetni) Then Exit Function
    End If

    Dim zipYol As String, acilan As String
    zipYol = Environ$("TEMP") & "\ApaydinOtomasyon_" & gSurum & ".zip"
    acilan = Environ$("TEMP") & "\ApaydinOtomasyon_" & gSurum
    SilVarsa zipYol
    DoCmd.Hourglass True
    If Not DosyaIndir(gAdres, zipYol) Then
        DoCmd.Hourglass False
        Mesaj "Yeni sürüm indirilemedi. İnternet bağlantısını kontrol edip daha sonra tekrar deneyin." & vbCrLf & _
              "Sorun sürerse Ayarlar > Güncelleme ve Güvenlik Bilgisi bölümüne bakın.", vbExclamation
        Exit Function
    End If
    If DosyaOzeti(zipYol) <> gOzet Then
        DoCmd.Hourglass False
        SilVarsa zipYol
        Mesaj "İndirilen dosya doğrulanamadı (bozuk ya da eksik indi); güncelleme yapılmadı.", vbExclamation
        Exit Function
    End If
    gecici = acilan & "\ApaydinOtomasyon." & gUzanti
    If Not ZipAc(zipYol, acilan, gecici) Then
        DoCmd.Hourglass False
        Mesaj "İndirilen güncelleme açılamadı. Antivirüs engellemiş olabilir; " & _
              "Ayarlar > Güncelleme ve Güvenlik Bilgisi bölümüne bakın.", vbExclamation
        Exit Function
    End If
    SilVarsa zipYol
    DoCmd.Hourglass False
    AyarYaz "GuncellemeNotlari", Left$(gNotlar, 250)
    GuncellemeDenetle = True
    GuncellemeyiBaslat gecici
    Exit Function
Hata:
    DoCmd.Hourglass False
    If elle Then Mesaj "Güncelleme denetlenemedi: " & Err.Description, vbExclamation
End Function

' GitHub indirmeleri başka bir sunucuya yönlenir; tarayıcı tabanlı XMLHTTP bu
' yönlendirmeyi "Erişim engellendi" ile keser. Önce sunucu tabanlı istemciler denenir.
Private Function DosyaIndir(ByVal adres As String, ByVal hedef As String) As Boolean
    Dim sinif As Variant
    SysCmd acSysCmdSetStatus, "Yeni sürüm indiriliyor..."
    For Each sinif In Array("MSXML2.ServerXMLHTTP.6.0", "WinHttp.WinHttpRequest.5.1", "MSXML2.XMLHTTP.6.0")
        If IstemciyleIndir(CStr(sinif), adres, hedef) Then
            DosyaIndir = True
            Exit For
        End If
    Next
    SysCmd acSysCmdClearStatus
End Function

Private Function IstemciyleIndir(ByVal sinif As String, ByVal adres As String, ByVal hedef As String) As Boolean
    On Error GoTo Hata
    Dim h As Object, st As Object
    Set h = CreateObject(sinif)
    If sinif = "MSXML2.ServerXMLHTTP.6.0" Then h.setTimeouts 15000, 15000, 30000, INDIRME_SURESI * 1000
    If sinif = "WinHttp.WinHttpRequest.5.1" Then h.SetTimeouts 15000, 15000, 30000, INDIRME_SURESI * 1000
    h.Open "GET", adres, False
    h.send
    If h.Status <> 200 Then Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write h.responseBody
    st.SaveToFile hedef, 2
    st.Close
    IstemciyleIndir = True
    Exit Function
Hata:
End Function

' Zip'i Windows'un kendi zip desteğiyle açar, beklenen dosya çıkana kadar bekler
Private Function ZipAc(ByVal zipYol As String, ByVal klasor As String, ByVal beklenen As String) As Boolean
    On Error GoTo Hata
    Dim kabuk As Object, kaynak As Object, hedef As Object, bas As Single, boyut As Double
    If Fso().FolderExists(klasor) Then Fso().DeleteFolder klasor, True
    Fso().CreateFolder klasor
    Set kabuk = CreateObject("Shell.Application")
    Set kaynak = kabuk.Namespace(CVar(zipYol))
    Set hedef = kabuk.Namespace(CVar(klasor))
    If kaynak Is Nothing Or hedef Is Nothing Then Exit Function
    hedef.CopyHere kaynak.Items, 4 + 16 + 1024
    bas = Timer
    boyut = -1
    Do
        DoEvents
        Sleep 300
        If Fso().FileExists(beklenen) Then
            If FileLen(beklenen) > 0 And FileLen(beklenen) = boyut Then
                ZipAc = True
                Exit Function
            End If
            boyut = FileLen(beklenen)
        End If
    Loop While Timer - bas < 60 And Timer >= bas
    Exit Function
Hata:
End Function

Private Function DosyaOzeti(ByVal yol As String) As String
    On Error GoTo Hata
    Dim f As Integer, b() As Byte, n As Long
    n = FileLen(yol)
    If n = 0 Then Exit Function
    ReDim b(0 To n - 1)
    f = FreeFile
    Open yol For Binary Access Read Shared As #f
    Get #f, , b
    Close #f
    DosyaOzeti = Sha256Bayt(b)
    Exit Function
Hata:
    Close
End Function

' ---------------------------------------------------------------
' Kurulum: Guncelle.vbs yazılır, çalıştırılır, program kapanır
' ---------------------------------------------------------------
Private Sub GuncellemeyiBaslat(ByVal gecici As String)
    Dim klasor As String, hedef As String, betik As String, f As Object
    klasor = CurrentProject.Path
    hedef = klasor & "\ApaydinOtomasyon." & gUzanti
    betik = klasor & "\" & BETIK_ADI
    BetikYaz betik
    Set f = Fso().CreateTextFile(klasor & "\" & AYAR_DOSYASI, True, True)
    f.WriteLine CurrentProject.FullName
    f.WriteLine gecici
    f.WriteLine hedef
    f.WriteLine klasor & "\Yedek"
    f.WriteLine SysCmd(acSysCmdAccessDir) & "MSACCESS.EXE"
    f.WriteLine gSurum
    f.Close
    CreateObject("WScript.Shell").Run """" & Environ$("SystemRoot") & "\System32\wscript.exe"" """ & betik & """", 0, False
    CikisOnay = True
    DoCmd.Quit acQuitSaveNone
End Sub

Private Sub BSatir(ByRef metin As String, ByVal satir As String)
    metin = metin & Replace(satir, "`", Chr$(34)) & vbCrLf
End Sub

' Güncelleyici betik (yalnızca ASCII; tırnak yerine ` yazılır)
Private Sub BetikYaz(ByVal yol As String)
    Dim m As String, f As Object
    BSatir m, "' Apaydin Otomasyon - guncelleyici. Program tarafindan olusturulur; elle calistirmayin."
    BSatir m, "Option Explicit"
    BSatir m, "Dim fso, sh, klasor, f, eski, gecici, hedef, yedek, accessExe, surum, i, ok, hata, yedekYol, ts"
    BSatir m, "Set fso = CreateObject(`Scripting.FileSystemObject`)"
    BSatir m, "Set sh = CreateObject(`WScript.Shell`)"
    BSatir m, "klasor = fso.GetParentFolderName(WScript.ScriptFullName)"
    BSatir m, "Set f = fso.OpenTextFile(klasor & `\guncelleme.ayar`, 1, False, -1)"
    BSatir m, "eski = f.ReadLine: gecici = f.ReadLine: hedef = f.ReadLine"
    BSatir m, "yedek = f.ReadLine: accessExe = f.ReadLine: surum = f.ReadLine"
    BSatir m, "f.Close"
    BSatir m, "ts = Year(Now) & Right(`0` & Month(Now), 2) & Right(`0` & Day(Now), 2) & `_` & _"
    BSatir m, "     Right(`0` & Hour(Now), 2) & Right(`0` & Minute(Now), 2) & Right(`0` & Second(Now), 2)"
    BSatir m, "If Not fso.FolderExists(yedek) Then fso.CreateFolder yedek"
    BSatir m, "yedekYol = yedek & `\` & fso.GetBaseName(eski) & `_eski_` & ts & `.` & fso.GetExtensionName(eski)"
    BSatir m, "' 1. Programin kapanmasini bekle, eski surumu Yedek klasorune tasi (en fazla 60 sn)"
    BSatir m, "ok = False"
    BSatir m, "For i = 1 To 120"
    BSatir m, "    On Error Resume Next"
    BSatir m, "    Err.Clear"
    BSatir m, "    fso.MoveFile eski, yedekYol"
    BSatir m, "    If Err.Number = 0 Then ok = True"
    BSatir m, "    On Error GoTo 0"
    BSatir m, "    If ok Then Exit For"
    BSatir m, "    WScript.Sleep 500"
    BSatir m, "Next"
    BSatir m, "If Not ok Then"
    BSatir m, "    Sonuc `HATA|Program dosyasi kullanimda oldugu icin degistirilemedi.`"
    BSatir m, "    Calistir eski"
    BSatir m, "    WScript.Quit"
    BSatir m, "End If"
    BSatir m, "' 2. Yeni surumu yerine koy"
    BSatir m, "On Error Resume Next"
    BSatir m, "If LCase(hedef) <> LCase(eski) And fso.FileExists(hedef) Then fso.MoveFile hedef, yedek & `\` & fso.GetBaseName(hedef) & `_eski_` & ts & `.` & fso.GetExtensionName(hedef)"
    BSatir m, "Err.Clear"
    BSatir m, "fso.CopyFile gecici, hedef, True"
    BSatir m, "hata = Err.Number"
    BSatir m, "On Error GoTo 0"
    BSatir m, "If hata <> 0 Or Not fso.FileExists(hedef) Then"
    BSatir m, "    On Error Resume Next"
    BSatir m, "    fso.MoveFile yedekYol, eski"
    BSatir m, "    On Error GoTo 0"
    BSatir m, "    Sonuc `HATA|Yeni surum yerine konamadi (hata ` & hata & `). Antivirus engellemis olabilir.`"
    BSatir m, "    Calistir eski"
    BSatir m, "    WScript.Quit"
    BSatir m, "End If"
    BSatir m, "On Error Resume Next"
    BSatir m, "fso.DeleteFolder fso.GetParentFolderName(gecici), True"
    BSatir m, "On Error GoTo 0"
    BSatir m, "' 3. Masaustu kisayollari eski dosyayi gosteriyorsa yenisine cevir"
    BSatir m, "If LCase(hedef) <> LCase(eski) Then Kisayollar eski, hedef"
    BSatir m, "Sonuc `TAMAM|` & surum"
    BSatir m, "Calistir hedef"
    BSatir m, ""
    BSatir m, "Sub Sonuc(metin)"
    BSatir m, "    Dim t"
    BSatir m, "    Set t = fso.CreateTextFile(klasor & `\guncelleme.sonuc`, True, True)"
    BSatir m, "    t.WriteLine metin"
    BSatir m, "    t.Close"
    BSatir m, "End Sub"
    BSatir m, ""
    BSatir m, "' Programi acmadan once gecici guncelleme dosyalarini (bu betik dahil) temizle"
    BSatir m, "Sub Calistir(yol)"
    BSatir m, "    On Error Resume Next"
    BSatir m, "    fso.DeleteFile klasor & `\guncelleme.ayar`, True"
    BSatir m, "    fso.DeleteFile WScript.ScriptFullName, True"
    BSatir m, "    On Error GoTo 0"
    BSatir m, "    sh.Run Chr(34) & accessExe & Chr(34) & ` ` & Chr(34) & yol & Chr(34), 1, False"
    BSatir m, "End Sub"
    BSatir m, ""
    BSatir m, "Sub Kisayollar(eskiYol, yeniYol)"
    BSatir m, "    Dim masa, d, lnk"
    BSatir m, "    On Error Resume Next"
    BSatir m, "    For Each masa In Array(sh.SpecialFolders(`Desktop`), sh.SpecialFolders(`AllUsersDesktop`))"
    BSatir m, "        If fso.FolderExists(masa) Then"
    BSatir m, "            For Each d In fso.GetFolder(masa).Files"
    BSatir m, "                If LCase(fso.GetExtensionName(d.Name)) = `lnk` Then"
    BSatir m, "                    Set lnk = sh.CreateShortcut(d.Path)"
    BSatir m, "                    If LCase(lnk.TargetPath) = LCase(eskiYol) Then"
    BSatir m, "                        lnk.TargetPath = yeniYol"
    BSatir m, "                        lnk.Save"
    BSatir m, "                    End If"
    BSatir m, "                End If"
    BSatir m, "            Next"
    BSatir m, "        End If"
    BSatir m, "    Next"
    BSatir m, "End Sub"
    Set f = Fso().CreateTextFile(yol, True, False)
    f.Write m
    f.Close
End Sub

' Açılışta: önceki güncellemenin sonucunu gösterir
Public Sub GuncellemeSonucuGoster()
    On Error Resume Next
    Dim yol As String, f As Object, s As String
    ' Önceki sürümlerin güncelleyicisi betiği klasörde bırakıyordu
    SilVarsa CurrentProject.Path & "\" & BETIK_ADI
    If Not Fso().FileExists(CurrentProject.Path & "\" & SONUC_DOSYASI) Then SilVarsa CurrentProject.Path & "\" & AYAR_DOSYASI
    yol = CurrentProject.Path & "\" & SONUC_DOSYASI
    If Not Fso().FileExists(yol) Then Exit Sub
    Set f = Fso().OpenTextFile(yol, 1, False, -1)
    s = f.ReadLine
    f.Close
    SilVarsa yol
    SilVarsa CurrentProject.Path & "\" & AYAR_DOSYASI
    If Left$(s, 6) = "TAMAM|" Then
        Mesaj "Program " & Mid$(s, 7) & " sürümüne güncellendi." & NotMetni(Ayar("GuncellemeNotlari"))
    ElseIf Left$(s, 5) = "HATA|" Then
        Mesaj "Güncelleme yapılamadı: " & Mid$(s, 6) & vbCrLf & vbCrLf & _
              "Antivirüs programı güncelleme dosyalarını engellemiş olabilir. " & _
              "Ayarlar > Güncelleme ve Güvenlik Bilgisi bölümüne bakın.", vbExclamation
    End If
End Sub

Public Function GuncellemeBilgiMetni() As String
    Dim t As String, son As String
    t = "Sürüm " & URUN_SURUM & IIf(KorumaliSurum(), " (korumalı)", "")
    If GelistirmeKopyasi() Then
        GuncellemeBilgiMetni = t & "   •   Geliştirme kopyası: otomatik güncelleme yapılmaz."
        Exit Function
    End If
    t = t & "   •   Deneme sürümleri: " & IIf(GelistiriciAnahtariVar(), "açık (geliştirici anahtarı)", "yazılımcı açarsa")
    son = Ayar("GuncellemeSonDenetim")
    If Len(son) >= 16 Then t = t & "   •   Son denetim: " & Mid$(son, 9, 2) & "." & Mid$(son, 6, 2) & "." & Left$(son, 4) & " " & Mid$(son, 12, 5)
    GuncellemeBilgiMetni = t
End Function

Public Function GuncellemeElleDenetle()
    GuncellemeDenetle True
    If FormAcikMi("frm_Ayarlar") Then Forms("frm_Ayarlar")!lblGuncelleme.Caption = GuncellemeBilgiMetni()
End Function

' ---------------------------------------------------------------
' Güncelleme ve güvenlik bilgisi ekranı   (form: frm_GuncellemeBilgi)
' ---------------------------------------------------------------
Public Function GuncellemeBilgisiAc()
    DoCmd.OpenForm "frm_GuncellemeBilgi", WindowMode:=acDialog
End Function

Public Function GuncellemeBilgiYukle()
    On Error Resume Next
    Forms("frm_GuncellemeBilgi")!txtKlasor = CurrentProject.Path
End Function

Public Function WindowsGuvenligiAc()
    On Error Resume Next
    CreateObject("WScript.Shell").Run "windowsdefender://threatsettings/"
    If Err.Number <> 0 Then Mesaj "Windows Güvenliği açılamadı. Başlat menüsünden 'Windows Güvenliği'ni açın.", vbExclamation
End Function

Public Function ProgramKlasorunuAc()
    On Error Resume Next
    CreateObject("WScript.Shell").Run "explorer.exe """ & CurrentProject.Path & """"
End Function

Public Function GuncellemeBilgiKapat()
    DoCmd.Close acForm, "frm_GuncellemeBilgi"
End Function
