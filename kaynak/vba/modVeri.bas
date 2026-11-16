Option Compare Database
Option Explicit

' =============================================================
' VERİ DOSYASI (bağlantı ve sürüm yükseltme)
' -------------------------------------------------------------
' Program (ApaydinOtomasyon.accdb) yalnızca form, rapor ve kodu
' içerir. Tablolar ayrı veri dosyasındadır:
'     <program klasörü>\Veri\ApaydinVeri.accdb
' Program açılırken (GirisYukle) bağlantı denetlenir:
'   - Bağlantı sağlamsa devam edilir.
'   - Kopuksa önce varsayılan konuma, yoksa kullanıcının seçtiği
'     dosyaya yeniden bağlanılır.
'   - Veri dosyasının sürümü (tbl_Ayarlar: VeriSurum) programın
'     beklediğinden eskiyse VeriGuncelle ile yükseltilir.
' Her yükseltmeden önce veri dosyasının yedeği Yedek klasörüne alınır.
' =============================================================

Public Const VERI_SURUM As Long = 3
Public Const VERI_DOSYA_ADI As String = "ApaydinVeri.accdb"

' Bağlı tabloların gösterdiği veri dosyası
Public Function VeriYolu() As String
    On Error Resume Next
    Dim c As String, p As Long
    c = CurrentDb.TableDefs("tbl_Ayarlar").Connect
    p = InStr(1, c, "DATABASE=", vbTextCompare)
    If p > 0 Then VeriYolu = Mid$(c, p + 9)
End Function

Public Function VarsayilanVeriYolu() As String
    VarsayilanVeriYolu = CurrentProject.Path & "\Veri\" & VERI_DOSYA_ADI
End Function

Public Function OrnekVeriYolu() As String
    OrnekVeriYolu = CurrentProject.Path & "\ornek\ApaydinVeri_ornek.accdb"
End Function

Private Function DosyaVar(ByVal yol As String) As Boolean
    On Error Resume Next
    If yol = "" Then Exit Function
    DosyaVar = (Len(Dir(yol)) > 0)
End Function

' Bağlantı çalışıyor mu? (tablo açılabiliyor mu)
Private Function BaglantiSaglam() As Boolean
    On Error GoTo Yok
    Dim rs As DAO.Recordset
    If Not DosyaVar(VeriYolu()) Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 Anahtar FROM tbl_Ayarlar", dbOpenSnapshot)
    rs.Close
    BaglantiSaglam = True
    Exit Function
Yok:
    BaglantiSaglam = False
End Function

' Tüm bağlı tabloları verilen veri dosyasına yeniden bağlar
Public Function TablolariBagla(ByVal yol As String) As Boolean
    On Error GoTo Hata
    Dim db As DAO.Database, td As DAO.TableDef
    Set db = CurrentDb
    For Each td In db.TableDefs
        If Len(td.Connect) > 0 Then
            td.Connect = ";DATABASE=" & yol
            td.RefreshLink
        End If
    Next
    db.TableDefs.Refresh
    TablolariBagla = True
    Exit Function
Hata:
    Mesaj "Veri dosyasına bağlanılamadı:" & vbCrLf & yol & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Function

' Veri dosyasında yeni oluşturulmuş tablolar varsa programa bağlar
Private Sub EksikTablolariBagla(ByVal yol As String)
    On Error Resume Next
    Dim vdb As DAO.Database, td As DAO.TableDef, ad As String, yeni As DAO.TableDef, tmp As String
    Set vdb = DBEngine.OpenDatabase(yol)
    For Each td In vdb.TableDefs
        ad = td.Name
        If Left$(ad, 4) <> "MSys" And Left$(ad, 1) <> "~" Then
            Err.Clear
            tmp = CurrentDb.TableDefs(ad).Name
            If Err.Number <> 0 Then
                Err.Clear
                Set yeni = CurrentDb.CreateTableDef(ad)
                yeni.Connect = ";DATABASE=" & yol
                yeni.SourceTableName = ad
                CurrentDb.TableDefs.Append yeni
            End If
        End If
    Next
    vdb.Close
End Sub

Private Function VeriDosyasiSec() As String
    On Error GoTo Hata
    Dim fd As Object
    Set fd = Application.FileDialog(3)    ' msoFileDialogFilePicker
    fd.Title = "Veri dosyasını seçin (" & VERI_DOSYA_ADI & ")"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Access veri dosyası", "*.accdb"
    If fd.Show = -1 Then VeriDosyasiSec = fd.SelectedItems(1)
    Exit Function
Hata:
    VeriDosyasiSec = ""
End Function

' Program açılışında çağrılır. False dönerse program kapatılmalıdır.
Public Function VeriBaglantisiHazir() As Boolean
    Dim yol As String
    If Not BaglantiSaglam() Then
        yol = VarsayilanVeriYolu()
        ' İlk açılış: veri dosyası yoksa örnek veriyle başlatmayı öner
        If Not DosyaVar(yol) And DosyaVar(OrnekVeriYolu()) Then
            If Soru("Veri dosyası bulunamadı." & vbCrLf & vbCrLf & _
                    "Örnek verilerle yeni bir veri dosyası oluşturulsun mu?" & vbCrLf & yol) Then
                On Error Resume Next
                If Dir(CurrentProject.Path & "\Veri", vbDirectory) = "" Then MkDir CurrentProject.Path & "\Veri"
                FileCopy OrnekVeriYolu(), yol
                On Error GoTo 0
            End If
        End If
        If Not DosyaVar(yol) Then
            Mesaj "Veri dosyasına ulaşılamadı:" & vbCrLf & VeriYolu() & vbCrLf & vbCrLf & _
                  "Lütfen veri dosyasını (" & VERI_DOSYA_ADI & ") seçin.", vbExclamation
            yol = VeriDosyasiSec()
            If yol = "" Then Exit Function
        End If
        If Not TablolariBagla(yol) Then Exit Function
        If Not BaglantiSaglam() Then Exit Function
    End If
    VeriBaglantisiHazir = VeriGuncelle()
End Function

' Kur.vbs tarafından çağrılır
Public Function VeriGuncelleKurulum(ByVal yol As String) As String
    TestModuAyarla True
    If VeriYolu() <> yol Then TablolariBagla yol
    If VeriGuncelle() Then
        VeriGuncelleKurulum = "Veri sürümü " & VeriSurumu()
    Else
        VeriGuncelleKurulum = "HATA: " & SonMesaj
    End If
    TestModuAyarla False
End Function

' Kur.vbs bu yordamı çağırarak modüllerin derlendiğini denetler
Public Function KurulumDerlemeKontrol() As String
    Dim s As String
    s = SezonBul(Date) & "|" & ProgramBasligi() & "|" & VERI_SURUM
    KurulumDerlemeKontrol = s
End Function

Public Function VeriSurumu() As Long
    Dim s As String
    s = Ayar("VeriSurum")
    If IsNumeric(s) Then VeriSurumu = CLng(s) Else VeriSurumu = 1
End Function

Private Sub VeriYedekle(ByVal yol As String)
    On Error Resume Next
    Dim klasor As String
    ' Yedek klasörü veri klasörünün yanındadır: ...\Veri\ApaydinVeri.accdb -> ...\Yedek
    klasor = Left$(yol, InStrRev(yol, "\") - 1)
    klasor = Left$(klasor, InStrRev(klasor, "\") - 1) & "\Yedek"
    If Dir(klasor, vbDirectory) = "" Then MkDir klasor
    FileCopy yol, klasor & "\ApaydinVeri_surum" & VeriSurumu() & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".accdb"
End Sub

' ---------------------------------------------------------------
' Sürüm yükseltme
' ---------------------------------------------------------------
Public Function VeriGuncelle() As Boolean
    On Error GoTo Hata
    Dim yol As String, surum As Long, vdb As DAO.Database
    surum = VeriSurumu()
    If surum >= VERI_SURUM Then
        VeriGuncelle = True
        Exit Function
    End If
    yol = VeriYolu()
    VeriYedekle yol
    Set vdb = DBEngine.OpenDatabase(yol)

    If surum < 2 Then Surum2 vdb

    vdb.Close
    Set vdb = Nothing
    TablolariBagla yol
    EksikTablolariBagla yol

    If surum < 2 Then Surum2Veri
    If surum < 3 Then Surum3Veri

    AyarYaz "VeriSurum", VERI_SURUM
    VeriGuncelle = True
    Exit Function
Hata:
    Mesaj "Veri dosyası güncellenemedi (sürüm " & surum & " -> " & VERI_SURUM & "):" & vbCrLf & Err.Description, vbCritical
    On Error Resume Next
    If Not vdb Is Nothing Then vdb.Close
End Function

Private Function AlanVar(ByVal vdb As DAO.Database, ByVal tablo As String, ByVal alan As String) As Boolean
    On Error Resume Next
    Dim s As String
    s = vdb.TableDefs(tablo).Fields(alan).Name
    AlanVar = (Err.Number = 0)
    Err.Clear
End Function

Private Function IndeksVar(ByVal vdb As DAO.Database, ByVal tablo As String, ByVal ad As String) As Boolean
    On Error Resume Next
    Dim s As String
    s = vdb.TableDefs(tablo).Indexes(ad).Name
    IndeksVar = (Err.Number = 0)
    Err.Clear
End Function

Private Sub AlanEkle(ByVal vdb As DAO.Database, ByVal tablo As String, ByVal alan As String, ByVal tip As String)
    If Not AlanVar(vdb, tablo, alan) Then vdb.Execute "ALTER TABLE " & tablo & " ADD COLUMN " & alan & " " & tip, dbFailOnError
End Sub

Private Sub IndeksEkle(ByVal vdb As DAO.Database, ByVal tablo As String, ByVal ad As String, ByVal alanlar As String)
    If Not IndeksVar(vdb, tablo, ad) Then vdb.Execute "CREATE INDEX " & ad & " ON " & tablo & " (" & alanlar & ")", dbFailOnError
End Sub

' Parti No'yu tekrar kullanılabilir yapmak için PartiKodu metne çevrilir
' (105, 105/1 ...), sezon, brüt/dara, yağ fiyatı ve tahsilat alanları eklenir.
Private Sub Surum2(ByVal vdb As DAO.Database)
    Dim rel As DAO.Relation, i As Long

    ' Müşteri
    AlanEkle vdb, "tbl_Musteri", "BabaAdi", "TEXT(50)"

    ' Parti: PartiKodu sayı -> metin
    If vdb.TableDefs("tbl_Parti").Fields("PartiKodu").Type <> dbText Then
        If IndeksVar(vdb, "tbl_Parti", "UX_tbl_Parti_PartiKodu") Then vdb.Execute "DROP INDEX UX_tbl_Parti_PartiKodu ON tbl_Parti", dbFailOnError
        For i = vdb.Relations.Count - 1 To 0 Step -1
            Set rel = vdb.Relations(i)
            If rel.Fields.Count > 0 Then
                If rel.Table = "tbl_Parti" And rel.Fields(0).Name = "PartiKodu" Then vdb.Relations.Delete rel.Name
            End If
        Next
        vdb.Execute "ALTER TABLE tbl_Parti ALTER COLUMN PartiKodu TEXT(20)", dbFailOnError
    End If
    AlanEkle vdb, "tbl_Parti", "Sezon", "TEXT(9)"
    AlanEkle vdb, "tbl_Parti", "BrutKg", "DOUBLE"
    AlanEkle vdb, "tbl_Parti", "DaraKg", "DOUBLE"
    AlanEkle vdb, "tbl_Parti", "YagFiyati", "CURRENCY"
    AlanEkle vdb, "tbl_Parti", "KabulAciklama", "TEXT(255)"
    AlanEkle vdb, "tbl_Parti", "TahsilatTuru", "TEXT(20)"
    AlanEkle vdb, "tbl_Parti", "TahsilatTutari", "CURRENCY"
    AlanEkle vdb, "tbl_Parti", "TahsilatYontemi", "TEXT(20)"
    AlanEkle vdb, "tbl_Parti", "TahsilatKasaID", "LONG"
    IndeksEkle vdb, "tbl_Parti", "IX_Parti_SezonKod", "Sezon, PartiKodu"
    IndeksEkle vdb, "tbl_Parti", "IX_Parti_Tarih", "Tarih"

    ' Kasa: nakit kasa ve banka hesapları
    AlanEkle vdb, "tbl_Kasa", "KasaTuru", "TEXT(10)"
    AlanEkle vdb, "tbl_Kasa", "BankaAdi", "TEXT(100)"
    AlanEkle vdb, "tbl_Kasa", "IBAN", "TEXT(34)"
    vdb.Execute "UPDATE tbl_Kasa SET KasaTuru='Nakit' WHERE KasaTuru Is Null", dbFailOnError

    ' Hızlandırma: otomatik kayıt silme ve ekstre sorguları için indeksler
    IndeksEkle vdb, "tbl_CariHareket", "IX_Cari_Kaynak", "KaynakTipi, KaynakID"
    IndeksEkle vdb, "tbl_KasaHareket", "IX_Kasa_Kaynak", "KaynakTipi, KaynakID"
    IndeksEkle vdb, "tbl_StokHareket", "IX_Stok_Kaynak", "KaynakTipi, KaynakID"
    IndeksEkle vdb, "tbl_KasaHareket", "IX_Kasa_KasaTarih", "KasaID, Tarih"
    IndeksEkle vdb, "tbl_PartiBidon", "IX_Bidon_No", "BidonNo"
End Sub

' Yapı değiştikten sonra (tablolar yeniden bağlıyken) verileri düzenler
Private Sub Surum2Veri()
    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = CurrentDb

    ' Ayarlar
    If Ayar("HakyagiOrani") = "" Then AyarYaz "HakyagiOrani", IIf(Ayar("VarsayilanHakyagiOrani") = "", "12", Ayar("VarsayilanHakyagiOrani"))
    If Ayar("YagFiyati") = "" Then AyarYaz "YagFiyati", "200"
    If Ayar("AktifSezon") = "" Then AyarYaz "AktifSezon", SezonBul(Date)

    ' Partilerin sezonu ve eski peşin tahsilatlar
    Set rs = db.OpenRecordset("SELECT * FROM tbl_Parti", dbOpenDynaset)
    Do Until rs.EOF
        rs.Edit
        If IsNull(rs!sezon) Then rs!sezon = SezonBul(rs!tarih)
        If IsNull(rs!TahsilatTuru) Then
            If UcretliMi(rs!odemeTipi) And Nz(rs!PesinTahsilEdildi, False) Then
                rs!TahsilatTuru = "Peşin"
                rs!TahsilatTutari = Nz(rs!TLTutari, 0)
                rs!TahsilatYontemi = "Nakit"
                rs!TahsilatKasaID = VarsayilanKasaID()
                If IsNull(rs!TahsilatTarihi) Then rs!TahsilatTarihi = Nz(rs!TeslimTarihi, rs!tarih)
            Else
                rs!TahsilatTuru = "Yok"
            End If
        End If
        If IsNull(rs!SikimTarihi) And Nz(rs!CikanYag_Kg, 0) > 0 Then rs!SikimTarihi = rs!tarih
        If IsNull(rs!brutKg) And Nz(rs!ZeytinMiktari_Kg, 0) > 0 Then
            rs!brutKg = rs!ZeytinMiktari_Kg
            rs!daraKg = 0
        End If
        rs.Update
        rs.MoveNext
    Loop
    rs.Close

    ' Otomatik kayıtları yeni kurallarla yeniden yaz (tutarlar korunur)
    Set rs = db.OpenRecordset("SELECT PartiNo FROM tbl_Parti", dbOpenSnapshot)
    Do Until rs.EOF
        PartiMuhasebeYaz rs!partiNo
        TeslimMuhasebeYaz rs!partiNo
        BidonlariEsle rs!partiNo
        rs.MoveNext
    Loop
    rs.Close
End Sub

' Sürüm 3: ilk kurulumda Türkçe harfleri bozuk yazılmış otomatik kayıtları
' (ör. "Peþin", "Sýkým Ücreti") düzeltir ve otomatik kayıtları yeniden yazar.
Private Sub Surum3Veri()
    Dim db As DAO.Database
    Set db = CurrentDb
    db.Execute "UPDATE tbl_Parti SET TahsilatTuru='Peşin' WHERE TahsilatTuru Like 'Pe*' AND TahsilatTuru<>'Peşin'", dbFailOnError
    db.Execute "UPDATE tbl_Parti SET TahsilatTuru='Kısmi' WHERE TahsilatTuru Like 'K*' AND TahsilatTuru<>'Kısmi'", dbFailOnError
    db.Execute "UPDATE tbl_Parti SET TahsilatTuru='Yok' WHERE TahsilatTuru Is Null", dbFailOnError
    TumPartiKayitlariniYenidenYaz
End Sub
