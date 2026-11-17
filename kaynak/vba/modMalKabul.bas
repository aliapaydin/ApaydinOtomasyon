Option Compare Database
Option Explicit

' =============================================================
' MAL KABUL (ZEYTİN GİRİŞİ)   (form: frm_MalKabul)
' -------------------------------------------------------------
' Müşterinin getirdiği zeytin tartılır ve parti açılır:
'   Parti No (sezon içinde önerilir, değiştirilebilir), tarih,
'   müşteri (yanındaki "+ Yeni" ile hızlı cari eklenir), cinsi,
'   özellik, ödeme tipi, bidonlar, brüt / dara / net kg.
' Kaydedilen partiye o anki hakyağı oranı ve 1 kg yağ fiyatı
' (Ayarlar) kopyalanır; sonradan Ayarlar değişse de parti etkilenmez.
' A5 Mal Kabul Fişi önizlenir veya yazdırılır.
' Üretimi yapılmış partinin kabul bilgileri değiştirilebilir, ancak
' silinemez (önce Üretim ekranından üretim geri alınmalıdır).
' =============================================================

Private Const FORM_ADI As String = "frm_MalKabul"
Private Const KUTULAR As String = "txtPartiNo,txtPartiKodu,cmbMusteri,txtMusteriBilgi,cmbGirisPersonel,cmbCinsi,cmbOzellik," & _
    "cmbOdemeTipi,txtBidonDetay,txtKabulAciklama,txtBrut,txtDara,txtNet,txtDurum"

Private Function F() As Form
    Set F = Forms(FORM_ADI)
End Function

Public Function MalKabulYukle()
    On Error Resume Next
    Dim frm As Form
    Set frm = F()
    frm!cmbListeFiltre = "Üretim Bekleyenler"
    MalKabulTemizle
    MalKabulYeni
End Function

Public Function MalKabulTemizle()
    Dim frm As Form
    Set frm = F()
    KutulariTemizle frm, Split(KUTULAR, ",")
    frm!txtTarih = Date
    frm!txtSezon = AktifSezon()
    frm!lstPartiler = Null
    MalKabulListeFiltrele
End Function

Public Function MalKabulYeni()
    Dim frm As Form
    MalKabulTemizle
    Set frm = F()
    frm!txtPartiKodu = PartiNoOner(AktifSezon())
    frm!cmbOdemeTipi = "Hakyağı"
    frm!cmbCinsi = "Yağlık Zeytin"
    frm!txtDara = 0
    frm!txtDurum = "Yeni"
    frm!cmbMusteri.SetFocus
End Function

' Tarih değişince sezonu gösterir; yeni kayıtta numara önerisini yeniler
Public Function MalKabulTarihDegisti()
    Dim frm As Form, sezon As String
    Set frm = F()
    If Not TarihMi(frm!txtTarih) Then Exit Function
    sezon = SezonBul(frm!txtTarih)
    If sezon <> Nz(frm!txtSezon, "") Then
        frm!txtSezon = sezon
        If Bos(frm!txtPartiNo) Then frm!txtPartiKodu = PartiNoOner(sezon)
    End If
End Function

Public Function MalKabulMusteriSecildi()
    Dim frm As Form, rs As DAO.Recordset, s As String
    Set frm = F()
    frm!txtMusteriBilgi = Null
    If Bos(frm!cmbMusteri) Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT Telefon, TCKN_VKN, BabaAdi FROM tbl_Musteri WHERE MusteriID=" & CLng(frm!cmbMusteri), dbOpenSnapshot)
    If Not rs.EOF Then
        If Not Bos(rs!BabaAdi) Then s = "Baba: " & rs!BabaAdi
        If Not Bos(rs!Telefon) Then s = s & IIf(s = "", "", "  •  ") & rs!Telefon
        If Not Bos(rs!TCKN_VKN) Then s = s & IIf(s = "", "", "  •  ") & rs!TCKN_VKN
    End If
    rs.Close
    frm!txtMusteriBilgi = MetinVeyaNull(s)
End Function

' Brüt / dara yazılırken (OnChange) ve alandan çıkınca (AfterUpdate) çağrılır.
' Yazılmakta olan alanın değeri henüz Value'ya geçmediği için .Text okunur.
Public Function MalKabulHesapla()
    Dim frm As Form, brut As Double, dara As Double
    Set frm = F()
    brut = AlanSayisi(frm, "txtBrut")
    dara = AlanSayisi(frm, "txtDara")
    If brut > 0 Then
        frm!txtNet = brut - dara
    Else
        frm!txtNet = Null
    End If
End Function

Private Function AlanSayisi(ByVal frm As Form, ByVal ad As String) As Double
    Dim aktif As String
    On Error Resume Next
    aktif = Screen.ActiveControl.Name
    On Error GoTo 0
    If aktif = ad Then
        AlanSayisi = Sayi(frm.Controls(ad).Text)
    Else
        AlanSayisi = Sayi(frm.Controls(ad).Value)
    End If
End Function

Public Function MalKabulBidonDuzenle()
    Dim frm As Form
    Set frm = F()
    frm!txtBidonDetay = MetinVeyaNull(BidonlariAyristir(frm!txtBidonDetay))
End Function

' "+ Yeni" düğmesi: hızlı müşteri ekleme penceresi
Public Function MalKabulYeniCari()
    DoCmd.OpenForm "frm_HizliCari", , , , , , FORM_ADI
End Function

' ---------------------------------------------------------------
' Kaydet / Güncelle
' ---------------------------------------------------------------
Public Function MalKabulKaydet()
    MalKabulYazIslem False
End Function

Public Function MalKabulGuncelle()
    MalKabulYazIslem True
End Function

Private Function MalKabulKontrol(ByVal partiNo As Long) As Boolean
    Dim frm As Form, kod As String, sezon As String, cakisan As String
    Set frm = F()
    kod = Trim$(Nz(frm!txtPartiKodu, ""))
    If kod = "" Then
        Mesaj "Parti No giriniz.", vbExclamation: Exit Function
    End If
    If Not TarihMi(frm!txtTarih) Then
        Mesaj "Tarih giriniz.", vbExclamation: Exit Function
    End If
    sezon = SezonBul(frm!txtTarih)
    If PartiNoMesgul(sezon, kod, partiNo) Then
        Mesaj "Parti No " & kod & " bu sezonda yağı henüz teslim edilmemiş başka bir partide kullanılıyor." & vbCrLf & _
              "Önerilen boş numara: " & PartiNoOner(sezon), vbExclamation
        Exit Function
    End If
    If Bos(frm!cmbMusteri) Then
        Mesaj "Müşteri seçiniz (yoksa '+ Yeni' ile ekleyin).", vbExclamation: Exit Function
    End If
    If Sayi(frm!txtBrut) <= 0 Then
        Mesaj "Brüt kg giriniz.", vbExclamation: Exit Function
    End If
    If Sayi(frm!txtDara) < 0 Or Sayi(frm!txtDara) >= Sayi(frm!txtBrut) Then
        Mesaj "Dara, brüt kilodan küçük olmalıdır.", vbExclamation: Exit Function
    End If
    If Not UcretliMi(frm!cmbOdemeTipi) And Not HakyagiMi(frm!cmbOdemeTipi) Then
        Mesaj "Ödeme tipini seçiniz (Hakyağı veya Sıkım Ücreti).", vbExclamation: Exit Function
    End If
    If partiNo <> 0 Then
        If Not IsNull(DLookup("SikimTarihi", "tbl_Parti", "PartiNo=" & partiNo)) Then
            If Abs(Nz(DLookup("ZeytinMiktari_Kg", "tbl_Parti", "PartiNo=" & partiNo), 0) - Sayi(frm!txtNet)) > 0.001 Then
                If Not Soru("Bu partinin üretimi yapılmış. Net zeytin kilosu değişirse randıman da değişir." & vbCrLf & "Devam edilsin mi?") Then Exit Function
            End If
        End If
    End If
    cakisan = BidonCakismasi(BidonlariAyristir(frm!txtBidonDetay), sezon, partiNo)
    If cakisan <> "" Then
        If Not Soru("Şu bidonlar teslim edilmemiş başka bir partide de kayıtlı: " & cakisan & vbCrLf & "Yine de kaydedilsin mi?") Then Exit Function
    End If
    MalKabulKontrol = True
End Function

Private Sub MalKabulYazIslem(ByVal guncelle As Boolean)
    On Error GoTo Hata
    Dim frm As Form, partiNo As Long, rs As DAO.Recordset, islemde As Boolean, yeni As Boolean
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
    yeni = (partiNo = 0)

    MalKabulHesapla
    MalKabulBidonDuzenle
    If Not MalKabulKontrol(partiNo) Then Exit Sub

    DBEngine.BeginTrans: islemde = True
    Set rs = KayitAc("tbl_Parti", "PartiNo", partiNo)
    rs!PartiKodu = Trim$(frm!txtPartiKodu)
    rs!tarih = CDate(frm!txtTarih)
    rs!sezon = SezonBul(frm!txtTarih)
    rs!musteriID = CLng(frm!cmbMusteri)
    rs!Cinsi = MetinVeyaNull(frm!cmbCinsi)
    rs!Ozellik = MetinVeyaNull(frm!cmbOzellik)
    rs!GirisPersonelID = SayiVeyaNull(frm!cmbGirisPersonel)
    rs!brutKg = Sayi(frm!txtBrut)
    rs!daraKg = Sayi(frm!txtDara)
    rs!ZeytinMiktari_Kg = Sayi(frm!txtNet)
    rs!odemeTipi = IIf(UcretliMi(frm!cmbOdemeTipi), "Sıkım Ücreti", "Hakyağı")
    rs!BidonDetay = MetinVeyaNull(frm!txtBidonDetay)
    rs!KabulAciklama = MetinVeyaNull(frm!txtKabulAciklama)
    If yeni Then
        ' Ayarlardaki oran ve fiyat partiye kopyalanır
        rs!HakYagiOrani_Yuzde = VarsayilanOran()
        rs!YagFiyati = VarsayilanYagFiyati()
        rs!PesinTahsilEdildi = False
        rs!TahsilatTuru = "Yok"
    End If
    rs.Update
    rs.Bookmark = rs.LastModified
    partiNo = rs!partiNo
    rs.Close

    ' Üretimi yapılmış partide ödeme tipi / miktar değişmiş olabilir
    If Not yeni Then UretimYenidenHesapla partiNo
    BidonlariEsle partiNo
    DBEngine.CommitTrans: islemde = False

    frm!txtPartiNo = partiNo
    frm!txtSezon = SezonBul(frm!txtTarih)
    frm!txtDurum = PartiDurumuNo(partiNo)
    frm!lstPartiler.Requery
    If Soru(IIf(guncelle, "Parti güncellendi.", "Parti " & frm!txtPartiKodu & " kaydedildi.") & vbCrLf & vbCrLf & _
            "Mal Kabul Fişi yazdırılsın mı?") Then MalKabulFisiAc partiNo, "yazdir"
    Exit Sub
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "Kayıt sırasında hata: " & Err.Description, vbExclamation
End Sub

Public Function PartiDurumuNo(ByVal partiNo As Long) As String
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT SikimTarihi, CikanYag_Kg, TeslimTarihi FROM tbl_Parti WHERE PartiNo=" & partiNo, dbOpenSnapshot)
    If Not rs.EOF Then PartiDurumuNo = PartiDurumu(rs!SikimTarihi, rs!CikanYag_Kg, rs!TeslimTarihi)
    rs.Close
End Function

' ---------------------------------------------------------------
' Liste
' ---------------------------------------------------------------
Public Function MalKabulListeFiltrele()
    On Error Resume Next
    Dim frm As Form, w As String
    Set frm = F()
    w = " WHERE Sezon=" & SqlMetin(AktifSezon())
    Select Case Nz(frm!cmbListeFiltre, "")
        Case "Üretim Bekleyenler": w = w & " AND Durum='" & DURUM_KABUL & "'"
        Case "Teslim Bekleyenler": w = w & " AND Durum<>'" & DURUM_TESLIM & "'"
        Case "Tüm Sezonlar": w = ""
    End Select
    frm!lstPartiler.RowSource = "SELECT * FROM qry_MalKabulListe" & w & " ORDER BY TarihHam DESC, PartiSira DESC;"
End Function

Public Function MalKabulListedenSec()
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
    frm!txtSezon = Nz(rs!sezon, SezonBul(rs!tarih))
    frm!cmbMusteri = rs!musteriID
    frm!cmbCinsi = rs!Cinsi
    frm!cmbOzellik = rs!Ozellik
    frm!cmbGirisPersonel = rs!GirisPersonelID
    frm!cmbOdemeTipi = rs!odemeTipi
    frm!txtBidonDetay = rs!BidonDetay
    frm!txtKabulAciklama = rs!KabulAciklama
    frm!txtBrut = Nz(rs!brutKg, rs!ZeytinMiktari_Kg)
    frm!txtDara = Nz(rs!daraKg, 0)
    frm!txtNet = rs!ZeytinMiktari_Kg
    frm!txtDurum = PartiDurumu(rs!SikimTarihi, rs!CikanYag_Kg, rs!TeslimTarihi)
    rs.Close
    MalKabulMusteriSecildi
    Exit Function
Hata:
    Mesaj "Parti açılamadı: " & Err.Description, vbExclamation
End Function

' ---------------------------------------------------------------
' Sil
' ---------------------------------------------------------------
Public Function MalKabulSil()
    Dim frm As Form, id As Long
    Set frm = F()
    If Bos(frm!txtPartiNo) Then
        Mesaj "Silmek için önce listeden bir parti seçin (çift tıklayın).", vbExclamation
        Exit Function
    End If
    id = CLng(frm!txtPartiNo)
    If Not IsNull(DLookup("TeslimTarihi", "tbl_Parti", "PartiNo=" & id)) Then
        Mesaj "Yağı teslim edilmiş parti silinemez. Önce teslimi geri alın.", vbExclamation: Exit Function
    End If
    If Not IsNull(DLookup("SikimTarihi", "tbl_Parti", "PartiNo=" & id)) Then
        Mesaj "Üretimi yapılmış parti silinemez. Önce Üretim ekranından üretimi geri alın.", vbExclamation: Exit Function
    End If
    If Not Soru("Parti " & frm!txtPartiKodu & " silinsin mi?") Then Exit Function
    If PartiSilIslem(id) Then
        MalKabulYeni
        Mesaj "Parti silindi."
    End If
End Function

' ---------------------------------------------------------------
' Mal Kabul Fişi (A5)
' ---------------------------------------------------------------
Public Function MalKabulFisi(ByVal cikti As String)
    Dim frm As Form
    Set frm = F()
    If Bos(frm!txtPartiNo) Then
        Mesaj "Önce partiyi kaydedin veya listeden seçin.", vbExclamation: Exit Function
    End If
    MalKabulFisiAc CLng(frm!txtPartiNo), cikti
End Function
