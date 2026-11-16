Option Compare Database
Option Explicit

' =============================================================
' PARTİ ORTAK İŞLEMLERİ
' -------------------------------------------------------------
' Bir parti üç aşamadan geçer:
'   1. Mal Kabul  (frm_MalKabul)  : müşteri, zeytin brüt/dara/net,
'                                   bidonlar, ödeme tipi
'   2. Üretim     (frm_Uretim)    : sıkım tarihi, çıkan yağ,
'                                   hakyağı / sıkım ücreti
'   3. Yağ Teslim (frm_CicerTeslim): teslim ve tahsilat
'
' Parti No (PartiKodu) sezon içinde önerilir, elle değiştirilebilir.
' Yağı teslim edilmiş partinin numarası aynı sezonda yeniden
' kullanılabilir. Teslim geri alınırken numara başka bir açık partide
' kullanılıyorsa eski parti "105/1", "105/2" ... olarak yeniden adlanır.
'
' Sıkım ücreti = çıkan yağ x hakyağı oranı x 1 kg yağ fiyatı
'   (oran ve fiyat Ayarlar'dan gelir, partiye kopyalanır;
'    Ayarlar değişse de eski partiler etkilenmez)
'
' Otomatik kayıtlar:
'   Üretim (KaynakTipi PARTI):
'     Sıkım Ücreti -> Cari BORÇ (sıkım ücreti)
'     Hakyağı      -> Stok GİRİŞ (Hak Yağı ürünü, yağ x oran)
'   Teslim (KaynakTipi TESLIM):
'     Peşin / Kısmi tahsilat -> Cari ALACAK + Kasa/Banka TAHSİLAT
' =============================================================

Public Const DURUM_KABUL As String = "Kabul"
Public Const DURUM_URETILDI As String = "Üretildi"
Public Const DURUM_TESLIM As String = "Teslim Edildi"

' ---------------------------------------------------------------
' Parti numarası
' ---------------------------------------------------------------
' Sezonda açık (teslim edilmemiş) bir partide kullanılıyor mu?
Public Function PartiNoMesgul(ByVal sezon As String, ByVal kod As String, Optional ByVal haricPartiNo As Long = 0) As Boolean
    PartiNoMesgul = DCount("*", "tbl_Parti", "Sezon=" & SqlMetin(sezon) & " AND PartiKodu=" & SqlMetin(kod) & _
        " AND TeslimTarihi Is Null AND PartiNo<>" & haricPartiNo) > 0
End Function

' Önerilen numara: sezonda açık (teslim edilmemiş) partilerin en büyük
' numarasının bir fazlası. Açık parti yoksa 1'den başlar; böylece teslim
' edilen partilerin numaraları sırası gelince yeniden kullanılır.
Public Function PartiNoOner(ByVal sezon As String) As String
    Dim rs As DAO.Recordset, enBuyuk As Long, n As Long
    Set rs = CurrentDb.OpenRecordset("SELECT PartiKodu FROM tbl_Parti WHERE Sezon=" & SqlMetin(sezon) & _
        " AND TeslimTarihi Is Null", dbOpenSnapshot)
    Do Until rs.EOF
        n = Val(Nz(rs!PartiKodu, "0"))
        If n > enBuyuk Then enBuyuk = n
        rs.MoveNext
    Loop
    rs.Close
    n = enBuyuk + 1
    Do While PartiNoMesgul(sezon, CStr(n))
        n = n + 1
    Loop
    PartiNoOner = CStr(n)
End Function

' Teslimi geri alınan partinin numarası doluysa 105 -> 105/1 yapar.
' Yeni numarayı döndürür (değişmediyse eskisini).
Public Function PartiNoCakismaCoz(ByVal partiNo As Long) As String
    Dim sezon As String, kod As String, kok As String, k As Long, yeni As String
    sezon = Nz(DLookup("Sezon", "tbl_Parti", "PartiNo=" & partiNo), "")
    kod = Nz(DLookup("PartiKodu", "tbl_Parti", "PartiNo=" & partiNo), "")
    PartiNoCakismaCoz = kod
    If Not PartiNoMesgul(sezon, kod, partiNo) Then Exit Function
    kok = kod
    If InStr(kok, "/") > 0 Then kok = Left$(kok, InStr(kok, "/") - 1)
    k = 1
    Do
        yeni = kok & "/" & k
        If DCount("*", "tbl_Parti", "Sezon=" & SqlMetin(sezon) & " AND PartiKodu=" & SqlMetin(yeni) & " AND PartiNo<>" & partiNo) = 0 Then Exit Do
        k = k + 1
    Loop
    CurrentDb.Execute "UPDATE tbl_Parti SET PartiKodu=" & SqlMetin(yeni) & " WHERE PartiNo=" & partiNo, dbFailOnError
    PartiNoCakismaCoz = yeni
End Function

' Ekranda / fişte görünen parti adı
Public Function PartiAdi(ByVal partiNo As Long) As String
    PartiAdi = "Parti " & Nz(DLookup("PartiKodu", "tbl_Parti", "PartiNo=" & partiNo), partiNo)
End Function

' ---------------------------------------------------------------
' Durum ve hesaplar
' ---------------------------------------------------------------
Public Function PartiDurumu(ByVal sikimTarihi As Variant, ByVal cikanYag As Variant, ByVal teslimTarihi As Variant) As String
    If Not IsNull(teslimTarihi) Then
        PartiDurumu = DURUM_TESLIM
    ElseIf Not IsNull(sikimTarihi) And Nz(cikanYag, 0) > 0 Then
        PartiDurumu = DURUM_URETILDI
    Else
        PartiDurumu = DURUM_KABUL
    End If
End Function

Public Function SikimUcreti(ByVal yag As Double, ByVal oran As Double, ByVal fiyat As Double) As Currency
    SikimUcreti = Round(yag * oran / 100 * fiyat, 2)
End Function

Public Function HakyagiKg(ByVal yag As Double, ByVal oran As Double) As Double
    HakyagiKg = Round(yag * oran / 100, 2)
End Function

Public Function VarsayilanOran() As Double
    VarsayilanOran = AyarSayi("HakyagiOrani", 12)
End Function

Public Function VarsayilanYagFiyati() As Double
    VarsayilanYagFiyati = AyarSayi("YagFiyati", 200)
End Function

' Partinin cari bakiyesi (+ ise müşteri bu partiden borçlu)
Public Function PartiKalanBorc(ByVal partiNo As Long) As Currency
    PartiKalanBorc = Nz(DSum("Nz(Borc,0)-Nz(Alacak,0)", "tbl_CariHareket", _
        "PartiNo=" & partiNo & " AND Nz(IptalMi,False)=False"), 0)
End Function

' ---------------------------------------------------------------
' Otomatik kayıtlar
' ---------------------------------------------------------------
' Üretim kayıtları (sıkım ücreti borcu / hakyağı stok girişi)
Public Sub PartiMuhasebeYaz(ByVal partiNo As Long)
    Dim rs As DAO.Recordset, acik As String, tutar As Currency, tarih As Date
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Parti WHERE PartiNo=" & partiNo, dbOpenSnapshot)
    OtomatikKayitlariSil KAYNAK_PARTI, partiNo
    If rs.EOF Then rs.Close: Exit Sub
    If IsNull(rs!musteriID) Or IsNull(rs!SikimTarihi) Or Nz(rs!CikanYag_Kg, 0) <= 0 Then rs.Close: Exit Sub
    acik = "Parti " & rs!PartiKodu
    tarih = rs!SikimTarihi

    If UcretliMi(rs!odemeTipi) Then
        tutar = Nz(rs!TLTutari, 0)
        If tutar > 0 Then
            If Nz(rs!YagFiyati, 0) > 0 Then
                acik = acik & " sıkım ücreti (" & Nz(rs!CikanYag_Kg, 0) & " kg yağ x %" & Nz(rs!HakYagiOrani_Yuzde, 0) & _
                    " x " & Format$(Nz(rs!YagFiyati, 0), "#,##0.00") & " TL)"
            Else
                acik = acik & " sıkım ücreti"
            End If
            CariYaz tarih, rs!musteriID, partiNo, "Sıkım Ücreti", acik, tutar, 0, KAYNAK_PARTI, partiNo, True
        End If
    ElseIf HakyagiMi(rs!odemeTipi) Then
        If Nz(rs!HakYagi_Kg, 0) > 0 Then
            StokYaz tarih, HakYagiUrunID(), rs!musteriID, partiNo, "Hakyağı Giriş", _
                acik & " %" & Nz(rs!HakYagiOrani_Yuzde, 0) & " hakyağı (" & Nz(rs!CikanYag_Kg, 0) & " kg yağdan)", _
                rs!HakYagi_Kg, 0, KAYNAK_PARTI, partiNo, True
        End If
    End If
    rs.Close
End Sub

' Teslimde alınan tahsilat (peşin / kısmi)
Public Sub TeslimMuhasebeYaz(ByVal partiNo As Long)
    Dim rs As DAO.Recordset, tutar As Currency, tarih As Date, acik As String, yontem As String, kasa As Long
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl_Parti WHERE PartiNo=" & partiNo, dbOpenSnapshot)
    OtomatikKayitlariSil KAYNAK_TESLIM, partiNo
    If rs.EOF Then rs.Close: Exit Sub
    tutar = Nz(rs!TahsilatTutari, 0)
    If IsNull(rs!musteriID) Or tutar <= 0 Then rs.Close: Exit Sub
    If Nz(rs!TahsilatTuru, "Yok") = "Yok" Then rs.Close: Exit Sub
    tarih = Nz(rs!TahsilatTarihi, Nz(rs!TeslimTarihi, rs!tarih))
    yontem = Nz(rs!TahsilatYontemi, "Nakit")
    kasa = Nz(rs!TahsilatKasaID, 0)
    acik = "Parti " & rs!PartiKodu & " sıkım ücreti tahsilatı (" & LCase$(Nz(rs!TahsilatTuru, "")) & ", " & yontem & ")"
    CariYaz tarih, rs!musteriID, partiNo, IIf(rs!TahsilatTuru = "Kısmi", "Kısmi Tahsilat", "Peşin Tahsilat"), _
        acik, 0, tutar, KAYNAK_TESLIM, partiNo, True
    KasaYaz tarih, rs!musteriID, partiNo, "Tahsilat", yontem, acik, tutar, 0, KAYNAK_TESLIM, partiNo, True, 0, kasa
    rs.Close
End Sub

' ---------------------------------------------------------------
' Bidonlar
' "11, 12 13;14" veya "11-14" gibi yazılan numaraları ayırır,
' tekrarları atar. Aralık (11-14) en fazla 200 bidonluk olabilir.
' Sonuç: "11, 12, 13, 14"
' ---------------------------------------------------------------
Public Function BidonlariAyristir(ByVal metin As Variant) As String
    Dim s As String, p() As String, i As Long, b As String, a As Long, z As Long, n As Long
    Dim goruldu As Object, sonuc As String, tire As Long
    Set goruldu = CreateObject("Scripting.Dictionary")
    s = Nz(metin, "")
    s = Replace(s, vbCrLf, ","): s = Replace(s, vbLf, ","): s = Replace(s, vbTab, ",")
    s = Replace(s, ";", ","): s = Replace(s, "+", ",")
    ' "11 - 14" -> "11-14"
    Do While InStr(s, " -") > 0: s = Replace(s, " -", "-"): Loop
    Do While InStr(s, "- ") > 0: s = Replace(s, "- ", "-"): Loop
    s = Replace(s, " ", ",")
    p = Split(s, ",")
    For i = LBound(p) To UBound(p)
        b = Trim$(p(i))
        If b <> "" Then
            tire = InStr(2, b, "-")
            If tire > 0 Then
                If IsNumeric(Left$(b, tire - 1)) And IsNumeric(Mid$(b, tire + 1)) Then
                    a = CLng(Left$(b, tire - 1)): z = CLng(Mid$(b, tire + 1))
                    If z >= a And z - a < 200 Then
                        For n = a To z
                            If Not goruldu.Exists(CStr(n)) Then goruldu(CStr(n)) = True: sonuc = sonuc & ", " & n
                        Next
                        b = ""
                    End If
                End If
            End If
            If b <> "" Then
                b = Left$(UCase$(b), 30)
                If Not goruldu.Exists(b) Then goruldu(b) = True: sonuc = sonuc & ", " & b
            End If
        End If
    Next
    If sonuc <> "" Then sonuc = Mid$(sonuc, 3)
    BidonlariAyristir = sonuc
End Function

' Sezonda açık başka bir partide kullanılan bidonlar ("12, 14" gibi)
Public Function BidonCakismasi(ByVal bidonlar As String, ByVal sezon As String, ByVal partiNo As Long) As String
    Dim p() As String, i As Long, b As String, sonuc As String, rs As DAO.Recordset, acik As Object
    If bidonlar = "" Then Exit Function
    Set acik = CreateObject("Scripting.Dictionary")
    Set rs = CurrentDb.OpenRecordset("SELECT x.BidonNo FROM tbl_PartiBidon AS x INNER JOIN tbl_Parti AS y ON x.PartiNo=y.PartiNo " & _
        "WHERE y.Sezon=" & SqlMetin(sezon) & " AND y.TeslimTarihi Is Null AND y.PartiNo<>" & partiNo, dbOpenSnapshot)
    Do Until rs.EOF
        acik(UCase$(Nz(rs!BidonNo, ""))) = True
        rs.MoveNext
    Loop
    rs.Close
    p = Split(bidonlar, ",")
    For i = 0 To UBound(p)
        b = UCase$(Trim$(p(i)))
        If b <> "" Then
            If acik.Exists(b) Then sonuc = sonuc & ", " & b
        End If
    Next
    If sonuc <> "" Then sonuc = Mid$(sonuc, 3)
    BidonCakismasi = sonuc
End Function

' tbl_Parti.BidonDetay -> tbl_PartiBidon. Teslim tarihi varsa teslim edildi işaretlenir.
Public Sub BidonlariEsle(ByVal partiNo As Long)
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim metin As String, parcalar() As String, yeni As String, eski As String, i As Long
    Dim teslim As Variant

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT BidonDetay, TeslimTarihi FROM tbl_Parti WHERE PartiNo=" & partiNo, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Sub
    metin = BidonlariAyristir(rs!BidonDetay)
    teslim = rs!teslimTarihi
    rs.Close

    If metin <> "" Then
        parcalar = Split(metin, ", ")
        For i = 0 To UBound(parcalar)
            yeni = yeni & "|" & parcalar(i)
        Next
    End If

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
                rs!TeslimEdildi = False
                rs.Update
            Next
            rs.Close
        End If
    End If

    If IsNull(teslim) Then
        db.Execute "UPDATE tbl_PartiBidon SET TeslimEdildi=False, TeslimTarihi=Null WHERE PartiNo=" & partiNo, dbFailOnError
    Else
        db.Execute "UPDATE tbl_PartiBidon SET TeslimEdildi=True, TeslimTarihi=" & SqlTarih(teslim) & _
            " WHERE PartiNo=" & partiNo, dbFailOnError
    End If
End Sub

' ---------------------------------------------------------------
' Parti silme (Mal Kabul ekranından)
' ---------------------------------------------------------------
Public Function PartiSilIslem(ByVal partiNo As Long) As Boolean
    On Error GoTo Hata
    Dim n As Long, elle As String, islemde As Boolean
    elle = "PartiNo=" & partiNo & " AND (OtomatikMi=False OR OtomatikMi Is Null)"
    n = DCount("*", "tbl_CariHareket", elle) + DCount("*", "tbl_KasaHareket", elle) + DCount("*", "tbl_StokHareket", elle)
    If n > 0 Then
        Mesaj "Bu partiye bağlı elle girilmiş " & n & " cari / kasa / stok hareketi var (ör. havale)." & vbCrLf & _
              "Önce o hareketleri silin.", vbExclamation
        Exit Function
    End If
    DBEngine.BeginTrans: islemde = True
    OtomatikKayitlariSil KAYNAK_PARTI, partiNo
    OtomatikKayitlariSil KAYNAK_TESLIM, partiNo
    CurrentDb.Execute "DELETE FROM tbl_PartiBidon WHERE PartiNo=" & partiNo, dbFailOnError
    CurrentDb.Execute "DELETE FROM tbl_Parti WHERE PartiNo=" & partiNo, dbFailOnError
    DBEngine.CommitTrans: islemde = False
    PartiSilIslem = True
    Exit Function
Hata:
    If islemde Then DBEngine.Rollback
    Mesaj "Silme sırasında hata: " & Err.Description, vbExclamation
End Function

' ---------------------------------------------------------------
' BAKIM: tüm partilerin otomatik kayıtlarını yeniden yazar
' (tutarlar yeniden hesaplanmaz, kayıtlı değerler kullanılır)
' ---------------------------------------------------------------
Public Sub TumPartiKayitlariniYenidenYaz()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT PartiNo FROM tbl_Parti", dbOpenSnapshot)
    Do Until rs.EOF
        PartiMuhasebeYaz rs!partiNo
        TeslimMuhasebeYaz rs!partiNo
        BidonlariEsle rs!partiNo
        rs.MoveNext
    Loop
    rs.Close
End Sub
