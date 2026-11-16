Option Compare Database
Option Explicit

' =============================================================
' KAYNAK DIŞA AKTARMA
' -------------------------------------------------------------
' Program içindeki form, rapor, modül ve sorguları program
' klasöründeki "kaynak" klasörüne UTF-8 metin olarak yazar
' (git ile sürüm takibi için). Ayarlar ekranından çalıştırılır.
' VBA modülleri aktarılmaz; onların asıl kaynağı kaynak\vba klasörüdür.
' Aynı işi program kapalıyken araclar\KaynakDisaAktar.vbs yapar;
' kaynaktan yeni program oluşturmak için araclar\Kur.vbs kullanılır.
' =============================================================

Public Function KaynakDisaAktar(Optional ByVal hedef As String = "") As String
    On Error GoTo Hata
    Dim ao As AccessObject, qd As DAO.QueryDef, n As Long, tmp As String
    If hedef = "" Then hedef = CurrentProject.Path & "\kaynak"
    KlasorOlustur hedef
    KlasorOlustur hedef & "\formlar"
    KlasorOlustur hedef & "\raporlar"
    KlasorOlustur hedef & "\vba"
    KlasorOlustur hedef & "\sorgular"
    tmp = Environ$("TEMP") & "\apaydin_kaynak.tmp"
    DoCmd.Hourglass True
    For Each ao In CurrentProject.AllForms
        NesneYaz acForm, ao.Name, hedef & "\formlar\" & ao.Name & ".txt", tmp
        n = n + 1
    Next
    For Each ao In CurrentProject.AllReports
        NesneYaz acReport, ao.Name, hedef & "\raporlar\" & ao.Name & ".txt", tmp
        n = n + 1
    Next
    ' Modüller dışa aktarılmaz: kaynak\vba asıl kaynaktır. Kur.vbs modülleri
    ' yüklerken Türkçe harfleri ChrW(...) biçimine çevirdiği için programdaki
    ' kod, kaynaktaki okunabilir metinle aynı değildir.
    For Each qd In CurrentDb.QueryDefs
        If Left$(qd.Name, 1) <> "~" Then
            Utf8Yaz hedef & "\sorgular\" & qd.Name & ".sql", qd.sql
            n = n + 1
        End If
    Next
    DoCmd.Hourglass False
    KaynakDisaAktar = hedef & "  (" & n & " nesne)"
    Exit Function
Hata:
    DoCmd.Hourglass False
    Mesaj "Kaynak dışa aktarılamadı: " & Err.Description, vbExclamation
End Function

Private Sub KlasorOlustur(ByVal k As String)
    If Dir(k, vbDirectory) = "" Then MkDir k
End Sub

Private Sub NesneYaz(ByVal tur As AcObjectType, ByVal ad As String, ByVal hedef As String, ByVal tmp As String)
    If Dir(tmp) <> "" Then Kill tmp
    Application.SaveAsText tur, ad, tmp
    Utf8Yaz hedef, MetinOku(tmp, (tur = acModule))
End Sub

' SaveAsText formları UTF-16, modülleri ANSI (Türkçe) yazar
Private Function MetinOku(ByVal yol As String, ByVal ansi As Boolean) As String
    Dim st As Object, b() As Byte
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open: st.LoadFromFile yol
    b = st.Read(2)
    st.Close
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "windows-1254"
    If UBound(b) >= 1 Then
        If b(0) = &HFF And b(1) = &HFE Then st.Charset = "unicode"
        If b(0) = &HEF And b(1) = &HBB Then st.Charset = "utf-8"
    End If
    st.Open: st.LoadFromFile yol
    MetinOku = st.ReadText
    st.Close
End Function

' BOM'suz UTF-8
Private Sub Utf8Yaz(ByVal yol As String, ByVal metin As String)
    Dim st As Object, b As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText metin
    st.Position = 3
    Set b = CreateObject("ADODB.Stream")
    b.Type = 1: b.Open
    st.CopyTo b
    b.SaveToFile yol, 2
    b.Close: st.Close
End Sub
