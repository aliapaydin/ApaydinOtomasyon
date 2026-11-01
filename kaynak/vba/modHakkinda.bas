Option Compare Database
Option Explicit

' =============================================================
' HAKKINDA   (form: frm_Hakkinda)
' -------------------------------------------------------------
' Ürün, sürüm ve geliştirici bilgileri. Bu bilgiler modOrtak'taki
' sabitlerden gelir ve programdan değiştirilemez.
' =============================================================

Private Const FORM_ADI As String = "frm_Hakkinda"

Public Function HakkindaYukle()
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_ADI)
    frm!lblUrun.Caption = BuyukHarf(URUN_ADI)
    frm!lblSurum.Caption = "Sürüm " & URUN_SURUM & "   •   Lisans: " & FirmaAdi()
    frm!lblGelistirici.Caption = GELISTIRICI
    frm!lblEposta.Caption = GELISTIRICI_EPOSTA
    frm!lblTelefon.Caption = GELISTIRICI_TELEFON
    frm!lblGithub.Caption = GELISTIRICI_GITHUB
    frm!lblTelif.Caption = "© " & Year(Date) & " " & GELISTIRICI & ". Tüm hakları saklıdır."
End Function

Public Function HakkindaGithubAc()
    On Error Resume Next
    Application.FollowHyperlink "https://" & GELISTIRICI_GITHUB
End Function

Public Function HakkindaEpostaAc()
    On Error Resume Next
    Application.FollowHyperlink "mailto:" & GELISTIRICI_EPOSTA
End Function