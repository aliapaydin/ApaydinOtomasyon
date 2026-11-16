SELECT k.KasaID, k.KasaAdi AS [Hesap Adı], Nz(k.KasaTuru,"Nakit") AS Tür, k.BankaAdi AS Banka, k.IBAN, IIf(k.Aktif,"Aktif","Pasif") AS Durum
FROM tbl_Kasa AS k
ORDER BY IIf(Nz(k.KasaTuru,"Nakit")="Nakit",0,1), k.KasaAdi;