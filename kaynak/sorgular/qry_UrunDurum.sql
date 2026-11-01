SELECT u.UrunID, u.UrunKodu AS Kod, u.UrunAdi AS [Ürün Adı], u.Birim, Nz(u.AcilisStok,0) AS Açılış, Format(u.AlisFiyati,"#,##0.00") AS [Alış Fiyatı], Format(u.SatisFiyati,"#,##0.00") AS [Satış Fiyatı], u.MinimumStok AS [Min Stok], Nz(u.AcilisStok,0) + Nz(DSum("Nz(Giris,0)-Nz(Cikis,0)","tbl_StokHareket","Nz(IptalMi,False)=False AND UrunID=" & u.UrunID),0) AS [Mevcut Stok], IIf(Nz(u.Aktif,True),"Evet","Hayır") AS Aktif
FROM tbl_Urun AS u
ORDER BY u.UrunAdi;
