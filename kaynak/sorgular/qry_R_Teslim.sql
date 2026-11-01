SELECT p.PartiNo, p.PartiKodu, p.Tarih AS PartiTarihi, p.MusteriID, m.AdSoyad AS Musteri, p.TeslimTarihi, p.TeslimAlan, per.AdSoyad AS TeslimEden, CDbl(Nz(p.CikanYag_Kg,0)) AS Yag, CDbl(Nz(p.HakYagi_Kg,0)) AS HakYag, CDbl(Nz(p.TeslimYag_Kg,0)) AS TeslimYag, p.OdemeTipi
FROM (tbl_Parti AS p LEFT JOIN tbl_Musteri AS m ON p.MusteriID=m.MusteriID) LEFT JOIN tbl_Personel AS per ON p.TeslimPersonelID=per.PersonelID
WHERE p.TeslimTarihi Is Not Null;
