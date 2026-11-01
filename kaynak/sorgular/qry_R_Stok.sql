SELECT s.StokHareketID, s.UrunID, u.UrunAdi AS Urun, u.Birim, s.Tarih, s.MusteriID, m.AdSoyad AS Musteri, p.PartiKodu, s.HareketTuru, s.Aciklama, CDbl(Nz(s.Giris,0)) AS Giris, CDbl(Nz(s.Cikis,0)) AS Cikis
FROM ((tbl_StokHareket AS s INNER JOIN tbl_Urun AS u ON s.UrunID=u.UrunID) LEFT JOIN tbl_Musteri AS m ON s.MusteriID=m.MusteriID) LEFT JOIN tbl_Parti AS p ON s.PartiNo=p.PartiNo
WHERE Nz(s.IptalMi,False)=False;
