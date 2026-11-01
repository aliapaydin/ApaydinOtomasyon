SELECT b.PartiBidonID, p.MusteriID, m.AdSoyad AS Musteri, m.Telefon, p.PartiKodu, p.Tarih, b.BidonNo, IIf(Nz(b.TeslimEdildi,False),"Teslim Edildi","Bekliyor") AS Durum, b.TeslimTarihi
FROM (tbl_PartiBidon AS b INNER JOIN tbl_Parti AS p ON b.PartiNo=p.PartiNo) LEFT JOIN tbl_Musteri AS m ON p.MusteriID=m.MusteriID;
