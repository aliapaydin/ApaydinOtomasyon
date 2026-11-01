SELECT b.PartiBidonID, p.MusteriID, m.AdSoyad AS Müşteri, p.PartiKodu AS Parti, b.BidonNo AS Bidon, IIf(Nz(b.TeslimEdildi,False),"Evet","Hayır") AS Teslim
FROM (tbl_PartiBidon AS b INNER JOIN tbl_Parti AS p ON b.PartiNo=p.PartiNo) LEFT JOIN tbl_Musteri AS m ON p.MusteriID=m.MusteriID
ORDER BY p.Tarih DESC , p.PartiKodu, b.PartiBidonID;
