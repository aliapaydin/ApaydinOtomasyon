SELECT k.KasaHareketID, k.Tarih, k.MusteriID, m.AdSoyad AS Musteri, p.PartiKodu, k.HareketTuru, k.OdemeYontemi, k.Aciklama, CDbl(Nz(k.Tahsilat,0)) AS Tahsilat, CDbl(Nz(k.Odeme,0)) AS Odeme
FROM (tbl_KasaHareket AS k LEFT JOIN tbl_Musteri AS m ON k.MusteriID=m.MusteriID) LEFT JOIN tbl_Parti AS p ON k.PartiNo=p.PartiNo
WHERE Nz(k.IptalMi,False)=False;
