SELECT c.CariHareketID, c.MusteriID, m.AdSoyad AS Musteri, m.Telefon, m.BankaAdi, m.IBAN AS IbanNo, c.Tarih, p.PartiKodu, c.HareketTuru, c.Aciklama, CDbl(Nz(c.Borc,0)) AS Borc, CDbl(Nz(c.Alacak,0)) AS Alacak
FROM (tbl_CariHareket AS c INNER JOIN tbl_Musteri AS m ON c.MusteriID=m.MusteriID) LEFT JOIN tbl_Parti AS p ON c.PartiNo=p.PartiNo
WHERE Nz(c.IptalMi,False)=False;
