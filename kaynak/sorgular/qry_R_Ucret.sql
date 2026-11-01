SELECT p.PartiNo, p.PartiKodu, p.Tarih, p.MusteriID, m.AdSoyad AS Musteri, CDbl(Nz(p.ZeytinMiktari_Kg,0)) AS Zeytin, CDbl(Nz(p.SimSikimUcreti,0)) AS BirimUcret, CDbl(Nz(p.TLTutari,0)) AS Ucret, IIf(Nz(p.PesinTahsilEdildi,False),"Tahsil Edilen","Tahsil Edilmeyen") AS Durum, p.TahsilatTarihi
FROM tbl_Parti AS p LEFT JOIN tbl_Musteri AS m ON p.MusteriID=m.MusteriID
WHERE p.OdemeTipi Like "S*";
