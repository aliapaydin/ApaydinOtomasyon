SELECT p.PartiNo, p.PartiKodu, p.Tarih, p.MusteriID, m.AdSoyad AS Musteri, CDbl(Nz(p.ZeytinMiktari_Kg,0)) AS Zeytin, CDbl(Nz(p.CikanYag_Kg,0)) AS Yag, IIf(Nz(p.ZeytinMiktari_Kg,0)>0,Nz(p.CikanYag_Kg,0)/p.ZeytinMiktari_Kg*100,Null) AS Randiman, p.OdemeTipi, CDbl(Nz(p.HakYagi_Kg,0)) AS HakYag, CDbl(Nz(p.TLTutari,0)) AS Ucret
FROM tbl_Parti AS p LEFT JOIN tbl_Musteri AS m ON p.MusteriID=m.MusteriID;
