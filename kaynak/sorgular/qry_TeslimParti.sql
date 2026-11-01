SELECT p.PartiNo, p.PartiKodu AS Parti, m.AdSoyad AS Müşteri, Format(p.Tarih,"dd.mm.yyyy") AS Tarih, IIf(p.TeslimTarihi Is Null,"Bekliyor","Teslim edildi") AS Durum
FROM tbl_Parti AS p LEFT JOIN tbl_Musteri AS m ON p.MusteriID=m.MusteriID
ORDER BY IIf(p.TeslimTarihi Is Null,0,1), p.Tarih DESC , p.PartiKodu DESC;
