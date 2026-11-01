SELECT a.IslemID, Format(a.Tarih,"dd.mm.yyyy") AS Tarih, a.IslemTuru AS İşlem, a.MusteriID, m.AdSoyad AS Cari, a.UrunID, u.UrunAdi AS Ürün, a.Miktar, Format(a.BirimFiyat,"#,##0.00") AS [Birim Fiyat], Format(Nz(a.ToplamTutar,0),"#,##0.00") AS Toplam, IIf(Nz(a.PesinOdendi,False),"Evet","Hayır") AS Peşin, a.Aciklama AS Açıklama
FROM (tbl_AlimSatim AS a INNER JOIN tbl_Musteri AS m ON a.MusteriID=m.MusteriID) INNER JOIN tbl_Urun AS u ON a.UrunID=u.UrunID
WHERE Nz(a.IptalMi,False)=False
ORDER BY a.Tarih DESC , a.IslemID DESC;
