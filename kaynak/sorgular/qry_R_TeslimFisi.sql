SELECT p.*, m.AdSoyad AS Musteri, m.Telefon, m.Adres AS MusteriAdres, m.TCKN_VKN, gp.AdSoyad AS GirisPersoneli, tp.AdSoyad AS TeslimEden
FROM ((tbl_Parti AS p LEFT JOIN tbl_Musteri AS m ON p.MusteriID=m.MusteriID) LEFT JOIN tbl_Personel AS gp ON p.GirisPersonelID=gp.PersonelID) LEFT JOIN tbl_Personel AS tp ON p.TeslimPersonelID=tp.PersonelID;
