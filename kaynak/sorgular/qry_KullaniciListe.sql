SELECT KullaniciID, KullaniciAdi AS [Kullanıcı Adı], AdSoyad AS [Ad Soyad], IIf(Yonetici,"Yönetici","Kullanıcı") AS Yetki, IIf(Aktif,"Evet","Hayır") AS Aktif, Format(SonGiris,"dd.mm.yyyy hh:nn") AS [Son Giriş]
FROM tbl_Kullanici
ORDER BY Yonetici, AdSoyad;
