SELECT k.KullaniciID, k.KullaniciAdi AS [Kullanıcı Adı], k.AdSoyad AS [Ad Soyad], IIf(k.Yonetici,"Yönetici","Kullanıcı") AS Yetki, IIf(k.Aktif,"Evet","Hayır") AS [Aktif mi], Format(k.SonGiris,"dd.mm.yyyy hh:nn") AS [Son Giriş]
FROM tbl_Kullanici AS k
ORDER BY k.Yonetici, k.AdSoyad;
