SELECT m.MusteriID, m.TCKN_VKN AS [TCKN/VKN], m.AdSoyad AS [Ad Soyad / Ünvan], m.Telefon, m.BankaAdi AS Banka, m.IBAN, m.CariTipi AS [Cari Tipi], Format(Nz(m.AcilisBakiyesi,0) + Nz(DSum("Nz(Borc,0)-Nz(Alacak,0)","tbl_CariHareket","Nz(IptalMi,False)=False AND MusteriID=" & m.MusteriID),0),"#,##0.00") AS Bakiye
FROM tbl_Musteri AS m
ORDER BY m.AdSoyad;
