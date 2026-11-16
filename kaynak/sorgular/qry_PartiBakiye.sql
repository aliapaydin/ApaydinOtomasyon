SELECT c.PartiNo, Sum(Nz(c.Borc,0)-Nz(c.Alacak,0)) AS Kalan
FROM tbl_CariHareket AS c
WHERE c.PartiNo Is Not Null AND Nz(c.IptalMi,False)=False
GROUP BY c.PartiNo;