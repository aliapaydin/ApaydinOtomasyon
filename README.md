<p align="center">
  <img src="Gorseller/apaydin_logo.png" width="140" alt="Logo">
</p>

<h1 align="center">Zeytinyağı Üretim Otomasyonu</h1>

<p align="center">
  Zeytinyağı fabrikaları için <b>mal kabul, üretim, yağ teslimi, sıkım ücreti / hakyağı, cari, kasa-banka, stok ve raporlama</b> programı.<br>
  Microsoft Access (VBA) ile yazılmıştır. Program ve veri ayrı dosyalardır; program güncellenince veriler korunur.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Microsoft%20Access-VBA-A4373A?logo=microsoftaccess&logoColor=white" alt="Access">
  <img src="https://img.shields.io/badge/sürüm-3.1-6B7F2A" alt="Sürüm">
  <img src="https://img.shields.io/badge/dil-Türkçe-C9A227" alt="Türkçe">
</p>

---

## Ekran görüntüleri

| Giriş | Ana menü |
|---|---|
| ![Giriş](docs/ekran-goruntuleri/giris.png) | ![Ana menü](docs/ekran-goruntuleri/ana-menu.png) |
| **Mal Kabul** | **Üretim** |
| ![Mal Kabul](docs/ekran-goruntuleri/mal-kabul.png) | ![Üretim](docs/ekran-goruntuleri/uretim.png) |
| **Yağ Teslim** | **Kasa / Banka** |
| ![Teslim](docs/ekran-goruntuleri/teslim.png) | ![Kasa](docs/ekran-goruntuleri/kasa.png) |
| **Ayarlar** | **Kullanıma kapalı** |
| ![Ayarlar](docs/ekran-goruntuleri/ayarlar.png) | ![Kilit](docs/ekran-goruntuleri/kilit.png) |

## Sürüm 3.0'da neler değişti

- Eski tek **Parti** ekranı üçe ayrıldı: **Mal Kabul → Üretim → Yağ Teslim**.
- **Sezonlar**: her sezon 1 Eylül'de başlar, aktif sezon Ayarlar'dan seçilir. Listeler ve raporlar sezona göre süzülür.
- **Parti No** sezon içinde otomatik önerilir, elle değiştirilebilir. Teslim edilmiş partinin numarası yeniden kullanılabilir.
- **Sıkım ücreti** artık çıkan yağa göre hesaplanır: `çıkan yağ × hakyağı oranı × 1 kg yağ fiyatı`.
- **Teslimde tahsilat**: peşin veya kısmi; nakit, havale ya da kart; kasa veya banka hesabı seçilir.
- **Banka hesapları** ve partiye bağlanan **havaleler**.
- **Program / veri ayrımı**, kaynaktan kurulum betiği (`araclar\Kur.vbs`), otomatik veri yükseltme.
- Hızlandırılmış ekstre sorguları, satışta stok kontrolü, bidon numarası aralıkları (`11-14 16`), A5 mal kabul fişi.

## Özellikler

**Mal Kabul**
- Parti No (sezon içinde önerilir), tarih, müşteri, cins / özellik, **brüt – dara = net kg**, bidon numaraları, ödeme tipi (Sıkım Ücreti / Hakyağı).
- **+ Yeni** ile ekrandan çıkmadan hızlı müşteri kaydı (ad, baba adı, telefon).
- Bidonlar aralıkla yazılabilir (`11-14 16` → 11, 12, 13, 14, 16). Başka bir partide teslim bekleyen bidon varsa uyarı verilir.
- **A5 Mal Kabul Fişi**: önizleme, PDF veya yazıcı.

**Üretim**
- Sıkım tarihi, çıkan yağ, randıman ve verim (kaç kg zeytinden 1 kg yağ).
- Hakyağı oranı (varsayılan %12) ve 1 kg yağ fiyatı (varsayılan 200 TL) Ayarlar'dan gelir ve partiye kopyalanır. Ayarlar sonradan değişse de eski partiler etkilenmez.
- **Sıkım Ücreti** seçildiyse ücret müşterinin carisine borç olarak yazılır. Örnek: 240 kg × %12 × 200 TL = 5.760 TL.
- **Hakyağı** seçildiyse ücret yazılmaz; `çıkan yağ × oran` kadar yağ fabrikanın stoğuna girer.

**Yağ Teslim**
- Parti özeti: gelen zeytin, çıkan yağ, randıman, hakyağı, kalan net yağ, sıkım ücreti, tahsil edilen, kalan borç, cari bakiye.
- Tahsilat: **Tahsil Edilmedi / Peşin / Kısmi**. Kısmide tutar girilir, kalanı borç olarak durur.
- Yöntem **Nakit / Havale / Kart** ve kasa veya banka hesabı. Tahsilat seçilen hesaba işlenir.
- **Teslimi Geri Al**: parti yeniden teslim bekler. Numara bu arada başka bir partiye verildiyse eski parti `105/1`, `105/2` olarak yeniden adlandırılır.
- **A5 teslim fişi** (teslim eden / teslim alan imza alanlarıyla).

**Cari, Kasa / Banka, Stok**
- Müşteri kartı: TCKN/VKN, ad soyad, **baba adı**, telefon, adres, banka, IBAN, açılış bakiyesi.
- Cari hareketler: borç / alacak / yürüyen bakiye, müşterinin bidonları.
- Kasa ve banka hesapları ayrı bakiyelerle tutulur. Sonradan gelen bir **havale partiye bağlanabilir**; tutar partinin kalan borcundan düşer.
- Stok: ürün bazında giriş / çıkış / kalan. Hakyağı stoğu otomatik oluşur. Satışta stok yetersizse uyarı verilir.

**Raporlar (A5: önizleme, PDF, yazıcı)**

Her raporda sezon seçilebilir; tarih aralığı ve sezon rapor başlığında görünür.

| Rapor | İçerik |
|---|---|
| Üretim Raporu | Gelen zeytin, çıkan yağ, randıman, hakyağı, ücretler |
| Sıkım Ücretleri | Parti bazında yağ, ücret, tahsil edilen, kalan |
| Müşteri Bidonları | Teslim edilen / bekleyen bidonlar |
| Yağ Teslimleri | Teslim alan, teslim eden, miktarlar |
| Alacaklar ve Borçlar | Seçilen tarih itibarıyla cari bakiyeler |
| Cari Hesap Ekstresi | Devir, borç, alacak, yürüyen bakiye |
| Stok Hareketleri | Ürün bazında devir, giriş, çıkış, kalan |
| Kasa Raporu | Hesap bazında devir, tahsilat, ödeme, bakiye |

**Kullanıcılar ve güvenlik**
- Profil kartlarıyla giriş: profil seçilir, şifre yazılır.
- **Yönetici** kullanıcı ekler, siler ve şifre belirler. Ayarlar ekranını yalnızca yönetici açar.
- Şifreler tuzlu özet (hash) olarak saklanır.
- Shift ile tasarım görünümüne geçiş kapalıdır (aşağıdaki *Geliştirici modu* bölümüne bakın).

**Yedekleme**
- **Ayarlar → Yedek Al**: takılı flaş bellekler ve diskler listelenir, istenirse **Klasör Seç** ile başka bir yer seçilir. Veri dosyası seçilen yerde `ApaydinOtomasyon_Yedek` klasörüne `ApaydinVeri_yyyy-aa-gg_ss-dd-nn.accdb` adıyla yazılır. Klasör yoksa oluşturulur; eski yedekler silinmez.
- Yedek, program açıkken de güvenle alınır: dosya önce geçici klasöre kopyalanır, sıkıştırılıp doğrulanır, sonra hedefe yazılır.
- **Ayarlar → Yedekten Geri Yükle** (yalnızca yönetici): takılı flaşlar, diskler, son yedek klasörü ve seçilen klasörler taranır. **En güncel yedek önerilir**, diğerleri ayrı listede görünür. Daha eski bir yedek seçilirse uyarı verilir.
- Geri yüklemede mevcut veri silinmez; bu bilgisayardaki `Yedek` klasörüne taşınır. Program kendiliğinden yeniden başlar ve yedeği yükler.
- **Çıkışta soru**: program kapatılırken (Çıkış düğmesi ya da pencerenin X'i) **Evet, Çık / Yedek Alıp Kapat / Hayır** sorulur. Yedek Alıp Kapat yedek ekranını açar, yedek alınınca program kapanır. Son yedeğin tarihi de bu ekranda görünür.

**Kullanım kontrolü (lisans)**
- Program her açılışta ve oturum kapatılıp giriş ekranına dönüldüğünde gizli `aliapaydin/app_licenses` deposundaki `apaydin_otomasyon/lisans.txt` dosyasını, yalnızca o depoyu okuyabilen bir erişim anahtarıyla okur.
- Her kurulumun ilk açılışta üretilen bir **lisans kodu** vardır (ör. `AO-7K3M9P`); Ayarlar → Kullanım Durumu'nda ve kilit ekranında görünür.
- Dosyada `HERKES = kapali` yazılırsa tüm kurulumlar, `AO-7K3M9P = kapali` yazılırsa yalnızca o kurulum kapanır. Kapalı programda yalnızca uyarı, iletişim bilgisi, **Tekrar Dene**, **Yedek Al** ve **Çıkış** görünür; müşteri kendi verisinin yedeğini her zaman alabilir.
- İnternet yoksa son bilinen durum geçerlidir (kapatılmış bir kurulum internet kesilince açılmaz). 7 gün boyunca hiç kontrol yapılamazsa giriş durdurulur.
- **Geliştirici anahtarı:** program klasöründe geçerli bir `gelistirici.anahtar` dosyası varsa lisans kontrolü hiç yapılmaz (7 gün sınırı da yoktur). Kodda yalnızca anahtarın SHA-256 özeti durur; dosya olmadan anahtar üretilemez. `*.anahtar` dosyaları `.gitignore` ile depodan dışarıda tutulur.

**Ayarlar**
- Firma adı, telefon, adres: tüm başlıklarda ve raporlarda görünür.
- Hakyağı oranı, 1 kg yağ fiyatı, aktif sezon.
- Kasa ve banka hesapları (hesap adı, tür, banka, IBAN).
- Yedekleme: **Yedek Al**, **Yedekten Geri Yükle**, son yedeğin tarihi ve yeri.
- Yönetici araçları: **Kaynakları Dışa Aktar**, **Kayıtları Yeniden Oluştur** (tüm otomatik cari / kasa / stok kayıtlarını baştan yazar), **Geliştirici Modu**.

## Klasör yapısı

```
ApaydinOtomasyon/
├── ApaydinOtomasyon.accdb     Program (formlar, raporlar, sorgular, kod). Veri içermez.
├── Veri/
│   └── ApaydinVeri.accdb      Veri dosyası (ilk açılışta oluşur, depoya girmez)
├── ornek/
│   └── ApaydinVeri_ornek.accdb  Uydurma örnek veriler
├── Yedek/                     Kurulum, veri yükseltme ve geri yükleme öncesi kopyalar (depoya girmez)
├── Raporlar/                  PDF çıktıları (depoya girmez)
├── Gorseller/                 Logo, simge, profil görselleri
├── docs/ekran-goruntuleri/
├── kaynak/                    Programın metin kaynakları
│   ├── vba/                   VBA modülleri (.bas)
│   ├── sorgular/              Sorgu SQL'leri
│   ├── formlar/               Form tasarımları (Access SaveAsText, UTF-8)
│   ├── raporlar/              Rapor tasarımları (Access SaveAsText, UTF-8)
│   ├── tablolar.txt           Tablo yapıları (bilgi amaçlı)
│   └── kaldirilan.txt         Artık kullanılmayan kaynak dosyaları
└── araclar/
    ├── Kur.vbs                Kaynaktan programı yeniden oluşturur
    ├── KaynakDisaAktar.vbs    Programdaki formları / raporları / sorguları kaynak klasörüne yazar
    ├── sablon.accdb           Türkçe kod sayfalı boş şablon (Kur.vbs kullanır)
    └── uretici/               Formları üreten Python betikleri (isteğe bağlı)
```

## Kurulum

1. Depoyu indirin (**Code → Download ZIP**) ve bir klasöre çıkarın.
2. `ApaydinOtomasyon.accdb` dosyasını Microsoft Access (2016 / 2019 / 2021 / 365) ile açın.
3. Sarı güvenlik çubuğu çıkarsa **İçeriği Etkinleştir**'e basın.
   Kalıcı çözüm: *Dosya → Seçenekler → Güven Merkezi → Güven Merkezi Ayarları → Güvenilen Konumlar → Yeni konum ekle* ile klasörü ekleyin.
4. İlk açılışta `Veri\ApaydinVeri.accdb` yoksa program **örnek veriyi kopyalamayı** önerir. Kabul ederseniz `ornek` klasöründeki dosya `Veri` klasörüne kopyalanır ve tablolar ona bağlanır.
5. `Gorseller` klasörü program dosyasıyla aynı klasörde durmalıdır (simge ve logolar buradan okunur).

### İlk giriş

| Kullanıcı | Şifre |
|---|---|
| **Yönetici** | `1234` |
| Zeynep Arslan (örnek kullanıcı) | `1234` |

> İlk girişten sonra **Ayarlar → Yönetici Şifresi**'nden şifreyi değiştirin, **Firma Bilgileri**'ne kendi firmanızı girin ve örnek kullanıcıyı silin.

Örnek veri dosyasındaki tüm müşteri, parti ve hareketler **uydurmadır**.

### Programı kapatma / açma (lisans.txt)

Gizli `app_licenses` deposunda `apaydin_otomasyon/lisans.txt` dosyasını açıp kalem simgesiyle düzenleyin ve kaydedin. Değişiklik genellikle 1 dakika içinde geçerli olur; program bir sonraki açılışta ya da kilit ekranındaki **Tekrar Dene** ile yeni durumu okur.

```
HERKES = acik                          # tüm kurulumlar açık
AO-7K3M9P = kapali                     # yalnızca bu kurulum kapalı
AO-4XR2PD = kapali | Ödeme bekleniyor  # bu kuruluma özel mesajla kapalı
MESAJ = Programın kullanımı durdurulmuştur. Lütfen yazılımcı ile iletişime geçin.
ILETISIM = Yazılımcı: ...
```

`#` işaretinden sonrası not olarak okunmaz.

**Okuma anahtarı:** Program dosyayı yalnızca `app_licenses` deposunu *okuma* yetkisi olan bir fine-grained token ile okur. Anahtar kaynakta karıştırılmış olarak durur (`kaynak/vba/modLisans.bas` → `JETON_KODU`); yeni anahtar için `python araclar/jeton_gizle.py github_pat_...` çıktısı buraya yazılıp program yeniden kurulur.

> - Anahtarın **süresi dolarsa** ya da silinirse programlar dosyayı okuyamaz ve son başarılı kontrolden 7 gün sonra kilitlenir.
> - `app_licenses` deposunu silmeyin, adını ya da dosyanın yerini değiştirmeyin.
> - Bu kontrol caydırıcıdır, kırılamaz değildir: `.accdb` dosyasında bilgili biri kodu değiştirebilir. Daha sıkı koruma için program `.accde` olarak dağıtılabilir.

### Otomatik güncelleme

- Program her açılışta gizli `app_licenses` deposundaki `apaydin_otomasyon/surum.txt` bildirimini okur. Daha yeni bir sürüm varsa `ZORUNLU = evet` ise sormadan, değilse sorarak kurar.
- Dosyalar bu deponun **Releases** bölümünden iner ve SHA-256 ile doğrulanır. 64 bit Access korumalı `.accde`, 32 bit Access `.accdb` sürümünü alır.
- Kurulumu program klasörüne yazılan `Guncelle.vbs` yapar: program kapanınca eski sürümü `Yedek`'e alır, yenisini koyar, masaüstü kısayollarını yeni dosyaya çevirir, programı yeniden açar.
- **Kademeli dağıtım:** `DENEME_*` satırlarını yalnızca geliştirici anahtarı olan kurulumlar ve `DENEME_KODLAR`'a yazılan kurulumlar alır; önek olmayan satırlar herkese gider.
- Kaynaktan kurulmuş geliştirme kopyasında (`.accdb` + `kaynak` klasörü + geliştirici anahtarı) otomatik güncelleme yapılmaz.
- Ayarlar → **Güncelleme**: sürüm, son denetim, "Güncellemeleri Denetle" ve antivirüs dışlama / güvenilen konum anlatımını içeren **Güncelleme ve Güvenlik Bilgisi** ekranı.

**Yeni sürüm yayınlama:**
1. `kaynak\vba\modOrtak.bas` içinde `URUN_SURUM`'u artırın, `araclar\Kur.vbs` ile programı kurun.
2. `araclar\SurumHazirla.vbs`'yi çalıştırın: `dagitim\<sürüm>\` altında `.accde`, `.accdb`, kurulum `.zip`'i ve `surum_satirlari.txt` oluşur.
3. GitHub'da `v<sürüm>` etiketiyle bir Release açıp üç dosyayı yükleyin.
4. `surum_satirlari.txt`'deki `DENEME_*` satırlarını `app_licenses/apaydin_otomasyon/surum.txt`'ye yazın; deneme bitince aynı değerleri önek olmadan yazarak herkese açın.

### Yedekten başka bilgisayara kurulum

Bilgisayar bozulursa ya da program başka bir bilgisayara taşınacaksa:

1. Programı indirip bir klasöre çıkarın.
2. **Ya** flaştaki en güncel yedeği `Veri` klasörüne `ApaydinVeri.accdb` adıyla kopyalayıp programı açın,
   **ya da** programı örnek veriyle açıp **Ayarlar → Yedekten Geri Yükle** ile flaştaki yedeği seçin.
3. Tüm müşteriler, partiler, hareketler, kullanıcılar ve ayarlar geri gelir.

Elle kopyalarken programın kapalı olduğundan emin olun (klasörde `ApaydinVeri.laccdb` dosyası olmamalı).

### Program güncellemesi ve veri

- Program ve veri ayrı dosyalar olduğu için yeni sürümde yalnızca `ApaydinOtomasyon.accdb` değiştirilir; `Veri` klasörüne dokunulmaz.
- Program açılırken veri dosyasının sürümüne bakar (`tbl_Ayarlar.VeriSurum`). Eskiyse önce `Yedek` klasörüne yedek alır, sonra yeni alanları ekleyip eski kayıtları dönüştürür.
- Veri dosyasına ulaşılamazsa program dosyayı seçmenizi ister ve tabloları ona yeniden bağlar.

## Kullanım

- **Ana menü**: Mal Kabul, Üretim, Yağ Teslim, Müşteriler, Cari Hareketler, Kasa / Banka, Stok Hareketleri, Ürün Kartları, Alım / Satım, Raporlar, Personel; yöneticiye ek olarak Kullanıcılar ve Ayarlar. Aktif sezon menünün sağ üstünde görünür.
- Formlardaki düğmeler aynıdır: **Yeni Kayıt, Kaydet, Güncelle, Sil, Temizle, Kapat**. Kayıtlı bir kaydı değiştirmek için listede çift tıklayıp **Güncelle**'ye basın.
- Raporlarda sezon veya tarih aralığı seçin (Bugün / Bu Hafta / Bu Ay / Bu Yıl kısayolları), gerekirse müşteri / ürün / durum süzün, sonra **Önizle**, **PDF Kaydet** veya **Yazdır (A5)**. PDF'ler program klasöründeki `Raporlar` klasörüne kaydedilir.

## Geliştirme

### Geliştirici modu

Normal açılışta Shift ile tasarım görünümüne geçiş kapalıdır. Yönetici olarak **Ayarlar → Geliştirici Modu (Shift)**'e basın, programı kapatın ve **Shift basılıyken** açın. Bir sonraki normal açılışta izin kendiliğinden yeniden kapanır.

### Kaynaktan kurulum (`araclar\Kur.vbs`)

Asıl kaynak `kaynak/` klasöründeki metin dosyalarıdır. Programı kaynaktan yeniden oluşturmak için Access kapalıyken `araclar\Kur.vbs` dosyasına çift tıklayın. Betik sırasıyla:

1. Mevcut programı `Yedek` klasörüne yedekler, `kaldirilan.txt`'deki eski kaynakları `Yedek\kaynak_eski`'ye taşır.
2. Veri dosyası yoksa oluşturur: eski tek parça programda tablolar varsa onları aktarır, yoksa örnek veriyi kopyalar.
3. `sablon.accdb`'den boş bir program açar; modülleri, tablo bağlantılarını, sorguları, formları ve raporları yükler.
4. Kodu derler ve denetler, açılış ayarlarını yapar (giriş formu, simge, başlık, Shift kapalı), veri yükseltmesini çalıştırır.
5. Sıkıştırıp `ApaydinOtomasyon.accdb` olarak kaydeder. Ayrıntılar `araclar\son_kurulum.log` dosyasına yazılır.

Formlarda veya raporlarda Access içinden değişiklik yaptıysanız **Ayarlar → Kaynakları Dışa Aktar** (ya da `araclar\KaynakDisaAktar.vbs`) ile `kaynak/` klasörünü güncelleyin. VBA modülleri dışa aktarılmaz; onların asıl kaynağı `kaynak/vba`'dır.

### Türkçe karakterler hakkında

Windows'un kod sayfası Türkçe değilse (ör. 1252) VBA dizgilerindeki `ı ş ğ İ` bozulur. Bunu önlemek için:

- `Kur.vbs`, Türkçe kod sayfasıyla hazırlanmış `sablon.accdb`'yi kullanır.
- Yükleme sırasında Türkçe karakter içeren dizgileri `ChrW(...)` ifadelerine çevirir. Kaynak dosyalarda kod okunabilir Türkçe kalır.
- Mesaj kutuları Unicode `MessageBoxW` ile gösterilir.

Bu yüzden modülleri VBA düzenleyicisinden değil, `kaynak/vba` altındaki `.bas` dosyalarından düzenleyip `Kur.vbs`'i çalıştırın.

### Kayıt mantığı

- Hareket tablolarındaki `KaynakTipi / KaynakID / OtomatikMi` alanları, bir kaydın hangi ekrandan otomatik üretildiğini tutar. Otomatik kayıtlar elle değiştirilemez; kaynağı güncellenince yeniden yazılır.
- Kaynak tipleri: `PARTI` (sıkım ücreti borcu veya hakyağı stoğu), `TESLIM` (teslimde tahsilat), `ALIMSATIM`, `KASA`, `MANUEL`.

## Yol haritası

- [ ] Bulut yedekleme / çok kullanıcılı bulut veritabanı
- [ ] Belirli aralıklarla otomatik yedek hatırlatması
- [ ] Banka ekstresi içe aktarma
- [ ] Toplu SMS / e-posta ile teslim bildirimi

## Geliştirici

**Ali Apaydın**

- E-posta: [apaydin.a@gmail.com](mailto:apaydin.a@gmail.com)
- GitHub: [github.com/aliapaydin](https://github.com/aliapaydin)

© 2026 Ali Apaydın. Tüm hakları saklıdır.
