# Hurda Analiz Sistemi

xxx firmasına özelleştirilmiş bir analiz aracı.

**Özel mülkiyet — tüm hakları saklıdır.** Hak sahibinin yazılı izni olmadan kullanım, değiştirme ve dağıtım yasaktır. Ayrıntılar `LICENSE` dosyasındadır.

Beş Excel girdisi üzerinden Y / Ç / HRD fişlerini ve 202 / 204 / 206 stoklarını yıl ve ay bazında analiz eder. PPM, malzeme maliyeti, hata kodu pivotu, Pareto ve PDF/PNG grafik aktarımı içerir. Hesap kuralları `KURULUM.md` içindedir.

## Web yayını

Python/Flask sunucusu gerektirir; GitHub Pages üzerinde çalışmaz. `render.yaml` Render dağıtım tanımıdır. Render'da **New → Blueprint** ile bu özel depoyu bağlayın. Kaynak kod erişimini yalnızca bu depoyla sınırlayın. Kalıcı disk veya veritabanı eklemeyin.

1. Yerel ortamda `python tools/password_hash.py` çalıştırın; site şifresini gizli giriş alanında iki kez girin. Windows'ta ayrı bir pencere açılır; **Şifre özetini panoya kopyala** düğmesi özeti panoya aktarır. Yerel `.venv` kurulumu varsa `tools/SifreHazirla.cmd` dosyasına çift tıklamak da aynı pencereyi açar ve PowerShell script izni gerektirmez.
2. Üretilen hash'i Render ortamında `APP_PASSWORD_HASH` değerine girin. Düz şifreyi veya hash'i GitHub'a koymayın. Şifreyi yalnızca erişmesini istediğiniz kişilerle paylaşın.
3. Blueprint `SECRET_KEY` değerini otomatik üretir. Başlangıç komutu tek Gunicorn işlemi kullanır; oturumlar RAM'de olduğundan işlem sayısını artırmayın.
4. Yayın tamamlandığında HTTPS adresini açın. `/health` uç noktası yalnızca servis sağlığını döndürür.

Şifre olmadan analiz ekranı ve API kullanılamaz. Yapılandırma eksikse web uygulaması başlatılmaz. Şifreyi değiştirmek için yeni hash oluşturun, ortam değişkenini değiştirin ve servisi yeniden başlatın. Mevcut oturumlar da sonlanır.

## Veri saklama

- XLS/XLSX içerikleri multipart aşamasından itibaren yalnızca RAM'de işlenir; uygulama geçici Excel dosyası yazmaz.
- Girdi ve sonuçlar kullanıcı oturumlarına ayrılır; veritabanı, tarayıcı localStorage/IndexedDB veya analiz geçmişi bulunmaz.
- Web oturumu 30 dakika kullanılmazsa en geç sonraki 30 saniyelik temizlikte veriler RAM'den çıkarılır. Çıkış/Oturumu temizle düğmeleri anında temizler. Sunucu yeniden başladığında veriler kaybolur.
- Sekmeyi kapatmak sunucunun hemen haberdar olmasını garanti etmez; hemen temizlemek için çıkış düğmesini kullanın.
- Web Excel/PDF/PNG çıktıları RAM'de oluşturulup tarayıcıya gönderilir; sunucuda çıktı klasörü oluşmaz. İndirilen dosyalar kendi cihazınızda kalır.
- Çerez yalnızca imzalı erişim bilgisi, rastgele oturum kimliği ve CSRF güvenlik belirteci içerir. Excel veya analiz verisi içermez. Yanıtlar `Cache-Control: no-store` kullanır. Formlar ve değişiklik yapan API istekleri oturuma bağlı CSRF belirteciyle doğrulanır.
- Uygulama hata günlüklerine dosya içeriği, maliyet, stok kodu veya şifre yazmaz. Sağlayıcının ağ/erişim günlükleri ve altyapı politikaları kendi kontrolündedir; uygulama sağlayıcının hiç teknik kayıt tutmadığını iddia etmez.
- Web sürümünde seçilen dosyalar analiz için sunucuya aktarılır. Sunucuya hiç göndermemek istenirse ayrıca tarayıcı içinde çalışan bir sürüm gerekir.

## Excel aktarımı

Web sürümü Microsoft Excel kurulumuna ihtiyaç duymaz. Her seçili grup için analiz ve Pareto XLSX dosyalarını ZIP olarak indirir. Dosyalarda formüller ve hesaplanmış önbellek değerleri, yerel Excel grafikleri, yenilenebilir gerçek PivotTable ve analiz dışı stoklar bulunur. Web raporlarının düzeni koddan üretilir; eski şirket şablonlarının logo ve baskı düzeninin birebir kopyası değildir.

Yerel Windows modu, özel `template_files` dosyaları ve kurulu Microsoft Excel ile eski şablonları koruyan aktarımı kullanır. Bu şablonlar geçmiş üretim verisi içerdiğinden depoya dahil edilmez. Yerel çıktılar masaüstüne yazılır. Web modu (`APP_MODE=web`) bunları kullanmaz.

## Yerel geliştirme ve test

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pytest -q
.\.venv\Scripts\python.exe app.py --port 8770
```

`APP_MODE=local` yalnızca localhost bağlantılarını kabul eder. Gerçek üretim dosyaları, şablonlar, çıktılar ve ortam değişkenleri Git'e dahil edilmez. Testler sentetik veri kullanır; özel şablon / Windows Excel isteyen kontroller uygun ortam yoksa atlanır.
