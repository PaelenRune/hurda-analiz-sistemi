# Hurda Analiz

`Başlat.ps1` dosyasını PowerShell ile çalıştırın. Uygulama yalnızca bu bilgisayarda `http://127.0.0.1:8770/` adresinde açılır.

Beş dosyayı kendi alanlarına yükleyin: Hurdalar, Tüketim, Gramaj, Maliyet, Parçalanmış Hurda. Parçalanmış hurda bulunmayan dönemlerde ilgili kutuyu işaretleyin. Son girdi yüklenince mevcut filtrelerle analiz başlar. Filtreleri değiştirince **Analizi başlat** düğmesini kullanın.

Girdi dosyaları değiştirilmez. Yüklenen dosyalar diske yazılmadan RAM'den okunur; ayrıştırılan veriler yalnızca çalışan oturumun belleğinde tutulur. Yerel oturum iki saat, web oturumu 30 dakika kullanılmazsa en geç sonraki 30 saniyelik temizlikte temizlenir. Oturumu temizle düğmesi de verileri temizler. Uygulama yeniden başlatıldığında geçmiş analizler yüklenmez. Web yayını ve şifre ayarları README.md içindedir.

## Kurallar

- Fiş grupları tam önek ile Y, Ç, HRD olarak ayrılır. HRD, H kayıtlarını kapsamaz.
- Stok grupları 202 kablo, 204 terminal, 206 IDC. Kod uzantıları korunur.
- Parçalanmış dosyada hazır bileşenler ana miktarlara eklenir. 500'lü ürünler hesaplanmaz. Dosyada `HURDA` sayfası varsa bileşenler oradan okunur; diğer pivot sayfaları yeniden eklenmez. Bu girdinin yalnızca ek bileşenleri içermesi gerekir.
- Parçalanmış tabloda tarih veya fiş grubu bulunmuyorsa gelişmiş filtrelerdeki açık seçimler kullanılır. Dosyada dönem/fiş bilgisi varsa bu bilgi önceliklidir.
- Birim maliyet Maliyet girdisinden, stok/yıl/ay içindeki son tarihten alınır. TL/EUR malzeme maliyetleri kullanılır; işçilik dahil edilmez. Aynı son tarihte çelişen değerler varsa analiz durur; dosya düzeltilmelidir.
- Eksik maliyet kullanıcıya sorulur. Evet seçilirse ilgili stok/ay bütün hesaplardan çıkarılır. Eksik tüketim veya gramaj için de aynı kontrol uygulanır. Sıfır maliyet geçerli bir değerdir; boş maliyet değildir.
- PPM yalnızca analize giren stokların tüketimiyle hesaplanır. Kablo kg/kg, diğerleri adet/adet. Çok aylı PPM, toplam hurda / toplam tüketim oranıdır.
- Gramaj dosyasındaki MT→GR/KG ve AD→GR/KG çevrimleri desteklenir. Kablo grubunda farklı kaynak birimleri varsa ham miktarlar birbiriyle toplanmaz; Pareto ortak kg birimiyle gösterilir. Stok satırlarında kaynak birim korunur.
- Boş hata kodları her ay için kayıt sayısının moduyla doldurulur. Eşitlikte önceki beş takvim ayındaki, aynı seçili fiş/stok kapsamındaki aday kodların kayıt sayıları karşılaştırılır. Çözülemeyen eşitlik belirsiz kalır. Orijinal kod ve atama bilgisi detayda korunur.
- Hedefler kullanıcı tarafından girilir. Örnek dosyalardaki hedefler veya geçmiş yıl sayıları güncel veriye otomatik taşınmaz.

## Çıktılar

Excel aktarımı kurulu Microsoft Excel gerektirir. Örnek XLS/XLSX dosyaları salt okunur şablonlardır; hücre biçimleri ve sayfa düzenleri korunarak yeni XLSX dosyaları üretilir. Aylık tablolar, formüllü toplamlar, gerçek Excel grafikleri, yerel PivotTable ve hariç tutulan stoklar bulunur. Seçilen grupların analiz ve Pareto dosyaları masaüstündeki **Çıktı/Analiz_tarih_saat** klasörüne yazılır. Önceki çıktılar ezilmez. Seçilmeyen ayların eski verileri temizlenir.

PDF/PNG grafikleri, arayüzde gösterilen grafiklerle aynı verilerden üretilir. PDF çok sayfalıdır; tüm PNG'ler seçilirse ZIP indirilir. Tek grafik doğrudan PNG olarak indirilir.

## Yeniden kurulum

Python 3.12+ ve Microsoft Excel kurulu bir Windows bilgisayarda:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app.py --port 8770
```

Altı örnek Excel şablonu `template_files` klasöründedir. Bu dosyaları aynı adlarla koruyun. Testler: `.\.venv\Scripts\python.exe -m pytest -q`.
