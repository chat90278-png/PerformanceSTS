# sozlesme-app

## Windows masaüstü uygulaması oluşturma

`calistir.bat` dosyasına çift tıklayın. İlk çalıştırmada tek dosyalık `dist\STS.exe` uygulamasını üretir ve uygulamayı başlatır.

Oluşan `dist\STS.exe` bağımsız taşınabilir uygulamadır; yalnızca bu EXE dosyasını başka bir Windows bilgisayara kopyaladığınızda Python kurulumu veya ek klasör gerekmeden çalışır. Masaüstüne kısayol oluşturulmaz.

> Not: Uygulama ikonu için `src\ui\assets\sts_icon.ico` dosyası repoda hazır bulunmalıdır; build sırasında ayrıca ikon üretimi yapılmaz. `calistir.bat`, PyInstaller ikon formatı dönüşümü isteyebileceği için Pillow paketini de kontrol eder/yükler.

## Kapsamlı PDF kullanım kılavuzu

Uygulamadaki **📘 Kullanım Kılavuzu** penceresinin sağ üst köşesindeki **📄 Kapsamlı PDF Kılavuzu** butonu, kapsamlı dokümanı varsayılan PDF görüntüleyiciyle açar.

Beklenen dosya adı `STS_KullanmaKılavuzu.pdf` olmalıdır. Uygulama PDF'yi yalnızca uygulamanın bulunduğu klasörde arar: paketli kullanımda `.exe` dosyasının bulunduğu klasör, kaynak koddan çalıştırmada ise `app.py` dosyasının bulunduğu klasör esas alınır. PDF dosyası farklı bir klasördeyse buton tarafından bulunamaz; dosyayı uygulama klasörüne kopyalayın.
