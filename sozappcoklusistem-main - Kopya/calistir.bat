@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "APP_NAME=STS"
set "DIST_EXE=%~dp0dist\STS.exe"

echo ============================================================
echo  STS tek dosya masaustu uygulamasi hazirlama / calistirma
echo ============================================================
echo.

if exist "%DIST_EXE%" (
    echo Hazir tek dosya uygulama bulundu: "%DIST_EXE%"
    echo STS baslatiliyor...
    start "STS" "%DIST_EXE%"
    exit /b 0
)

set "PY_CMD=py -3"
%PY_CMD% --version >nul 2>&1
if errorlevel 1 set "PY_CMD=python"
%PY_CMD% --version >nul 2>&1
if errorlevel 1 (
    echo HATA: Bu bilgisayarda paketleme icin Python bulunamadi.
    echo Not: Bu gereksinim sadece bu BAT ile EXE uretmek icindir.
    echo Olusan dist\STS.exe baska bilgisayarlarda Python veya ek klasor olmadan calisir.
    pause
    exit /b 1
)

if not exist "%~dp0src\ui\assets\sts_icon.ico" (
    echo HATA: Uygulama ikonu bulunamadi: "%~dp0src\ui\assets\sts_icon.ico"
    echo Lutfen sts_icon.ico dosyasini src\ui\assets klasorune ekleyin.
    pause
    exit /b 1
)

echo PyInstaller kontrol ediliyor...
%PY_CMD% -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo PyInstaller yukleniyor...
    %PY_CMD% -m pip install --upgrade pyinstaller
    if errorlevel 1 (
        echo HATA: PyInstaller yuklenemedi.
        pause
        exit /b 1
    )
)

echo Pillow ikon donusturme destegi kontrol ediliyor...
%PY_CMD% -c "import PIL" >nul 2>&1
if errorlevel 1 (
    echo Pillow yukleniyor...
    %PY_CMD% -m pip install --upgrade pillow
    if errorlevel 1 (
        echo HATA: Pillow yuklenemedi.
        echo Not: PyInstaller, ikon dosyasi tam ICO formatinda degilse Pillow ile donusturme yapar.
        pause
        exit /b 1
    )
)

echo STS tek dosya portable uygulamasi olusturuluyor...
%PY_CMD% -m PyInstaller --noconfirm --clean STS.spec
if errorlevel 1 (
    echo HATA: Uygulama paketlenemedi.
    pause
    exit /b 1
)

if not exist "%DIST_EXE%" (
    echo HATA: Beklenen dosya bulunamadi: "%DIST_EXE%"
    pause
    exit /b 1
)

if exist "%~dp0STS_KullanmaKılavuzu.pdf" (
    echo Kapsamli PDF kilavuzu dist klasorune kopyalaniyor...
    copy /Y "%~dp0STS_KullanmaKılavuzu.pdf" "%~dp0dist\STS_KullanmaKılavuzu.pdf" >nul
)

echo.
echo Tamamlandi.
echo Tek dosya uygulama: "%DIST_EXE%"
echo Uygulama adi: STS
echo Kapsamli PDF kilavuzu kullanmak icin STS_KullanmaKılavuzu.pdf dosyasini STS.exe ile ayni klasorde tutun.
echo Bu STS.exe dosyasini baska Windows bilgisayarlara kopyalayip Python veya ek klasor olmadan calistirabilirsiniz.
echo Masaustune kisayol olusturulmadi.
echo.
echo STS baslatiliyor...
start "STS" "%DIST_EXE%"
exit /b 0
