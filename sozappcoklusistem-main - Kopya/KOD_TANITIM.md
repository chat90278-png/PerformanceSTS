# STS Kod Tanıtım Dosyası

Bu doküman, **KONFİGÜRASYON YÖNETİMİ SÖZLEŞME TAKİP SİSTEMİ (STS)** kod tabanına sonradan müdahale edecek geliştiriciler için hazırlanmıştır. Amaç; uygulamanın genel mimarisini, kullanılan kütüphaneleri, özel davranışları ve her `.py` dosyasının sorumluluklarını hızlıca anlaşılır hale getirmektir.

> Kısa özet: Bu proje, Excel dosyasını veri deposu olarak kullanan PySide6 tabanlı bir Windows masaüstü sözleşme takip uygulamasıdır. Platform, sözleşme, sistem, teslimat/kabul, kullanıcı, bileşen ve etiket yönetimi yapar; Excel okuma/yazma işlemlerini `openpyxl` ile, arayüzü PySide6 ile, uzun süren işlemleri Qt worker/thread yapısıyla yürütür.

## 1. Projenin Genel Amacı

Uygulama sözleşme takip süreçlerini Excel üzerinde standart bir veri modeline oturtur. Kullanıcı uygulama açılışında bir `.xlsx` veya `.xlsm` dosyası seçer; uygulama bu dosyadaki platform sayfalarını, kullanıcı/bileşen/etiket yardımcı sayfalarını ve sözleşme bloklarını okuyarak arayüzde listeler. Kullanıcı arayüz üzerinden sözleşme, sistem, teslimat, kabul, etiket ve platform ayarlarını değiştirir; değişiklikler tekrar Excel dosyasına yazılır.

Temel işlev kümeleri şunlardır:

- Excel dosyası bağlama, oluşturma, okuma ve kaydetme.
- Platform bazlı sözleşme listesi, gelişmiş filtreleme ve arama.
- Ana sözleşme ve SD/ek sözleşme ilişkileri.
- Sözleşme altında sistem tanımlama, sistem bileşen miktarları ve sistem tip şablonları.
- Teslimat/kabul kayıtları, planlanan/teslim edilen miktar takibi ve otomatik kabul oluşturma.
- Kullanıcı, bileşen, platform, platform logosu ve etiket yönetimi.
- Takvim görünümü, yaklaşan terminler ve özet/popup ekranları.
- Excel üzerinde değişiklik kayıtları, sürüm bilgisi, görsel stiller ve koşullu biçimlendirme.

## 2. Ana Mimari

Proje katmanları genel olarak şu şekilde ayrılmıştır:

```text
app.py                         Ana uygulama, ana pencere ve büyük dialog akışları
src/config/                    Uygulama sabitleri, tema renkleri, dosya/sayfa adları
src/domain/                    Domain düzeyi sabitler ve alternatif veri modelleri
src/models/                    Uygulamanın aktif dataclass veri modelleri
src/services/                  Excel veri deposu, versiyon, cache ve performans servisleri
src/workers/                   QThread içinde çalışan Excel yükleme/kaydetme worker sınıfları
src/ui/                        Tekrar kullanılabilir UI parçaları, dialoglar, takvim, özet ekranı
tools/                         Geliştirici/asset yardımcı scriptleri
auto_accept.py                 Otomatik kabul/teslimat dialog akışı
```

### 2.1. Veri Kaynağı ve Kalıcılık

- Ana kalıcılık katmanı `src/services/excel_store.py` içindeki `ExcelStore` sınıfıdır.
- Excel workbook içinde platform sayfaları sözleşme verilerini tutar.
- Sistem yardımcı sayfaları şunları içerir: `Sistem Bileşenleri`, `Kullanıcılar`, `Etiketler`, `Config`, `_Meta`, değişiklik kayıtları ve platform logo sayfası.
- Excel sayfalarında sözleşmeler blok mantığıyla tutulur: ana sözleşme satırları, sistem satırları, teslimat/kabul satırları ve toplam satırları aynı platform sayfasında yer alır.
- `openpyxl` ile hücre okuma/yazma, stil, merge, kolon genişliği, koşullu biçimlendirme ve gizli sayfa işlemleri yapılır.

### 2.2. UI Mimarisi

- UI katmanı PySide6 üzerine kuruludur.
- `MainWindow`, platform ve sözleşme listesi ana ekranını yönetir.
- `ContractWorkWindow`, tek bir sözleşme ailesi üzerinde çalışma ekranıdır.
- Büyük dialogların bir bölümü hâlâ `app.py` içindedir; daha yeni/refactor edilmiş UI parçaları `src/ui/` altına taşınmıştır.
- Tarih alanlarında özel takvim popup'ı (`src/ui/date_picker.py`) kullanılır.
- Takvim takip ekranı `src/ui/tarih.py`, özet ekranı `src/ui/ozet.py`, kullanım kılavuzu `src/ui/kullanim_kilavuzu.py` dosyalarındadır.

### 2.3. Worker/Thread Mimarisi

Excel dosyaları büyük olabileceği için yükleme ve kaydetme işlemlerinin bir kısmı UI thread'i dışında çalıştırılır:

- `ExcelLoadWorker`: Excel dosyasını açar, platformları ve sözleşme indeksini oluşturur.
- `ComponentSaveWorker`: bileşen listesini asenkron kaydeder.
- `UserSaveWorker`: kullanıcı listesini asenkron kaydeder.
- `ContractSaveWorker`: sözleşme ailesi kaydını asenkron işler.
- Worker sınıfları Qt `Signal` kullanarak ilerleme, tamamlanma ve hata bilgisini UI'a iletir.

### 2.4. Özel/Önemli Özellikler

- **Excel merkezli mimari:** Uygulamanın veritabanı gibi davranan ana dosya Excel workbook'tur.
- **Platform sayfası ayrımı:** Platformlar Excel sayfası olarak temsil edilir. Sistem sayfaları `CORE_SHEETS` ve `EXTRA_SYSTEM_SHEET_NAMES` ile ayıklanır.
- **Türkçe normalize işlemleri:** Sayfa adları, arama ve sıralama için Türkçe karakterleri normalize eden yardımcı fonksiyonlar vardır.
- **Sözleşme ailesi:** Ana sözleşme ve SD/ek sözleşmeler birlikte ele alınır; kayıt ve indeks güncelleme bu aile mantığına göre yapılır.
- **Sistem tipi şablonları:** Sık kullanılan sistem bileşen kombinasyonları Excel'de `SistemTipleri` sayfasında saklanabilir.
- **Etiket sistemi:** Etiket tanımları ve sözleşme atamaları aynı `Etiketler` sayfasında farklı kayıt tipleriyle tutulur.
- **Değişiklik logları:** Sözleşme güncellemeleri yıllık Excel log dosyalarına yazılacak şekilde tasarlanmıştır.
- **Sürüm yönetimi:** Workbook içinde `_Meta` sayfasında `STS_vA1` benzeri sürüm bilgileri tutulur.
- **Yerel cache adımı:** `LocalCacheDB` SQLite tabanlı indeks cache için hazırlanmıştır; listeleme/filtreleme performansını iyileştirmek amacı taşır.
- **Kullanıcı deneyimi:** Toast, busy overlay, stat card, özel tarih seçici, ikonlu takvim ve kullanım kılavuzu gibi UI yardımcıları bulunur.

## 3. Kullanılan Başlıca Kütüphaneler

### 3.1. Harici Kütüphaneler

- **PySide6**
  - Masaüstü arayüzünün ana çatısıdır.
  - `QApplication`, `QMainWindow`, `QDialog`, `QTableWidget`, `QComboBox`, `QDateEdit`, `QThread`, `Signal`, `QPainter`, `QPixmap` gibi sınıflar kullanılır.
  - Dialoglar, tablolar, özel delegeler, popup'lar, overlay'ler, takvim ve toast bildirimleri bu kütüphane ile oluşturulur.

- **openpyxl**
  - Excel workbook okuma/yazma katmanıdır.
  - `Workbook`, `load_workbook`, `Font`, `PatternFill`, `Alignment`, `Border`, `Side`, `get_column_letter` gibi API'ler kullanılır.
  - Sayfa oluşturma, hücre yazma, stil verme, merge/unmerge, conditional formatting, gizli sayfa ve satır/kolon yönetimi bu kütüphane ile yapılır.

### 3.2. Standart Python Kütüphaneleri

- **pathlib**: Dosya yollarını platform bağımsız yönetmek için kullanılır.
- **datetime / calendar / time**: Tarih hesaplama, termin, T0+ay, kabul tarihi, takvim ve performans ölçümü için kullanılır.
- **dataclasses**: Veri taşıyıcı modelleri sade tanımlamak için kullanılır.
- **typing**: Tip ipuçları için kullanılır.
- **re**: Sayfa adı temizleme, sürüm formatı parse etme, arama/normalize işlemleri için kullanılır.
- **sqlite3**: Yerel indeks cache katmanı için kullanılır.
- **json**: Performans kayıtlarını JSON Lines formatında saklamak için kullanılır.
- **threading / contextlib**: ExcelStore içinde kayıt kilidi ve context manager destekleri için kullanılır.
- **getpass / socket**: İşlemi yapan kullanıcı/aktör bilgisini bulmak için kullanılır.
- **base64**: Platform logo görsellerini Excel hücrelerinde saklanabilir metne dönüştürmek için kullanılır.
- **ctypes / sys / os**: Windows AppUserModelID, paketli uygulama yolu ve çalışma ortamı kontrolleri için kullanılır.
- **struct**: `.ico` dosyası üretim scriptinde binary icon formatı yazmak için kullanılır.

## 4. Çalışma Akışı

1. `app.py` çalışır, Windows ikon kimliği ayarlanır ve PySide6 uygulaması başlatılır.
2. Kullanıcı Excel dosyasını `WorkbookStartDialog` ile seçer veya sürükle-bırak yapar.
3. `MainWindow.start_excel_load()` bir `ExcelLoadWorker` başlatır.
4. Worker `ExcelStore.open_or_create()` ve read-only indeksleme adımlarıyla workbook'u analiz eder.
5. Ana ekranda platform listesi, sözleşme tablosu, filtreler, yaklaşan terminler ve bağlantı/sürüm göstergeleri güncellenir.
6. Kullanıcı sözleşme açarsa `ContractWorkWindow` devreye girer.
7. Sözleşme ekranında ana bilgiler, sistemler, teslimatlar, kabuller, etiketler ve özet tablolar düzenlenir.
8. Kaydetmede `ExcelStore.write_contract()` veya ilgili servis metotları Excel'e yazar; worker kullanılıyorsa ilerleme UI'a sinyal olarak döner.
9. Gerekirse indeks, takvim ve açık ekranlar yenilenir.

## 5. Dosya Bazlı Kod Tanıtımı

Aşağıdaki bölümde repodaki her `.py` dosyası tek tek açıklanmıştır.

---

### `app.py`

Projenin ana giriş ve orkestrasyon dosyasıdır. Uygulamanın büyük bölümü bu dosyada bulunur: ana pencere, sözleşme çalışma penceresi ve birçok yönetim dialogu burada tanımlıdır.

İçerdiği başlıca fonksiyonlar:

- `app_icon_path()`: Windows `.ico` varsa onu, yoksa SVG logoyu döndürür.
- `configure_windows_app_identity()`: Windows taskbar ikonunun doğru görünmesi için AppUserModelID ayarlar.
- `normalize_sheet_name()`, `is_system_sheet_name()`, `safe_sheet_name()`: Excel sayfa adlarını normalize eder, sistem sayfalarını platformlardan ayırır ve güvenli sayfa adı üretir.
- `to_iso()`, `parse_iso_date()`, `add_months()`, `iso_or_blank()`: Tarih dönüşüm ve T0+ay hesaplama yardımcılarıdır.
- `contract_date_picker_events()`: Sözleşme/sistem terminlerini tarih seçici popup için event listesine çevirir.
- `as_number()`, `fmt_num()`: Excel veya UI değerlerini sayıya çevirir ve okunabilir formatlar.
- `tag_chip_style()`: Etiket butonları için renkli chip stili üretir.
- `section_label()`, `configure_table()`, `fill_table()`: UI tablo ve bölüm yardımcılarıdır.

Başlıca sınıflar ve sorumlulukları:

- `ElidedLabel`: Uzun metni alan daraldığında üç nokta ile gösteren QLabel türevi.
- `FilterableHeaderView`: Sözleşme tablosu başlıklarına filtre popup'ı ekler; değer, tarih aralığı ve gün aralığı filtrelerini yönetir.
- `StyledDialog`: Ortak dialog tabanı; footer durum mesajı gösterme altyapısı sağlar.
- `UserManagerDialog`: Excel'deki kullanıcı listesini gösterir, kullanıcı ekleme/silme/kaydetme işlemlerini yapar.
- `ComponentManagerDialog`: Sistem bileşenlerini ve platform bazlı aktifliklerini yönetir.
- `ContractDialog`: Yeni ana sözleşme veya SD sözleşme oluşturma formudur; zorunlu alan, tarih, kullanıcı ve duplicate kontrolü yapar.
- `ContractEditDialog`: Mevcut sözleşmenin ana bilgilerini düzenler.
- `TagAssignDialog`: Bir sözleşmeye etiket seçme/atama dialogudur.
- `TagManagerDialog`: Etiket tanımı oluşturma, düzenleme, silme ve kullanım/atama görüntüleme ekranıdır.
- `SystemDialog`: Tek sistem ekleme/düzenleme ekranıdır; sistem adı, T0, termin, durum ve bileşen seçimini yönetir.
- `MultiSystemDialog`: Aynı sözleşmeye birden fazla sistem taslağı ekleme ekranıdır; çoğaltma, silme, sistem tipi uygulama ve toplu kabul destekler.
- `DeliveryDialog`: Teslimat/kabul kaydı ekleme/düzenleme ekranıdır; planlanan ve teslim edilen miktarları karşılaştırır.
- `ContractWorkWindow`: Bir sözleşme ailesinin ana çalışma ekranıdır. Sistem listesi, teslimatlar, etiket overlay'i, SD gezinme, durum türetme, Excel'e kaydetme ve özet tablolar burada yönetilir.
- `MainWindow`: Uygulamanın ana penceresidir. Excel yükleme, platform listesi, sözleşme indeks tablosu, filtreler, takvim, kullanıcı/bileşen/platform/etiket yönetim ekranları ve sözleşme açma akışlarını yönetir.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Ana ekrana yeni filtre veya kolon eklemek.
- Sözleşme kartı/çalışma penceresi davranışını değiştirmek.
- Yeni sözleşme, sistem, teslimat veya etiket dialog davranışını düzenlemek.
- SD sözleşme aile mantığı veya otomatik durum türetme kurallarını değiştirmek.
- Uygulama açılış/yükleme akışında değişiklik yapmak.

---

### `auto_accept.py`

Otomatik kabul/teslimat oluşturma ekranını içerir. `ContractWorkWindow` içinden çağrılır ve sistem bileşenlerini kabul kayıtlarına dağıtmayı kolaylaştırır.

İçerdiği yardımcı fonksiyonlar:

- `as_number()`, `fmt_num()`: Miktar değerlerini normalize eder ve gösterime uygun formatlar.
- `parse_iso_date()`, `to_iso()`, `iso_or_blank()`, `add_months()`: Kabul tarihi ve T0+ay hesaplamalarında kullanılır.
- `open_auto_accept_dialog()`: Uygun sistem için dialogu açan dış erişim fonksiyonudur.

`AutoAcceptDialog` sınıfının işlevleri:

- Kabul sayısına göre kartlar oluşturur.
- Her kabul kartında durum, tarih, not ve bileşen miktar tablosu gösterir.
- “Planlananı teslim edilene doldur”, “tüm sistemi doldur”, “kalanı doldur” gibi hızlı doldurma aksiyonları sağlar.
- Fazla atanan bileşenleri bulur ve satırları görsel olarak uyarır.
- Henüz atanmamış/kalan bileşen miktarlarını panelde gösterir.
- Kaydetmede oluşturulan kabul kayıtlarını çalışma penceresindeki sistem teslimat listesine ekler.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Otomatik kabul oluşturma kuralı değişecekse.
- Teslimat/kabul miktar validasyonu değişecekse.
- Kabul kartı UI'ı veya hızlı doldurma butonları değişecekse.

---

### `src/__init__.py`

`src` klasörünü Python paketi haline getiren boş/işaretleyici dosyadır. İş mantığı içermez.

---

### `src/config/__init__.py`

`src.config` paket işaretleyicisidir. İş mantığı içermez.

---

### `src/config/app_config.py`

Uygulama düzeyi konfigürasyon sabitlerini içerir.

Önemli içerikler:

- `APP_TITLE`: Ana pencere başlığı.
- `DEFAULT_FILE`: Varsayılan Excel dosya adı.
- `APP_ICON_PATH`, `APP_ICON_ICO_PATH`, `APP_ID`: Logo/ikon ve Windows uygulama kimliği.
- `COMP_SHEET`, `USERS_SHEET`, `PLATFORM_LOGO_SHEET`, `TAG_SHEET`: Excel yardımcı sayfa adları.
- `TAG_KIND_DEF`, `TAG_KIND_ASSIGN`: Etiket sayfasında tanım ve atama satırlarını ayıran kayıt tipleri.
- `LOG_FOLDER_NAME`: Değişiklik kayıtlarının yazılacağı klasör adı.
- Renk sabitleri: `NAVY`, `LIGHT`, `CARD`, `HEAD`, `BLUE`, `GREEN`, `GRID`, `TEXT_MUTED`.
- `BASE_HEADERS`, `MAIN_TOTAL_LABEL`, `SYSTEM_TOTAL_SUFFIX`: Excel platform sayfalarındaki ana kolon/toplam metinleri.
- `TR_MONTHS`, `TR_WEEKDAYS`: Türkçe ay/gün adları.
- `LOG_HEADERS`, `TAG_HEADERS`: Log ve etiket sayfası kolonları.
- `EXTRA_SYSTEM_SHEET_NAMES`: Platform sayılmaması gereken ek teknik sayfa adları.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Excel yardımcı sayfa adı değişirse.
- Uygulama başlığı, ikon yolu veya varsayılan dosya adı değişirse.
- Türkçe ay/gün metinleri veya standart renk paleti güncellenecekse.

---

### `src/domain/__init__.py`

`src.domain` paket işaretleyicisidir. İş mantığı içermez.

---

### `src/domain/constants.py`

Domain seviyesinde UI ve Excel katmanlarının ortak kullandığı sabitleri içerir.

Önemli içerikler:

- `CORE_SHEETS`: Platform olmayan çekirdek Excel sayfaları.
- `HEADER_ROW`, `SUBHEADER_ROW`, `DATA_START_ROW`: Platform sayfalarındaki standart başlık ve veri başlangıç satırları.
- `MAIN_COLUMN_HEADERS`: Sözleşme platform sayfalarının ana kolon başlıkları.
- `STATUS_VALUES`: Standart durum değerleri: `Başlanmadı`, `Teslimata Hazırlanıyor`, `Parçalı Teslimat`, `Teslim Edildi`.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Excel platform sayfası kolon düzeni değişirse.
- Standart durum seti değişirse.
- Başlık/veri başlangıç satırı farklılaşırsa.

---

### `src/domain/models.py`

Domain tarafı için alternatif/özet dataclass modelleri içerir. Aktif UI akışlarında daha çok `src/models/app_models.py` kullanılsa da bu dosya domain temsilini sade tutar.

Modeller:

- `Platform`: Platform adı, sayfa adı ve aktiflik bilgisi.
- `User`: Kullanıcı adı, birim, e-posta, telefon, aktiflik ve not.
- `ComponentDefinition`: Bileşen adı, versiyon, birim, aktiflik, kullanım ve platform eşleşmeleri.
- `ContractSummary`: Platform, sözleşme no, kullanıcı, tip, durum, içerik, satır ve arama metni.
- `SystemRecord`: Sistem adı ve bileşen miktarları.
- `DeliveryRecord`: Teslimat/kabul adı, durum, kabul tarihi, not, planlanan ve teslim edilen miktarlar.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Domain katmanı bağımsızlaştırılırsa.
- Servis/UI dışı test veya API benzeri bir katman için sade modeller genişletilecekse.

---

### `src/models/__init__.py`

`src.models` paket işaretleyicisidir. İş mantığı içermez.

---

### `src/models/app_models.py`

Uygulamanın ana veri taşıyıcı dataclass modellerini içerir. ExcelStore, UI dialogları ve worker katmanı ağırlıklı olarak bu modelleri kullanır.

Modeller:

- `ComponentDef`: Bileşen adı, versiyon, birim, aktiflik, kullanım sayısı ve platform bazlı aktiflik sözlüğü.
- `ContractInfo`: Sözleşme no, platform, kullanıcı, Yİ/YD, sözleşme tipi, imza/T0/termin tarihleri, durum, not, kabul tarihi ve SD aile anchor bilgileri.
- `SystemInfo`: Sistem adı, bileşen miktarları, T0, T0+ay, termin, durum ve kabul tarihi.
- `DeliveryInfo`: Teslimat/kabul adı, durum, kabul tarihi, not, planlanan ve teslim edilen miktarlar, T0/T0+ay/termin alanları.
- `TagDef`: Etiket adı, renk, not ve aktiflik bilgisi.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Sözleşme, sistem, teslimat veya etiket veri modeline yeni alan eklenirse.
- Excel okuma/yazma ile UI formları arasında taşınacak ortak veri yapısı değişirse.

---

### `src/services/__init__.py`

`src.services` paket işaretleyicisidir. İş mantığı içermez.

---

### `src/services/excel_store.py`

Projenin en kritik servis dosyasıdır. Excel dosyası ile tüm okuma/yazma, stil, indeks, sistem tipi ve log işlemlerini yönetir.

Dosya başındaki yardımcı fonksiyonlar:

- `normalize_sheet_name()`, `is_system_sheet_name()`, `safe_sheet_name()`: Sayfa adı normalize/filtre/güvenli ad üretimi.
- `to_iso()`, `parse_iso_date()`, `add_months()`, `iso_or_blank()`: Tarih yardımcıları.
- `as_number()`, `fmt_num()`: Miktar ve toplam formatlama.

`ExcelStore` sınıfının ana sorumlulukları:

- Workbook açma/oluşturma: `open_or_create()`, `reload_from_disk()`, `ensure_core()`.
- Kaydetme ve kilitleme: `save()`, `batch_save()`, `_invalidate_runtime_caches()`.
- Log altyapısı: `_log_folder()`, `_ensure_log_sheet()`, `_append_log_rows()`, `_contract_snapshot()`, `_contract_diff_logs()`.
- Sözleşme silme: `delete_contract()`.
- Platform/sayfa yönetimi: `platform_names()`, `all_sheet_names()`, `load_excluded_platforms()`, `save_excluded_platforms()`, `create_platform()`, `setup_platform_sheet()`, `rebuild_platform_headers()`.
- Bileşen yönetimi: `ensure_component_sheet()`, `load_components()`, `write_components()`, `assigned_components()`, `increment_component_usage()`.
- Kullanıcı yönetimi: `ensure_user_sheet()`, `load_users()`, `write_users()`.
- Platform logo yönetimi: `ensure_platform_logo_sheet()`, `set_platform_logo()`, `get_platform_logo_bytes()`.
- Etiket yönetimi: `ensure_tag_sheet()`, `load_tag_defs()`, `write_tag_defs()`, `upsert_tag_def()`, `delete_tag_def()`, `load_contract_tags()`, `save_contract_tags()`, `delete_contract_tags()`, `tag_usage_counts()`, `list_tag_assignments()`, `load_tag_snapshot()`, `rename_tag_assignments()`.
- Excel görsellik/stil: `style_component_sheet()`, `style_platform_rows()`, `style_platform_rows_range()`, `flush_pending_styles()`, `_apply_platform_cf_rules()`, `migrate_platform_cf_rules()`.
- Merge ve satır blok yönetimi: `_unmerge_data_cells_preserve_values()`, `_apply_visual_merges_for_block()`, `_repair_single_contract_block_merges()`, `_repair_contract_merges_near_row()`, `_contract_block_rows()`, `next_row()`.
- Sözleşme aile/SD işlemleri: `find_main_contract_info()`, `next_sd_code()`, `_physical_family_last_row()`, `_safe_insert_rows()`, `_sd_insert_row_from_anchor()`, `write_contract()`, `update_linked_sd_contract_numbers()`.
- İndeks ve detay okuma: `list_main_contracts()`, `all_contract_tags_map()`, `build_contract_index()`, `load_contract_structure()`.

Dosya sonundaki sistem tipi fonksiyonları:

- `_ensure_system_type_sheet()`: Sistem tipi şablon sayfasını oluşturur.
- `_system_type_platform_key()`: Platform anahtarını normalize eder.
- `list_system_type_names()`: Kaydedilmiş sistem tiplerini listeler.
- `get_system_type_components()`: Bir sistem tipinin bileşenlerini döndürür.
- `get_system_type_component_quantities()`: Sistem tipi bileşen miktarlarını döndürür.
- `save_system_type()`: Seçili bileşen/miktar kombinasyonunu sistem tipi olarak kaydeder.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Excel dosya formatı, kolonlar veya satır blok mantığı değişirse.
- Sözleşme yazma/okuma algoritması değişirse.
- Etiket, kullanıcı, bileşen, platform logo veya sistem tipi kalıcılığı değişirse.
- Excel stilleri, koşullu biçimlendirme veya merge davranışı değişirse.
- Değişiklik logları veya workbook sürümleme mantığı değişirse.

---

### `src/services/local_cache_db.py`

SQLite tabanlı yerel indeks cache servisini içerir. Amaç, Excel'den okunan ana sözleşme listesini SQLite üzerinde tutup listeleme/filtreleme/sıralama sorgularını daha hızlı yapabilmektir.

`LocalCacheDB` sınıfı:

- `__init__()`: DB bağlantısını açar ve `row_factory` ayarlar.
- `_setup()`: WAL modu, performans PRAGMA'ları ve `contracts` tablosunu oluşturur.
- `replace_contracts()`: Mevcut kontrat cache'ini topluca silip yeni satırlarla değiştirir.
- `query_contracts()`: Platform, arama, durum ve tarih aralığı filtreleriyle sözleşme satırlarını sorgular.
- `close()`: SQLite bağlantısını kapatır.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Ana liste performansı için DB cache aktif kullanılacaksa.
- Sözleşme indeksine yeni filtrelenebilir/sıralanabilir alan eklenecekse.
- Excel yerine daha kalıcı/harici DB katmanına geçiş planlanıyorsa.

---

### `src/services/perf_tracker.py`

Performans ölçümü ve kayıt dosyası yönetimi için yardımcı fonksiyonlar içerir.

Fonksiyonlar:

- `_log_path()`: Kullanıcı home klasörü altında STS performans log dosyası yolunu üretir.
- `record()`: Bir operasyonun süre, dosya boyutu, satır sayısı ve not bilgisini JSON Lines formatında kaydeder.
- `measure()`: Context manager olarak kod bloğu süresini ölçer ve `record()` çağırır.
- `load_records()`: Kayıt dosyasından performans kayıtlarını okur.
- `compute_stats()`: Operasyon bazında adet, ortalama, min, max ve son değer istatistikleri üretir.
- `file_size_mb()`: Dosya boyutunu MB cinsinden döndürür.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Excel yükleme/kaydetme performansını izlemek için yeni metrik eklenecekse.
- Log formatı veya saklama yeri değiştirilecekse.

---

### `src/services/version_manager.py`

Excel workbook sürüm yönetimini sağlar. Sürüm formatı `STS_vA1`, `STS_vA2`, ..., `STS_vA9`, `STS_vB1` şeklindedir.

Fonksiyonlar:

- `versioned_workbook_path()`: Workbook dosya adını sürüm adına göre yeniden üretir.
- `save_store_as_versioned_file()`: Store'u sürümlü dosya adıyla kaydeder; eski dosyayı silmeye çalışır.
- `parse_version()`: `STS_vA1` benzeri metni `(harf_index, build)` tuple'ına çevirir.
- `format_version()`: Harf index ve build değerinden sürüm metni üretir.
- `increment_version()`: Mevcut sürümü bir artırır; geçersizse `STS_vA1` döndürür.
- `read_version()`: Workbook `_Meta` sayfasından mevcut sürümü okur.
- `write_version()`: `_Meta` sayfasına sürüm, tarih ve aktör bilgisini yazar.
- `bump_version()`: Sürümü oku-artır-yaz akışını tek fonksiyonda yapar.
- `_default_actor()`: Bilgisayar adı ve kullanıcı adından aktör üretir.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Sürüm formatı veya artırma mantığı değişirse.
- Workbook sürüm bilgisinin yazıldığı `_Meta` yapısı değişirse.

---

### `src/ui/__init__.py`

`src.ui` paket işaretleyicisidir. İş mantığı içermez.

---

### `src/ui/contract/work_window_deliveries.py`

`ContractWorkWindow` içindeki teslimat/kabul paneli fonksiyonlarını modülerleştirmek için hazırlanmış yardımcı dosyadır. Fonksiyonlar `self` parametresiyle çalışma penceresi instance'ı üzerinde çalışır.

Fonksiyonlar:

- `refresh_delivery_table(self)`: Seçili sistemin teslimat/kabul listesini yeniler.
- `delivery_detail_widget(self, delivery)`: Teslimat detay satırı/panel widget'ını üretir.
- `edit_delivery(self, idx)`: Seçili teslimatı düzenleme akışını açar.
- `add_delivery(self)`: Seçili sisteme yeni teslimat/kabul ekleme akışını başlatır.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Teslimat tablosunun görünümü veya etkileşimi değişirse.
- Teslimat ekleme/düzenleme akışı `app.py` dışına taşınmaya devam ederse.

---

### `src/ui/contract/work_window_view.py`

`ContractWorkWindow` içindeki özet ve sağ panel yenileme fonksiyonlarını modülerleştirmek için hazırlanmıştır.

Fonksiyonlar:

- `update_system_metric_cards(self)`: Seçili sistem için teslimat, bileşen ve durum metrik kartlarını günceller.
- `refresh_summary_only(self)`: Sistem/bileşen özet tablosunu sadece özet düzeyinde yeniler.
- `refresh_right(self)`: Sağ paneli, seçili sistem ve teslimat bilgilerine göre komple yeniler.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Sözleşme çalışma ekranındaki sağ panel veya metrik kartları değişirse.
- `ContractWorkWindow` refactor edilerek parçalara ayrılmaya devam ederse.

---

### `src/ui/date_picker.py`

Özel tarih seçici popup'ını içerir. Standart `QDateEdit` yerine, uygulamanın termin/kabul olaylarını gösterebilen modern bir popup takvim sağlar.

Fonksiyon ve sınıflar:

- `parse_iso_date()`: ISO tarih metnini `date` nesnesine çevirir.
- `DatePickerDay`: Takvimdeki tek gün hücresidir; event işaretleri, tooltip ve seçili/bugün görünümü çizer.
- `DatePickerPopup`: Ay/yıl seçimi, önceki/sonraki ay gezinmesi, event yenileme, max tarih kısıtı ve gün seçme davranışını yönetir.
- `build_date_input()`: Uygulamada kullanılan tarih input bileşenini üretir; line edit + takvim butonu + popup bağlamasını yapar.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Tarih alanlarının görünümü veya validasyonu değişirse.
- Takvim popup'ında olay/termin gösterimi farklılaştırılırsa.
- Max tarih veya geçmiş/gelecek tarih seçme kuralları değişirse.

---

### `src/ui/delegates.py`

Qt tablo hücreleri için özel delegate sınıflarını içerir.

Sınıflar:

- `CompactNumberDelegate`: Sayısal hücrelerde metni ortalar; doğrudan editör oluşturmaz.
- `CenterTableDelegate`: Hücre metnini ortalar; doğrudan editör oluşturmaz.
- `DropdownDelegate`: Hücre içinde `QComboBox` editörü açar; seçenekleri dışarıdan verilir.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Tablolarda hücre hizalama, editör veya combobox davranışı değişirse.
- Yeni özel tablo editörü gerekiyorsa.

---

### `src/ui/dialogs/contract_summary_popup.py`

Ana sözleşme listesinde hızlı özet göstermek için kullanılan popup dialogudur.

İçerikler:

- `_configure_popup_table()`: Popup içindeki tabloların ortak görünümünü ayarlar.
- `ContractSummaryPopup`: Frameless/popup tarzı özet ekranı. Sözleşme bilgilerini, sistem sayılarını, teslimat sayılarını ve tablolu detayları gösterir. Sürüklenebilir davranış ve ESC/dış tık ile kapanma desteği vardır.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Sözleşme listesinde açılan hızlı özet popup'ı değişecekse.
- Popup'a yeni metrik, tablo veya görsel alan eklenecekse.

---

### `src/ui/dialogs/platforms.py`

Platform oluşturma ve platform sayfalarını aktif/pasif işaretleme dialoglarını içerir.

Fonksiyonlar:

- `safe_sheet_name()`: Platform adından Excel'e uygun sayfa adı üretir.
- `form_label()`: Dialog form etiketi oluşturur.

Sınıflar:

- `PlatformDialog`: Yeni platform adı alır, opsiyonel logo seçtirir, `ExcelStore.create_platform()` ile platform sayfasını oluşturur ve logo kaydeder.
- `PlatformManagerDialog`: Tüm Excel sayfalarını listeler, hangilerinin platform olmadığını `Config` üzerinden işaretler, yeni platform ekleme ve ayar kaydetme işlemlerini yapar.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Platform ekleme formu veya logo yükleme akışı değişirse.
- Excel'de platform dışı sayfa işaretleme mantığı değişirse.

---

### `src/ui/dialogs/workbook_start.py`

Uygulama açılışında Excel dosyası seçtiren veya sürükle-bırak ile alan dialogdur.

`WorkbookStartDialog` sınıfı:

- `.xlsx` ve `.xlsm` uzantılı dosyaları kabul eder.
- Dosya seçme butonu ve drag-drop alanı sunar.
- Seçilen yolu `selected_path` alanında saklar.
- Uygun olmayan dosya bırakılırsa uyarı gösterir.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Açılışta farklı dosya türü kabul edilecekse.
- Son kullanılan dosya, otomatik açılış veya yeni workbook oluşturma seçeneği eklenecekse.

---

### `src/ui/kullanim_kilavuzu.py`

Uygulama içi kullanım kılavuzu dialogunu içerir.

İçerikler:

- `application_directory()`: Paketli `.exe` ve kaynak kod çalıştırma durumlarını dikkate alarak uygulama klasörünü bulur.
- Kılavuz sayfa verileri: Başlık, açıklama, ekran görüntüsü ve ipucu içerikleri tanımlıdır.
- `GuideImage`: Kılavuzdaki ekran görüntülerini uygun ölçekle gösterir; görsel yoksa placeholder üretir.
- `UsageGuideDialog`: Sayfa sayfa kullanım kılavuzu UI'ını kurar; kapsamlı PDF kılavuzunu varsayılan görüntüleyici ile açabilir.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Uygulama içi yardım metinleri veya ekran görüntüleri güncellenecekse.
- PDF kılavuz dosya adı/yolu değişirse.
- Yeni kullanım kılavuzu sayfası eklenecekse.

---

### `src/ui/ozet.py`

Detaylı sözleşme özet dialogunu içerir. Bir sözleşme ailesinin ana sözleşme, SD sözleşmeler, sistemler, bileşenler ve teslimatlar açısından okunabilir raporunu üretir.

Yardımcı fonksiyonlar:

- `parse_iso_date()`, `display_date()`, `iso_display()`: Tarih gösterimi.
- `as_number()`, `fmt_num()`: Miktar formatlama.
- `norm_tr()`: Türkçe normalize.
- `status_kind()`, `status_label()`: Durum metnini kategori/etikete çevirme.
- `delivery_timing_text()`: Teslimat zamanlama metni üretme.

Sınıflar:

- `SummaryContext`: Özetlenecek sözleşme bağlamını taşır.
- `ContractSummaryDialog`: Büyük özet ekranıdır. Üst bar, meta bilgiler, durum kartları, tarih kartları, sistem kartları, bileşen tabloları, teslimat tabloları, kapsam seçimi ve detay açma işlemlerini yönetir.

Öne çıkan metot grupları:

- Veri yükleme: `load_data()`, `safe_load_context()`, `_initial_scope_from_item()`.
- UI kurulum: `build()`, `build_alt3_overview()`, `build_topbar()`, `build_footer()`.
- Kart üretimi: `alt3_side_card()`, `build_alert()`, `build_mini_card()`, `build_date_card()`, `build_count_card()`.
- Yenileme: `refresh_view()`, `refresh_contract_info_table()`, `refresh_alert()`, `refresh_dates()`, `refresh_systems()`, `refresh_components()`, `refresh_deliveries()`.
- Detay: `select_system()`, `open_detail()`, `clear_layout()`.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Sözleşme özet raporuna yeni alan/metrik eklenecekse.
- Özet ekran tasarımı veya kapsam seçimi değiştirilecekse.
- Sistem, bileşen veya teslimat tablolarının sunumu değişirse.

---

### `src/ui/tarih.py`

Sözleşme ve sistem terminlerini takvim görünümünde gösteren takip ekranıdır.

Fonksiyon ve sınıflar:

- `parse_calendar_iso_date()`: Takvim için tarih parse eder.
- `CalendarEventChip`: Gün hücresindeki küçük event chip'i; çift tıkla detay açabilir.
- `CalendarEventCard`: Sağ paneldeki event kartı; tıklama ile detay açabilir.
- `CalendarMorePopup`: Bir günde çok event varsa fazlasını popup olarak gösterir.
- `ContractCalendarWindow`: Ana takvim dialogudur.

`ContractCalendarWindow` özellikleri:

- Platform filtresi sunar.
- Aylık görünüm ve liste/yan panel istatistiklerini yönetir.
- Sözleşme terminleri ve sistem terminlerini ayrı event türleri olarak sınıflandırır.
- Bugüne dön, önceki/sonraki ay gezinmesi sağlar.
- Event detayını ana uygulamaya callback ile açtırabilir.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Takvimde gösterilecek event türleri değişirse.
- Termin renklendirme/sınıflandırma kuralları değişirse.
- Takvim görünümü veya platform filtresi genişletilecekse.

---

### `src/ui/theme.py`

Uygulamanın global Qt stylesheet metnini `STYLE` değişkeninde tutar. Buton, tablo, dialog, card, input, scroll bar, combo box ve çeşitli objectName tabanlı stiller bu dosyada tanımlıdır.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Uygulamanın genel renkleri, fontları, butonları veya tablo görünümleri değiştirilecekse.
- Yeni UI bileşenleri için objectName tabanlı stylesheet eklenecekse.

---

### `src/ui/toast.py`

Kısa süreli ekranda beliren toast bildirimi sağlar.

`ToastNotification` sınıfı:

- Verilen mesajı parent pencerenin sağ üstüne yerleştirir.
- `show_in()` ile görünür hale gelir ve süre sonunda kapanır.
- Başarı/bilgi tarzı geçici bildirimler için kullanılır.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Bildirim konumu, süresi, stili veya animasyonu değiştirilecekse.
- Hata/uyarı/başarı gibi farklı toast tipleri eklenecekse.

---

### `src/ui/widgets.py`

Küçük tekrar kullanılabilir widget yardımcıları içerir.

Fonksiyonlar:

- `stat_card(title, value)`: Başlık ve değer içeren küçük kart widget'ı üretir.
- `set_card_value(card, value)`: Kart içindeki değer label'ını günceller.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Dashboard/metrik kartlarının ortak tasarımı değişirse.
- Yeni basit widget factory fonksiyonları eklenecekse.

---

### `src/workers/__init__.py`

`src.workers.excel_workers` içeriğini dışa aktaran paket dosyasıdır. Worker sınıflarına `src.workers` üzerinden erişimi kolaylaştırır.

---

### `src/workers/excel_workers.py`

Excel okuma/yazma işlemlerini UI thread'i dışında yapmak için kullanılan worker sınıflarını içerir.

Yardımcı fonksiyonlar:

- `normalize_sheet_name()`, `is_system_sheet_name()`, `safe_sheet_name()`: Worker içinde de platform/sistem sayfa ayrımı için kullanılır.

Sınıflar:

- `ExcelLoadWorker`: Excel dosyasını yükler, platformları okur, etiket haritasını çıkarır ve sözleşme indeksini batch/progress sinyalleriyle üretir. Read-only hızlı indeks fonksiyonları içerir.
- `ComponentSaveWorker`: UI'dan gelen bileşen satırlarını `ComponentDef` listesine çevirir ve Excel'e yazar.
- `UserSaveWorker`: Kullanıcı listesini Excel'e yazar.
- `ContractSaveWorker`: Sözleşme ailesini kaydeder. Gerekirse ayrı store açar, versiyon/snapshot bilgisini kullanır ve sonuç bilgisini UI'a döndürür.
- `AnalyzeDialog`: Basit analiz/progress dialogu; çalıştırılan fonksiyonu progress bar ile sarar.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Excel yükleme performansı veya progress mesajları değişirse.
- Yeni asenkron kayıt işlemi eklenecekse.
- Worker sinyallerinin payload yapısı değiştirilecekse.

---

### `tools/generate_sts_ico.py`

Geliştirici yardımcı scriptidir. Harici görsel kütüphanesi kullanmadan piksel çizimiyle STS ikon dosyası üretir.

Fonksiyonlar:

- `blend()`, `put()`: Piksel renk karıştırma ve yazma.
- `rect()`, `rounded_rect()`, `line()`, `circle()`: Basit raster çizim primitive'leri.
- `draw_text()`: Mini bitmap font ile `STS` yazısı çizer.
- `make_image()`: Belirli boyutta ikon görselini üretir.
- `dib_for()`: Piksel matrisini Windows DIB verisine çevirir.
- `write_ico()`: Birden fazla boyuttaki ikon görsellerini `.ico` dosyası olarak yazar.
- `main()`: Varsayılan boyut setiyle `src/ui/assets/sts_icon.ico` üretir.

Bu dosyada değişiklik yapılması gereken tipik durumlar:

- Uygulama ikonu kodla yeniden üretilecekse.
- İkon renkleri, şekilleri veya boyut seti değiştirilecekse.

## 6. Özellikten Dosyaya Hızlı Rehber

- **Ana pencere / platform listesi / sözleşme listesi:** `app.py` içindeki `MainWindow`.
- **Sözleşme detay çalışma ekranı:** `app.py` içindeki `ContractWorkWindow`; yardımcılar `src/ui/contract/`.
- **Yeni sözleşme formu:** `app.py` içindeki `ContractDialog`.
- **Sözleşme düzenleme:** `app.py` içindeki `ContractEditDialog`.
- **Sistem ekleme/düzenleme:** `app.py` içindeki `SystemDialog` ve `MultiSystemDialog`.
- **Teslimat/kabul ekleme:** `app.py` içindeki `DeliveryDialog`, ayrıca `src/ui/contract/work_window_deliveries.py`.
- **Otomatik kabul:** `auto_accept.py`.
- **Excel okuma/yazma formatı:** `src/services/excel_store.py`.
- **Excel yükleme/kaydetme workerları:** `src/workers/excel_workers.py`.
- **Bileşen yönetimi:** `app.py` içindeki `ComponentManagerDialog`, kalıcılık `ExcelStore`.
- **Kullanıcı yönetimi:** `app.py` içindeki `UserManagerDialog`, kalıcılık `ExcelStore`.
- **Platform yönetimi ve logo:** `src/ui/dialogs/platforms.py`, kalıcılık `ExcelStore`.
- **Etiket yönetimi:** `app.py` içindeki `TagManagerDialog`, `TagAssignDialog`, kalıcılık `ExcelStore`.
- **Takvim/termin ekranı:** `src/ui/tarih.py`.
- **Sözleşme özet ekranı:** `src/ui/ozet.py` ve hızlı popup için `src/ui/dialogs/contract_summary_popup.py`.
- **Tarih seçici:** `src/ui/date_picker.py`.
- **Global tema:** `src/ui/theme.py`.
- **Toast bildirim:** `src/ui/toast.py`.
- **Kullanım kılavuzu:** `src/ui/kullanim_kilavuzu.py`.
- **Sürüm yönetimi:** `src/services/version_manager.py`.
- **Performans kayıtları:** `src/services/perf_tracker.py`.
- **SQLite cache hazırlığı:** `src/services/local_cache_db.py`.
- **Uygulama ikonu üretimi:** `tools/generate_sts_ico.py`.

## 7. Değişiklik Yaparken Dikkat Edilecek Noktalar

1. **Excel kolonlarını değiştirirken tek dosya yetmez.** `src/domain/constants.py`, `src/config/app_config.py`, `src/services/excel_store.py`, `src/workers/excel_workers.py` ve ilgili UI tabloları birlikte kontrol edilmelidir.
2. **Sözleşme blok yapısı hassastır.** Merge, toplam satırı, SD ekleme ve fiziksel satır insert işlemleri `ExcelStore` içinde birbirine bağlıdır.
3. **UI thread'i kilitlememeye dikkat edin.** Büyük Excel işlemleri worker ile yapılmalıdır.
4. **Tarih formatı ISO tutulur.** İç veri modelinde genellikle `YYYY-MM-DD` kullanılır; ekranda Türkçe/okunabilir format dönüştürülür.
5. **Platform olmayan sayfalar filtrelenmelidir.** Yeni teknik sayfa eklerseniz `CORE_SHEETS` veya `EXTRA_SYSTEM_SHEET_NAMES` listesine ekleyin.
6. **Etiket sayfası iki tip satır taşır.** Etiket tanımı ve sözleşme ataması aynı sayfada farklı `KayitTipi` değerleriyle tutulur.
7. **Dataclass alanı eklemek ExcelStore güncellemesi gerektirir.** Yeni alan UI formunda görünse bile Excel okuma/yazma metotları güncellenmezse kalıcı olmaz.
8. **Paketli Windows kullanımını unutmayın.** Dosya yolu değişikliklerinde `.exe` ve kaynak koddan çalıştırma senaryoları birlikte düşünülmelidir.

