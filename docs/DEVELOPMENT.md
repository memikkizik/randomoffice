# Geliştirme ve doğrulama

## Başlatma

Kurulu Windows sürümü: masaüstündeki **RandomOffice** simgesi veya Başlat menüsü.
Geliştirme sürümü: `Launch_RandomOffice.bat`.

İlk kurulum/güncelleme (kararlı EXE yolu, kısayollar, Windows dosya türü kaydı):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/install_desktop.ps1
```

Kayıt mevcut varsayılanları değiştirmez. Uygulama **Ayarlar → Windows varsayılan
uygulamaları** düğmesiyle DOCX/XLSX/PPTX için kullanıcı seçimine izin verir.
Belgeye çift tıklama başlangıç dosyasını açar; Ctrl+S mevcut dosyaya yazar,
Ctrl+Shift+S yeni konum seçtirir. Yeni belgeler eski dosya kimliğini sıfırlar.
Kalıcı native taslak `%LOCALAPPDATA%\RandomOffice\workspace-state.json` altında,
WebView profili aynı uygulama dizininde, hata günlükleri `logs/` altındadır.

Tarayıcı sürümü:

```powershell
uv sync --extra test --extra browser
uv run randomoffice serve --host 127.0.0.1 --port 8000
```

Tarayıcıda `http://127.0.0.1:8000` adresini açın.

## Testler

```powershell
uv run pytest -o addopts= -q
```

Tarayıcı testleri önce yüklü Microsoft Edge'i kullanır. Edge bulunmazsa:

```powershell
uv run python -m playwright install chromium
```

Son doğrulama: **180 başarılı test**, bunların **44'ü gerçek tarayıcı testi**.
Playwright kurulu değilse tarayıcı test modülü atlanır; tam doğrulama için
`browser` ek bağımlılığı kurulmalıdır.

## Kaynakların ayrımı

## Windows masaüstü paketi oluşturma

```powershell
.venv\Scripts\python.exe -m PyInstaller --noconfirm --distpath dist/standalone --workpath build/standalone RandomOffice.spec
.venv\Scripts\python.exe scripts/test_desktop_build.py
.venv\Scripts\python.exe scripts/package_desktop.py
```

Çıktı `dist/standalone/RandomOffice.exe` tek dosyalıdır; ayrı `_internal`
klasörü gerektirmez. Test, EXE'yi kaynaklardan bağımsız bir klasöre kopyalar;
Python DLL, .NET, Office dışa aktarımı ve gizli gerçek WebView açılışını doğrular.
Sonuç `artifacts/desktop-isolated-self-test.json` içine yazılır. Paketleme aracı
EXE içinde Python DLL bulunduğunu kontrol eder ve indirme ZIP'ini yeniler.
İndirme ZIP'i derleme verilerine dahil edilmez.

Eski çok dosyalı EXE'nin ZIP içinden tek başına açılması `Failed to load Python
DLL` hatasına neden olabilir. Yeni tek dosyalı sürümü kullanın; ZIP dağıtımını
önce **Tümünü Ayıkla** ile çıkarın. İlk açılış birkaç saniye sürebilir.

## Arayüz kaynakları

- `randomoffice/static/index.html`: tasarım ve DOM yapısı.
- `randomoffice/static/editor.js`: etkileşimler, taslaklar, dosya işlemleri.
- `randomoffice/static/workspace.js`: profesyonel araçlar, native başlangıç/taslak,
  Bul/Değiştir, sayfa düzeni, hücre biçimi, veri araçları ve Windows Ayarları.
- `randomoffice/static/formulas.js`: sınırlı ve açık formül yorumlayıcısı.
- `randomoffice/static/editor.css`: işlevsel yerleşim düzeltmeleri.
- `randomoffice/ui.py`: web ve masaüstüne aynı davranış kodunu yerleştirir.
- `randomoffice/editor_formats.py`, `ppt_richtext.py`: gerçek Office biçim geçişi.
- `randomoffice/windows_integration.py`: mevcut varsayılanlara dokunmadan
  kayıt planı, kurulum durumu ve Windows varsayılan uygulama ayarları.

HTML'deki eski gömülü script uyumluluk için dosyada kalır; uygulama bunu
ortak düzenleyici koduyla değiştirir. Yeni davranışlar HTML içine eklenmemelidir.
DOM kimlikleri veya olay işleyicileri değişirse tarayıcı testleri çalıştırılmalıdır.
Dosya sahipliği ve ortak çalışma durumu `ANTIGRAVITY_COLLABORATION.md` içindedir.

## Kapsam

Temel belge içeriği ve düzenleyicinin desteklediği biçimler yeniden oluşturulur.
Tüm Office grafiklerinin, animasyonlarının, makrolarının veya gelişmiş belge
nesnelerinin eksiksiz korunması desteklenmez. Excel yorumlayıcısı aritmetik,
hücre referansları, SUM/TOPLA, AVERAGE/ORTALAMA, MIN, MAX ve COUNT işlevlerini
destekler; tüm Excel işlevlerini uygulamaz.

Otomatik kayıt bu cihazdaki bir taslaktır. Kalıcı Office dosyası oluşturmak için
Kaydet/İndir kullanılmalıdır. Önceden paketlenmiş EXE kaynak değişikliklerini
otomatik almaz; yeni sürüm için `RandomOffice.spec` ile yeniden paketlenmelidir.
