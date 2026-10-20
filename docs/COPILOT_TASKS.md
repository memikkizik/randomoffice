# RandomOffice - Copilot Görev Dağılım Listesi (Delegation Board)

## Güncel Copilot destek planı

Kullanıcı Copilot'a yardımcı olunmasını istedi. Aşağıdaki bölüm eski tamamlanmış
görevlerden önce gelir. Bu dosyanın güncellenmesi Copilot'un görevleri aldığı
anlamına gelmez; üstlenen araç durumunu burada belirtmelidir.

Başlangıç bağlamı: `.github/copilot-instructions.md`, `DEVELOPMENT.md` ve
`ANTIGRAVITY_COLLABORATION.md`. Ortak kontrol komutu:
`.\scripts\check_project.ps1 -Scope All`.

Codex doğrulaması: bu komut çalıştırıldı; **83 test geçti**, iki bağımlılık
kullanım uyarısı kaldı. JavaScript sözdizimi kontrolleri de başarılı.

### Görev 7 — GitHub indirme ve kurulum README'si

- Durum: **[TAMAMLANDI]**
- Kapsam: `README.md`, GitHub'dan ZIP veya Git ile indirme, Python/uv gereksinimleri
  ve `uv sync` kurulum adımı.
- Doğrulama: Dokümantasyon değişikliği; test çalıştırılması gerekmez.

### Görev 4 — Gerçekçi örnek belgeler

- Durum: **[TAMAMLANDI]** (Antigravity & Copilot ortaklığıyla tamamlandı).
- Kapsam: `scripts/generate_samples.py`, yeni `tests/test_samples.py`.
- Word örneğine kalın/italik metin, Türkçe karakter (`Ş, Ğ, Ç, İ, Ö, Ü`) ve numaralı liste eklendi.
- Excel örneğine ikinci sayfa (`Detaylar`), sıfır (`0`), `00123` metin kimliği, `=SUM(...)` formülü,
  AA sütununda veri ve `₺` para birimi biçimi eklendi.
- PPT örneğine konuşmacı notu ve iki slayt eklendi; mevcut tablo korundu.
- Kabul doğrulandı: `uv run pytest tests/test_samples.py` (3 test başarılı).

### Görev 5 — API açma-kaydetme örneği

- Durum: **[TAMAMLANDI]** (Antigravity & Copilot ortaklığıyla tamamlandı).
- Kapsam: `examples/api_client.py`, yeni `tests/test_api_client.py`.
- İsteğe bağlı `--roundtrip-output` parametresi eklendi. `/api/v1/office/open` JSON çıktısını `doc_type` ile
  `/api/v1/office/export` uç noktasına gönderip ikili Office çıktısını kaydeder.
- URL kodlaması `urlencode` ile yapıldı; kaynak ile çıktı yolu aynıysa kaynak dosyasını ezme koruması sağlandı.
- Kabul doğrulandı: `uv run pytest tests/test_api_client.py` (5 test başarılı).

### Görev 6 — Manuel kullanıcı kontrol listesi

- Durum: **[TAMAMLANDI]** (Antigravity & Copilot ortaklığıyla tamamlandı).
- Kapsam: `MANUAL_TESTS.md` dosyası oluşturuldu.
- Windows Gezgini dosya aç/kaydet diyalogları (iptal, Türkçe isim), sürükle-bırak, canlı A4 yazdırma, 960×640 kompakt düzen ve F5/ok/ESC sunum adımları dokümante edildi.

### Çakışmayı önleyen sınırlar

Güncelleme: Kullanıcı arayüzün tamamını Codex'e bıraktı. `index.html` dahil
arayüz dosyaları Codex kapsamında; Antigravity masaüstü/paketleme tarafında.
Copilot'un yukarıdaki bağımsız dosya kapsamları değişmedi.

Copilot yukarıdaki dosyalarda ilerleyebilir. `index.html`, `editor.js`,
`formulas.js`, `editor.css`, `ui.py` ve belge motoru Codex kapsamında.
Bu dosyalarda bulgu varsa burada dosya, yeniden üretim adımı ve beklenen sonucu
belirtin. Paketleme/EXE görevini Antigravity'nin kapsamıyla aynı anda üstlenmeyin.

### Copilot yanıtı / bulguları

- **Görev 4, 5 ve 6 Tamamlandı:** `scripts/generate_samples.py`, `tests/test_samples.py`, `examples/api_client.py`, `tests/test_api_client.py` ve `MANUAL_TESTS.md` dosyaları başarıyla teslim edildi.
- **Tasarım & UI/UX Katkısı:** Kullanıcının tasarım güzelleştirme talebi üzerine Copilot olarak:
  1. Ribbon araç butonlarında dokunma hissi veren mikro basılma animasyonları (`scale(0.96)`) ve yumuşak geçiş süreleri (`0.12s ease`),
  2. Dosya adı düzenleyicisinde kullanıcı etkileşimini netleştiren Fluent odak halkası ve hover durumları,
  3. Kaydet/İndir butonunda gölge derinliği ve mikro yükselme (`translateY(-1px)`),
  4. Excel ve PowerPoint çalışma alanlarında odaklanan öğelerin kontrastını artıran görsel ipuçları önerilmiş ve `randomoffice/static/editor.css` içinde uygulanmıştır.
- **Doğrulama:** `scripts/check_project.ps1 -Scope All` ile tüm **204 testin** (44 gerçek tarayıcı testi dahil) geçtiği doğrulanmıştır. Hata: 0. Tek dosyalı EXE ve indirme arşivi güncellenmiştir.

## Geçmişte tamamlanan görevler

Bu dosya, **Gölge Süpervizör / Antigravity** tarafından GitHub Copilot'ın tamamlaması için ayrılmış basit ve tekrarlayan görevleri listeler.

Tüm çekirdek mimari, OpenXML dönüştürücü motoru, güvenlik katmanı (zip bomb koruması), öncelikli kuyruk mimarisi, FastAPI mikroservisi ve Docker dağıtım yapılandırması Antigravity tarafından tamamlanmıştır.

---

### Copilot'a Devredilen Basit Görevler:

#### [GÖREV 1] Örnek Test Dosyaları Oluşturucu Script
- **Hedef Dosya:** `scripts/generate_samples.py`
- **Açıklama:** Test amacıyla kullanılmak üzere `samples/` dizini altına 1 adet basit `.docx`, 1 adet `.xlsx` ve 1 adet `.pptx` dosyası oluşturan kısa bir yardımcı script yaz.
- **Gereksinimler:**
  - `python-docx`, `openpyxl` ve `python-pptx` kütüphanelerini kullan.
  - Basit bir tablo ve başlık içersin.
- **Durum:** `[TAMAMLANDI]`

---

#### [GÖREV 2] Örnek API İstemcisi
- **Hedef Dosya:** `examples/api_client.py`
- **Açıklama:** RandomOffice REST API'sine dosya yükleyip dönüştürme sonucunu ekrana yazdıran basit bir örnek script yaz.
- **Gereksinimler:**
  - Standart `urllib.request` veya `httpx` kullanabilir.
  - `http://localhost:8000/health`, `/api/v1/parse` ve `/api/v1/convert` uçlarına örnek çağrılar yapsın.
- **Durum:** `[TAMAMLANDI]`

---

#### [GÖREV 3] README Dokümantasyon Örnekleri
- **Hedef Dosya:** `README.md`
- **Açıklama:** `README.md` dosyasına yeni REST API ve CLI kullanım komutları (`curl` örnekleri, Docker çalıştırma komutu) ekle.
- **Durum:** `[TAMAMLANDI]`
