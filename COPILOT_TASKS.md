# RandomOffice - Copilot Görev Dağılım Listesi (Delegation Board)

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
- **Durum:** `[BEKLENİYOR]`

---

#### [GÖREV 2] Örnek API İstemcisi
- **Hedef Dosya:** `examples/api_client.py`
- **Açıklama:** RandomOffice REST API'sine dosya yükleyip dönüştürme sonucunu ekrana yazdıran basit bir örnek script yaz.
- **Gereksinimler:**
  - Standart `urllib.request` veya `httpx` kullanabilir.
  - `http://localhost:8000/health`, `/api/v1/parse` ve `/api/v1/convert` uçlarına örnek çağrılar yapsın.
- **Durum:** `[BEKLENİYOR]`

---

#### [GÖREV 3] README Dokümantasyon Örnekleri
- **Hedef Dosya:** `README.md`
- **Açıklama:** `README.md` dosyasına yeni REST API ve CLI kullanım komutları (`curl` örnekleri, Docker çalıştırma komutu) ekle.
- **Durum:** `[BEKLENİYOR]`
