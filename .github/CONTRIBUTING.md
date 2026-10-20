# RandomOffice Katkı Kılavuzu (Contributing Guide)

RandomOffice'e katkıda bulunmak istediğiniz için teşekkür ederiz! Bu proje tamamen açık kaynaklı olup, topluluk katkılarını memnuniyetle karşılamaktadır.

## 🚀 Başlangıç ve Geliştirme Ortamı

1. **Depoyu Çatallayın (Fork) ve Klonlayın:**
   ```bash
   git clone https://github.com/KULLANICI_ADINIZ/randomoffice.git
   cd randomoffice
   ```

2. **Bağımlılıkları Yükleyin:**
   Projede paket yöneticisi olarak [uv](https://docs.astral.sh/uv/) kullanılmaktadır:
   ```powershell
   uv sync --all-extras
   ```

3. **Geliştirme Dalı Açın:**
   ```bash
   git checkout -b ozellik/yeni-ozellik-adi
   ```

## 🧪 Testleri Çalıştırma

Tüm değişikliklerin mevcut işlevleri bozmadığını doğrulamak için testleri çalıştırın:

```powershell
# Hızlı birim testleri
uv run pytest -v --ignore=tests/test_editor_browser.py --ignore=tests/test_professional_browser.py

# Tam proje doğrulaması (tarayıcı testleri dahil)
powershell -ExecutionPolicy Bypass -File scripts/check_project.ps1 -Scope All
```

## 📋 Katkı Kuralları

- **Gizlilik:** Kod tabanına hiçbir harici izleme (telemetri), reklam veya üçüncü taraf API anahtarı eklenemez.
- **Biçimlendirme & Sözdizimi:** Python PEP 8 standartlarına ve modüler mimariye sadık kalınmalıdır.
- **Testler:** Eklenen her yeni özellik veya hata düzeltmesi için `tests/` altına karşılık gelen bir birim/entegrasyon testi eklenmelidir.

## 📥 Çekme İsteği (Pull Request) Süreci

1. Değişikliklerinizi açık ve anlaşılır bir commit mesajıyla kaydedin (`feat: ...`, `fix: ...`, `docs: ...`).
2. Dalınızı çatalladığınız depoya push edin: `git push origin ozellik/yeni-ozellik-adi`.
3. GitHub üzerinden ana depoya (`memikkizik/randomoffice:main`) bir **Pull Request (PR)** açın.
4. Otomatik CI testlerinin başarıyla tamamlandığından emin olun.

