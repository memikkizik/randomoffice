# RandomOffice — Manuel Kullanıcı Kontrol Listesi (Manual Test Matrix)

Bu belge, otomatik birim ve tarayıcı testlerinin ötesinde, gerçek Windows masaüstü (`RandomOffice.exe`) ve Web tarayıcı ortamlarında son kullanıcı deneyimini manuel olarak doğrulamak için hazırlanmış kontrol listesidir.

---

## 1. Test Ortamı ve Ön Koşullar

| Parametre | Ortam A (Kaynak Kod) | Ortam B (Paketlenmiş Standalone EXE) |
| :--- | :--- | :--- |
| **Başlatıcı** | `Launch_RandomOffice.bat` veya `Launch_RandomOffice.vbs` | `dist\RandomOffice\RandomOffice.exe` |
| **Python** | 3.12 (uv sanal ortamı) | Bağımsız (Gömülü Python çalışma zamanı) |
| **Görünüm Motoru** | Microsoft Edge WebView2 | Microsoft Edge WebView2 |
| **Ağ Durumu** | Tamamen yerel (In-process, sıfır port/soket) | Tamamen yerel (In-process, sıfır port/soket) |
| **Test Tarihi** | `2026-10-20` | `2026-10-20` |
| **Test Eden** | | |

---

## 2. Test Senaryoları & Adım Adım Kontrol Tablosu

### Kategori 1: Masaüstü Pencere Yönetimi & Güvenlik Duvarı
- [ ] **1.1 Güvenlik Duvarı & Defender Uyarısı Yokluğu**
  - **Adım:** Uygulamayı ilk kez çift tıklayarak çalıştırın.
  - **Beklenen Sonuç:** Windows Defender Güvenlik Duvarı "Ağ erişimine izin ver" penceresi AÇILMAMALIDIR. Uygulama sessizce ve doğrudan başlamalıdır.
  - **Gerçek Sonuç:** `[ ] Başarılı  [ ] Başarısız  (Not: ____________)`

- [ ] **1.2 960×640 Kompakt Ekran Uyumluluğu**
  - **Adım:** Pencereyi fareyle minimum boyutuna (960×640) küçültün.
  - **Beklenen Sonuç:** Alttaki mavi durum çubuğu (`.desktop-status-bar`) ekrandan taşmamalı veya kaybolmamalıdır; çalışma alanı kaydırılabilir kalmalıdır.
  - **Gerçek Sonuç:** `[ ] Başarılı  [ ] Başarısız  (Not: ____________)`

---

### Kategori 2: Dosya Diyalogları (Aç, Kaydet, İptal)
- [ ] **2.1 Yerel Dosya Açma & İptal Etme**
  - **Adım:** `Dosya` -> `Aç` (veya `Ctrl+O`) tıklayın. Windows Gezgini dosya seçicisi açıldığında hiçbir dosya seçmeden `İptal` (Cancel) butonuna basın.
  - **Beklenen Sonuç:** Uygulama donmamalı, çökmemeli; mevcut açık belge bozulmadan yerinde kalmalıdır.
  - **Gerçek Sonuç:** `[ ] Başarılı  [ ] Başarısız  (Not: ____________)`

- [ ] **2.2 Türkçe Karakterli Dosya Adı ile Kaydetme**
  - **Adım:** Üst başlıktaki dosya adını `Yıllık_Bütçe_Raporu_Örnek_Çağlayan.xlsx` yapın ve `Kaydet / İndir` butonuna basın.
  - **Beklenen Sonuç:** Windows Gezgini kaydetme penceresi bu Türkçe ismi doğru önermeli; kaydedilen dosya diske tam ve Türkçe karakter bozulması olmadan yazılmalıdır.
  - **Gerçek Sonuç:** `[ ] Başarılı  [ ] Başarısız  (Not: ____________)`

- [ ] **2.3 Kaydedilen Dosyayı Yeniden Açma (Roundtrip Doğrulaması)**
  - **Adım:** Kaydedilen dosyayı `Dosya` -> `Aç` ile tekrar açın.
  - **Beklenen Sonuç:** Tüm veriler, kalın/italik biçimler, formüller (`=SUM(...)`) veya slayt notları eksiksiz yüklenmelidir.
  - **Gerçek Sonuç:** `[ ] Başarılı  [ ] Başarısız  (Not: ____________)`

- [ ] **2.4 Sürükle-Bırak (Drag & Drop) ile Dosya Açma**
  - **Adım:** Masaüstünden veya klasörden bir `.docx`, `.xlsx` veya `.pptx` dosyasını sürükleyip uygulama penceresinin üzerine bırakın.
  - **Beklenen Sonuç:** Belge otomatik olarak algılanmalı, ilgili sekmeye geçilmeli ve içeriği ekrana gelmelidir.
  - **Gerçek Sonuç:** `[ ] Başarılı  [ ] Başarısız  (Not: ____________)`

---

### Kategori 3: Yazdırma Sistemi (Backstage Print Hub)
- [ ] **3.1 Canlı Baskı Önizleme**
  - **Adım:** `Dosya` -> `Yazdır` sekmesine geçin.
  - **Beklenen Sonuç:** Sağ tarafta canlı A4 kağıt üzerinde belgenin ölçeklenmiş baskı görünümü yer almalıdır.
  - **Gerçek Sonuç:** `[ ] Başarılı  [ ] Başarısız  (Not: ____________)`

- [ ] **3.2 Yönlendirme ve Boyut Değişimi**
  - **Adım:** Yönlendirmeyi `Dikey`den `Yatay`a getirin.
  - **Beklenen Sonuç:** Önizleme kutusu yatay A4 orantısına geçmeli; baskı esnasında `@media print` kuralları doğru oryantasyonu uygulamalıdır.
  - **Gerçek Sonuç:** `[ ] Başarılı  [ ] Başarısız  (Not: ____________)`

- [ ] **3.3 Yazıcı / PDF Çıktısı İzolasyonu**
  - **Adım:** `Şimdi Yazdır` butonuna basarak sistem yazdırma penceresini açın.
  - **Beklenen Sonuç:** Yalnızca aktif belge yazdırılmalı; diğer kapalı uygulamaların (Excel/PowerPoint) boş alanları veya Ribbon menüsü kağıda taşmamalıdır.
  - **Gerçek Sonuç:** `[ ] Başarılı  [ ] Başarısız  (Not: ____________)`

---

### Kategori 4: PowerPoint Sunum Modu & Klavye Kısayolları
- [ ] **4.1 Tam Ekran Sunum Başlatma (F5)**
  - **Adım:** PowerPoint sekmesine geçin ve `F5` tuşuna veya şeritteki `Slayt Gösterisi` butonuna basın.
  - **Beklenen Sonuç:** Karartılmış tam ekran sunum modu açılmalı; ilk slayt ortalanmış şekilde görüntülenmelidir.
  - **Gerçek Sonuç:** `[ ] Başarılı  [ ] Başarısız  (Not: ____________)`

- [ ] **4.2 Slayt Geçişleri (Ok Tuşları / Boşluk / Geri)**
  - **Adım:** `Sağ Ok` veya `Boşluk` ile ilerleyin, `Sol Ok` ile geri gelin.
  - **Beklenen Sonuç:** Slayt HUD sayacı (`1 / 2` -> `2 / 2`) güncellenmeli ve doğru slayt içeriği ekrana gelmelidir.
  - **Gerçek Sonuç:** `[ ] Başarılı  [ ] Başarısız  (Not: ____________)`

- [ ] **4.3 Sunumdan Çıkış (ESC)**
  - **Adım:** Sunum modundayken klavyeden `ESC` tuşuna basın.
  - **Beklenen Sonuç:** Sunum ekranı kapanmalı ve düzenleyici çalışma alanına sorunsuz dönülmelidir.
  - **Gerçek Sonuç:** `[ ] Başarılı  [ ] Başarısız  (Not: ____________)`

- [ ] **4.4 Konuşmacı Notları Düzenleme**
  - **Adım:** Slayt altındaki konuşmacı notları kutusuna özel bir not yazın ve slaytlar arasında geçiş yapıp geri dönün.
  - **Beklenen Sonuç:** Yazılan not aktif slayt modelinde korunmalı ve silinmemelidir.
  - **Gerçek Sonuç:** `[ ] Başarılı  [ ] Başarısız  (Not: ____________)`

---

## 3. Manuel Test Değerlendirme Özeti

- [ ] Kaynak Kod Sürümü (`.venv` üzerinden `Launch_RandomOffice.bat`): **Tamamlandı**
- [ ] Paket Sürümü (`dist\RandomOffice\RandomOffice.exe`): **Tamamlandı**
- **Sonuç:** Uygulama genel bütünlüğü, yerel köprü mimarisi ve kullanıcı deneyimi onaylandı.

