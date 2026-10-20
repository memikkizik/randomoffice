# 🏢 RandomOffice — Ücretsiz ve Halka Açık Masaüstü & Web Office Paketi

> **Microsoft 365 veya Office abonelik ücreti ödemeden; Word (`.docx`), Excel (`.xlsx`) ve PowerPoint (`.pptx`) belgelerinizi doğrudan masaüstü uygulamasında veya web tarayıcınızda açın, düzenleyin, formülleri hesaplayın, slayt gösterisi yapın ve kaydedin!**

[![Web Office Suite](https://img.shields.io/badge/Office%20Suite-Word%20%7C%20Excel%20%7C%20PowerPoint-0078d4.svg)]()
[![Desktop Engine](https://img.shields.io/badge/Desktop-Native%20Edge%20WebView2-blue.svg)]()
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Zero Telemetry](https://img.shields.io/badge/Privacy-100%25%20Local%20%26%20Zero%20Telemetry-success.svg)]()
[![Zero Firewall Prompts](https://img.shields.io/badge/security-Zero%20Socket%20%2F%20No%20Defender%20Block-brightgreen.svg)]()
[![Tests](https://img.shields.io/badge/tests-204%20passed-success.svg)]()

---

## 📥 GitHub'dan İndirme ve Kurulum

### ZIP ile indirme

GitHub'da **Code → Download ZIP** seçeneğine tıklayın. ZIP dosyasını çıkarıp
proje klasörünü açın.

### Git ile indirme

```powershell
git clone https://github.com/memikkizik/randomoffice.git
cd randomoffice
```

### Gereksinimleri kurma

Windows'ta **Python 3.11 veya üzeri** ve
[uv](https://docs.astral.sh/uv/) kurulu olmalıdır:

```powershell
uv sync
```

Kurulumdan sonra uygulamayı [Başlatma ve Kullanım](#-başlatma-ve-kullanım)
bölümündeki komutlarla çalıştırabilirsiniz.

## 🎯 Projenin Amacı ve Vizyonu

Günümüzde Microsoft Office belgelerini (`.docx`, `.xlsx`, `.pptx`) açıp üzerinde çalışmak için kullanıcılar yüksek aylık/yıllık lisans ücretlerine (Microsoft 365 vb.) zorlanmaktadır.

**RandomOffice**, herkesin tamamen ücretsiz kullanabileceği halka açık, bağımsız bir **Masaüstü ve Web Office Süiti**dir:
- ❌ **Abonelik yok, lisans anahtarı yok, hesap zorunluluğu yok.**
- 🛡️ **Windows Defender Güvenlik Duvarı uyarısı yok:** Masaüstü modu yerel ağ portu açmadan (sıfır TCP soketi ile) doğrudan in-process çalışır; güvenlik duvarı engeline takılmaz.
- ⚡ **Orijinal Format Desteği:** DOCX, XLSX ve PPTX dosyalarını tek tıkla açar, düzenler ve orijinal ikili formatında kaydeder.
- 💻 **Hem Masaüstü Hem Web:** İster Windows masaüstü uygulaması olarak çift tıklayıp kullanın, ister web sunucusu olarak çalıştırın.

---

## 💻 3 Temel Program Tek Bir Pakette

| Uygulama | Desteklenen Dosya | Yapabilecekleriniz |
| :--- | :--- | :--- |
| **📝 Word (Belge)** | `.docx` | Zengin metin düzenleme (Kalın, İtalik, Altı Çizili, Renk Paleti, Vurgu), Başlık stilleri (H1, H2, H3), Tablo ve Resim ekleme, Yatay Cetvel, A4 baskı düzeni, Canlı sözcük/okuma süresi sayacı ve Ctrl+P yazdırma. |
| **📊 Excel (Tablo)** | `.xlsx` | Çok sayfalı tablolar (`Sayfa1`, `Sayfa2`, `+`), formül çubuğu (`fx`), hücre koordinatları (`A1`), dinamik formül motoru (`=SUM`, `=AVERAGE`, `=MIN`, `=MAX`, `=A1+B1`), ₺ para birimi ve % yüzde formatı. |
| **📽️ PowerPoint (Sunum)** | `.pptx` | Küçük resimli slayt kartları, yeni slayt/çoğaltma/silme, slayt temaları (Modern Beyaz, Koyu Gece, Lacivert, Zümrüt, Mercan), resim ekleme ve **F5 Tam Ekran Slayt Gösterisi Modu**. |

---

## 🚀 Başlatma ve Kullanım

### 1. Masaüstü Uygulaması Olarak Çalıştırma (Önerilen)
Masaüstünüzde otomatik oluşturulan **RandomOffice** kısayoluna çift tıklayabilir veya proje klasöründen şu dosyayı çalıştırabilirsiniz:
- 👉 **`Launch_RandomOffice.vbs`** (Sessiz, konsol penceresiz, doğrudan yerel masaüstü penceresi açar)
- Veya terminalden:
  ```powershell
  uv run randomoffice desktop
  ```

### 2. Web Sunucusu Olarak Çalıştırma
Tüm yerel ağda veya tarayıcıda çalıştırmak için:
```powershell
uv run uvicorn randomoffice.api.app:app --host 127.0.0.1 --port 8000
```
Tarayıcınızda açın:
👉 **[http://127.0.0.1:8000](http://127.0.0.1:8000)**

---

## 🌟 Öne Çıkan Gelişmiş Özellikler

1. **"Ne yapmak istiyorsunuz?" Akıllı Arama Çubuğu (Tell Me):**
   - Üst çubuktaki arama kutusuna *"yazdır"*, *"tablo"*, *"resim"*, *"kaydet"*, *"kalın"* gibi herhangi bir eylem yazıp Enter'a basarak doğrudan çalıştırabilirsiniz.
2. **Kusursuz Şerit (Ribbon) Sekmeleri:**
   - **Giriş:** Pano, Yazı Tipi, Renk Paletleri, Paragraf Hizalamaları, Stiller.
   - **Ekle:** Tablo, Resim, Bağlantı (Link), Sayfa Sonu, Yatay Çizgi, Tarih & Saat.
   - **Formüller (Excel):** Toplam, Ortalama, Min, Max, Para Birimi, Satır/Sütun Ekle ve Sil.
   - **Slayt Gösterisi (PowerPoint):** Baştan Başlat (F5), Geçerli Slayttan, Tema Seçici.
   - **Görünüm:** Cetvel Aç/Kapat, Koyu Tuval Modu, Tam Ekran ve Yakınlaştırma (%50 - %150).
3. **Masaüstü Dosya Köprüsü (Native Bridge):**
   - Masaüstü modunda Windows Dosya Gezgini Açma/Kaydetme iletişim kutuları doğrudan yerel işletim sistemi arayüzüyle çalışır.
4. **Sürükle & Bırak Desteği:**
   - Herhangi bir `.docx`, `.xlsx` veya `.pptx` dosyasını pencerenin üzerine bırakmanız açılması için yeterlidir.
5. **Otomatik Kaydetme:**
   - Yazdığınız her harf yerel tarayıcı önbelleğinde saklanır; tarayıcı kapansa dahi çalışmanız kaybolmaz.

---

## 🧪 Testler

Paket içindeki tüm motor, API, sözleşmeler ve gerçek tarayıcı senaryolarını çalıştırmak için:
```powershell
powershell -ExecutionPolicy Bypass -File scripts/check_project.ps1 -Scope All
```
*(204 testin tamamı %100 başarıyla geçmektedir; 44 gerçek Playwright/Edge tarayıcı senaryosu dahil.)*

---

## 📚 Dokümantasyon
Detaylı mimari, API kılavuzu ve manuel test belgeleri için [**`docs/`**](docs/) klasörüne göz atabilirsiniz:
- [**docs/DEVELOPMENT.md**](docs/DEVELOPMENT.md): Yerel ortam kurulumu ve derleme adımları.
- [**docs/MANUAL_TESTS.md**](docs/MANUAL_TESTS.md): Windows masaüstü doğrulama adımları.
- [**docs/COLLABORATION.md**](docs/COLLABORATION.md): Teknik mimari mutabakatları ve görev takibi.
- [**docs/BACKEND_REVIEW.md**](docs/BACKEND_REVIEW.md): OpenXML motor ve bağımlılık grafiği mimarisi.

---

## 🔒 Güvenlik, Gizlilik ve Açık Kaynak Güvencesi
- **Sıfır Telemetri & Sıfır Veri Toplama:** Belgeleriniz ve verileriniz asla dışarıdaki bir bulut sunucusuna gönderilmez. Tamamen yerel (air-gapped) mimaridedir.
- **Gizli API veya Sızıntı Riski Yok:** Kod tabanı hiçbir gizli API anahtarı, kişisel veri veya arka kapı içermez. MIT lisansı ile halka açıktır.
- **İnternet Bağımsız:** İnternet bağlantınız olmasa bile tüm masaüstü ofis özellikleri eksiksiz çalışır.
- **Güvenlik Duvarı Engeli Yok:** Yerel ağ dinlemesi yapmadan doğrudan in-process WebView2 motoruyla açılır.
