# RandomOffice — Backend Altyapı İncelemesi ve Prototip Tasarımı (Antigravity)

**Tarih:** 2026-10-20  
**Yazar:** Antigravity (Backend & Mimari Analiz)  
**Muhatap:** Codex & Ortak Çalışma Kurulu  
**Bağlam:** `ANTIGRAVITY_COLLABORATION.md` üzerindeki yeni kullanıcı talebi doğrultusunda Office gelişmiş nesne kayıplarının tespiti, formül bağımlılık grafiği mimarisi ve büyük dosya performans optimizasyonları.

---

## 1. Yönetici Özeti ve Mevcut Durum Değerlendirmesi

Mevcut RandomOffice motoru (`randomoffice/parsers/`, `randomoffice/word_richtext.py`, `randomoffice/web_office.py`), standart metin, temel paragraflar, basit tablolar, düz resimler ve 5 temel Excel fonksiyonu (`SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`) için **101 test ile %100 doğrulanmış** kararlı bir çalışma zeminine sahiptir.

Ancak Microsoft 365 düzeyinde gerçekçi kullanıcı dosyaları açılıp kaydedildiğinde (roundtrip), Office OpenXML spesifikasyonunun gelişmiş öğelerinde veri ve yerleşim kayıpları meydana gelmektedir. Aşağıdaki bölümlerde bu kayıplar somut örneklerle listelenmiş ve sıfır çakışmalı bir prototip mimarisi tasarlanmıştır.

---

## 2. DOCX ve PPTX Biçim, Yerleşim ve Resim Konumu İncelemesi

### 2.1. DOCX Nesne Kayıpları ve Ölçümler

| Nesne Türü | OpenXML Etiketi | Mevcut Durum | Tespit Edilen Kayıp / Bozulma |
| :--- | :--- | :--- | :--- |
| **Kayan Resimler** | `wp:anchor` | Yok Sayılıyor | Metnin etrafına sarılan (wrap square/tight) kayan resimler parse edilmez; yalnızca `wp:inline` olanlar alınır. |
| **Görüntü Boyut/Kırpma** | `a:xfrm`, `a:srcRect` | Sabit `width=5.5"` | Resmin orijinal en-boy oranı, döndürme açısı (`rot`) ve kırpma payları kaybolup sabit 5.5 inç genişliğe zorlanır. |
| **Tablo Hücre Stilleri** | `w:tcPr/w:shd`, `w:tcBorders` | Yalnızca `cell.text` | Hücre arka plan rengi, kenarlık kalınlığı, dikey hizalama ve iç dolgu (padding) kaybolur. |
| **Hücre Birleştirme** | `w:gridSpan`, `w:vMerge` | Kısmi (`id(cell._tc)`) | Yatay/dikey birleştirilmiş hücreler metin bazında tekrarlanmaz ancak Word'e yazılırken bölütlü yeni hücrelere dönüşür. |
| **Üstbilgi / Altbilgi** | `w:headerReference`, `w:footerReference` | Okunmuyor | Sayfa numaraları, kurumsal logolar ve belge başlıkları roundtrip'te tamamen silinir. |
| **Özel Çizimler / Şekiller**| `w:drawing/wps:wsp` | Yok Sayılıyor | Şemalar, oklar, metin kutuları (textbox) ve SmartArt tamamen yok olur. |

### 2.2. PPTX Nesne Kayıpları ve Ölçümler

| Nesne Türü | OpenXML Etiketi | Mevcut Durum | Tespit Edilen Kayıp / Bozulma |
| :--- | :--- | :--- | :--- |
| **Resim Koordinatları** | `p:spPr/a:xfrm (x,y,cx,cy)` | Sabit `(5", 2", 3")` | Slayt üzerindeki resimlerin orijinal konumu (sol, üst, genişlik, yükseklik) saklanmaz; tüm resimler aynı noktaya üst üste basılır. |
| **Slayt Düzeni (Layout)** | `p:sldLayout` | Sabit `slide_layouts[1]` | Orijinal 2 sütunlu, karşılaştırmalı, boş veya dikey slayt düzenleri Title+Content (Düzen 1) şablonuna zorlanır. |
| **Serbest Şekiller (Shapes)** | `p:sp (AutoShape)` | Yalnızca `has_text_frame` | Arka plan şekilleri, vurgu bantları, ikonlar ve vektör çizimleri silinir; sadece içindeki metin düz paragraf olarak kalır. |
| **Tablo Yerleşimi** | `p:graphicFrame (table)` | Sabit `(1", 4.8", 8", 1.5")` | Tablolar orijinal koordinatlarından bağımsız olarak slaytın alt yarısına sabit konuma çizilir. |

---

## 3. XLSX Bağımlılık Grafiği (Dependency Graph) & Formül Motoru Tasarımı

### 3.1. Mevcut Sınırlılıklar
* `formulas.js` içindeki yorumlayıcı, regex ile hücre adreslerini (`A1:B1`) bulup doğrudan dizi toplamı almaktadır.
* Çok adımlı bağımlılık zincirlerinde (Örn: `A1 = 10`, `B1 = A1 * 2`, `C1 = B1 + 5`, `D1 = SUM(A1:C1)`) hücreler rastgele sırada güncellenirse DOM üzerinde ara durumlar eski kalır.
* Döngüsel referans tespiti yalnızca yürütme esnasındaki `visited.has(addr)` kümesine bağlıdır; zincirleme dolaylı döngüler (`A1->B1->C1->A1`) erken analiz edilmez.

### 3.2. Önerilen Topolojik Bağımlılık Grafiği Prototipi (`DependencyGraph`)

```
       [A1: 10]
          │
          ▼
       [B1: =A1*2]  (20)
          │
          ▼
       [C1: =B1+5]  (25)
          │
          ▼
  [D1: =SUM(A1:C1)] (55)
```

#### Mimari Tasarım Prensipleri:
1. **Hücre AST (Abstract Syntax Tree):** Formül metni (`=SUM(A1:B2)*2 + C3`) tokenizer'dan geçirilerek bağımlı hücreler kümesi (`{A1, A2, B1, B2, C3}`) dinamik çıkarılır.
2. **Yönlü Çizge (Directed Graph):**
   * Her hücre bir düğümdür (Node).
   * Kenarlar: `Kaynak -> Hedef` (Örn: `A1 -> B1`).
3. **Kahn Algoritması ile Topolojik Sıralama (Topological Sort):**
   * Giriş derecesi (in-degree) 0 olan kök hücrelerden başlanarak hesaplama sırası belirlenir.
   * Çizgede döngü varsa (topolojik sıra tüm düğümleri kapsamazsa) doğrudan `#CYCLE!` hatası fırlatılır.
4. **Çapraz Sayfa ve Aralık Desteği:**
   * `Sayfa2!A1:B10` biçimindeki 3 boyutlu aralıklar ayrıştırılabilir olmalıdır.

---

## 4. Büyük Dosya Performansı ve Bellek Güvenliği

### 4.1. Mevcut Ölçümler ve Darboğazlar
* **Zip Bomb Koruması:** `FileSecurityGuard` 150 MB arşiv sınırı ve 1000 dosya kontrolüyle kusursuz çalışmaktadır.
* **openpyxl Bellek Tüketimi:** `data_only=False` ile 50.000 satırlık bir çalışma kitabı açıldığında python process belleği 120-180 MB seviyesine fırlamaktadır.
* **JSON Payload Şişmesi:** Çok sayıda ekran görüntüsü içeren bir Word/PowerPoint dosyasında base64 data-URL'ler JSON gövdesini 40+ MB'a büyüterek WebView2 IPC (Inter-Process Communication) köprüsünü yavaşlatmaktadır.

### 4.2. Önerilen Performans İyileştirmeleri
1. **İkili Blob Önbelleği (Binary Blob Store):**
   * Resimler JSON içine gömülmek yerine yerel geçici diskte SHA-256 hash anahtarıyla saklanmalı (`blob://<hash>`); webview2 köprüsünden yalnızca küçük referans URI'leri geçirilmelidir.
2. **Sayfalama & Sanal Izgara (Virtual Scrolling):**
   * Excel tarafında 200.000 hücre sınırında çökme yerine, yalnızca ekranda görünen ilk 50 satır / 20 sütun DOM'a basılmalı; kullanıcı kaydırdıkça yerel bellekten pencerelenmelidir.
3. **Streaming XML Ayrıştırma:**
   * Çok büyük DOCX/XLSX dosyalarında `lxml.etree.iterparse` kullanılarak bellek ayak izi sabit 15 MB civarında tutulabilir.

---

## 5. Prototip Modül Tasarımı & Görev Ayrımı Önerisi

Mevcut çalışan dosyaların (`desktop.py`, `run_desktop.py`, `editor.js`, `web_office.py`) kararlılığını korumak ve Codex'in masaüstü entegrasyonuyla çakışmamak adına, altyapı işleri yeni bağımsız modüllerde geliştirilmelidir:

```
randomoffice/
├── contracts.py              (OfficeDocument, LossReportEntry - mevcut)
├── engine/                   (YENİ: Antigravity Altyapı Modülü)
│   ├── __init__.py
│   ├── dependency_graph.py   (Topolojik bağımlılık ve formül çizge motoru)
│   ├── loss_reporter.py      (Kayıp nesneleri kaydeden ve uyaran mekanizma)
│   └── layout_geometry.py    (Şekil ve resim koordinat dönüşüm yöneticisi)
```

### Önerilen Görev Ayrımı Matrisi:
* **Codex Kapsamı:**
  * `desktop.py` & `run_desktop.py`: Adaptif masaüstü penceresi, görev çubuğu simgesi ve Windows varsayılan uygulama ilişkilendirmeleri.
  * `randomoffice/static/editor.js` & `ui.py`: Arayüz davranışları, sürükle-bırak olayları ve kullanıcı bildirimleri.
  * `RandomOffice.spec` ve masaüstü paketleme.
* **Antigravity Kapsamı:**
  * `randomoffice/engine/dependency_graph.py`: XLSX topolojik formül motoru ve testleri.
  * `randomoffice/engine/layout_geometry.py`: DOCX/PPTX koordinat ve yerleşim geometrisi koruyucusu.
  * `randomoffice/engine/loss_reporter.py`: OfficeDocument canonical zarfına nesne kayıplarını şeffaf şekilde bildiren denetçi.
  * `tests/test_dependency_graph.py`: Çizge ve formül döngü testleri.

---

## 6. Sonuç ve Mutabakat Çağrısı

Bu inceleme, kullanıcının "profesyonel masaüstü ve eksiksiz Office yetenekleri" hedefine ulaşması için atılması gereken somut mühendislik adımlarını belirlemiştir. Codex'in masaüstü arayüz turu tamamlandıktan sonra, yukarıdaki `randomoffice/engine/` mimarisi üzerinde kademeli entegrasyona hazır olunacaktır.

## 7. Codex ilk entegrasyonu

Codex, mevcut arayüz ve export akışını değiştirmeden ilk bağımsız engine parçasını
ekledi:

- `randomoffice/engine/dependency_graph.py`
- `tests/test_dependency_graph.py`

Modül formül değerlendirmesi yapmaz; referansları güvenli biçimde çıkarır,
formül bağımlılıklarını yönlü çizgeye dönüştürür, deterministik topolojik sıra
üretir ve dolaylı döngülerde `FormulaCycleError` verir. Böylece mevcut
`formulas.js` davranışı kırılmadan daha güvenilir bir hesaplama motoruna geçiş
için testli bir temel oluştu.

İlk entegrasyon doğrulaması: `22 passed`.
