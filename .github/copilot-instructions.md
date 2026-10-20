# RandomOffice: Copilot çalışma bağlamı

Önce `COPILOT_TASKS.md`, `ANTIGRAVITY_COLLABORATION.md` ve `DEVELOPMENT.md`
dosyalarını okuyun. Güncel görev ayrımı, eski tarihsel notlardan önce gelir.
Bir görevi üstlendiğinizde görev panosunda durumunu ve değiştireceğiniz dosyaları
yazın. Başka aracın görevi aldığını veya incelemeyi tamamladığını varsaymayın.

## Uygulama yapısı

- Python/FastAPI belge motoru ve web API'si; pywebview masaüstü köprüsü.
- `randomoffice/static/index.html` tasarım ve DOM yapısıdır.
- Çalışan davranış kodu `editor.js` ve `formulas.js`, yerleşim düzeltmeleri
  `editor.css` içindedir. `ui.py` bunları web ve masaüstüne birlikte yerleştirir.
- HTML içindeki eski gömülü JavaScript uygulamada değiştirilir. Oraya işlev
  düzeltmesi yazmayın; çalışmayacaktır.
- DOM kimliklerini ve olay işleyici adlarını değiştirmek entegrasyonu etkiler.
- Word dışa aktarımı `html` alanını; PPT `notes`, `tables`, `images`, `theme`
  alanlarını; XLSX `number_formats` alanını korumalıdır.
- Formüller ve hücre değerleri modelde tutulur. Gösterilen hesap sonucunu
  formülün üzerine yazmayın; para birimi için sayıyı metne çevirmeyin.

## Birlikte çalışma

- Çalışma ağacı önceden değiştirilmiş durumdadır. Kullanıcıya veya diğer
  araçlara ait değişiklikleri geri almayın; toplu formatlama yapmayın.
- Copilot için ayrılan dosya kapsamları `COPILOT_TASKS.md` içindedir.
- Ortak dosyada değişiklik gerekiyorsa önce bulguyu ve önerilen dosya ayrımını
  panoya yazın. Bağımsız görevlerde ilerlemeye devam edin.
- Türkçe metinleri UTF-8 olarak koruyun. Üretilmiş önbellek, sanal ortam,
  örnek ikili dosya veya EXE dosyalarını kaynak değişikliğiyle karıştırmayın.

## Doğrulama ve teslim

```powershell
.\scripts\check_project.ps1 -Scope All
```

Yalnızca belge motoru: `-Scope Core`; yalnızca kullanıcı akışları: `-Scope Browser`.
Kurulum ve Edge/Chromium gereksinimleri `DEVELOPMENT.md` içinde açıklanır.
Son kaydedilen temel sonuç 83 testtir; bunu yeni bir test çalıştırmasının sonucu
olarak sunmayın. Gerçek komutu, sonucu, değişen dosyaları ve kalan kısıtları
görev panosuna ekleyin. Başarısız testleri silerek veya atlayarak işi tamamlanmış
göstermeyin. Yeni bir EXE oluşturulmadıysa kaynak kodun çalışmasını EXE'nin
doğrulanmasıyla eş tutmayın.
