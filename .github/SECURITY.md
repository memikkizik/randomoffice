# Güvenlik Politikası (Security Policy)

RandomOffice, kullanıcı gizliliğine ve güvenliğine en üst düzeyde önem veren açık kaynaklı bir ofis süitidir.

## Desteklenen Sürümler

Aşağıdaki sürümler şu anda güvenlik güncellemeleri ve düzeltmeleri almaktadır:

| Sürüm | Destek Durumu |
| :--- | :--- |
| `2.x` (Master/Main) | :white_check_mark: Destekleniyor |
| `< 2.0` | :x: Desteklenmiyor |

## Güvenlik İlkelerimiz

1. **Sıfır Telemetri:** Uygulama kullanıcı davranışlarını izlemez, analitik göndermez ve harici sunucularla veri paylaşmaz.
2. **Sıfır Soket Güvenliği:** Masaüstü modu yerel ağ dinlemesi yapmaz; ağ üzerinden izinsiz sızmalara karşı tam korumalıdır.
3. **Güvenli Dosya Ayrıştırma:** OpenXML arşivleri sıkıştırma oranı sınırlarına (Zip Bomb koruması) ve devre dışı bırakılmış dış XML varlıklarına (XXE koruması) tabidir.

## Güvenlik Açığı Bildirimi

Bir güvenlik açığı keşfettiyseniz, lütfen bunu herkese açık GitHub Issue bölümünde paylaşmak yerine sorumlu ifşa (responsible disclosure) ilkelerine uygun olarak bildirin:

- **E-posta:** [memikkizik@gmail.com](mailto:memikkizik@gmail.com)
- **Güvenlik Başlığı:** `[SECURITY] RandomOffice Güvenlik Bildirimi`

Lütfen açığın yeniden üretilme adımlarını, etkilenen bileşeni ve potansiyel etkiyi detaylandırın. Bildiriminiz en geç 48 saat içerisinde incelenecektir.

