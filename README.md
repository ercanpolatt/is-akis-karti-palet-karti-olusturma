# Palet Tanıtım Kartı ve İş Akış Yönetim Sistemi

Üretim ve sevkiyat süreçlerinde palet tanıtım kartlarını dinamik olarak oluşturan, A4 sayfa yapısına göre düzenleyen ve yazdırma/PDF işlemlerini otomatikleştiren Excel tabanlı yönetim sistemi.

## 🚀 Özellikler

- **Veri Giriş Arayüzü (`VERI_GIRIS`):** Parti/lot, sipariş, sevk tarihi, koli adedi ve kalite durumu gibi bilgilerin girilebildiği kullanıcı dostu form.
- **Dinamik Kart Üretimi (`PALET_KARTLARI`):** Girilen palet sayısı kadar kartı (Örn: 1/12, 2/12...) otomatik olarak oluşturur.
- **A4 Optimize Baskı Formatı:** Her A4 sayfasına tam 2 adet kart sığacak şekilde satır yükseklikleri, kenar boşlukları ve makas kesim çizgileri ayarlanmıştır.
- **Tek Tıkla PDF Kaydetme:** Kartları otomatik olarak kullanıcının masaüstüne tarih-saat damgalı PDF olarak aktarır.
- **Veri Doğrulama (Açılır Liste):** Hatalı veri girişini önlemek için kasa içi adet ve kalite onay alanlarında açılır liste desteği.

## 📋 Gereksinimler

- Windows işletim sistemi
- Microsoft Excel (Makro destekli)
- Python 3.8+ (Excel dosyasını sıfırdan oluşturmak veya güncellemek isterseniz)

### Python Bağımlılıkları

```bash
pip install -r requirements.txt
```

## 🛠️ Kullanım

1. **Excel Dosyasını Sıfırdan Oluşturmak / Güncellemek İçin:**
   ```bash
   python build_pallet_system.py
   ```
   Bu komut çalıştırıldığında `Palet_Karti_Sistemi.xlsm` dosyası otomatik olarak oluşturulur ve içine gerekli VBA modülleri, butonlar ve ilk kart şablonları yerleştirilir.

2. **Günlük Kullanım:**
   - `Palet_Karti_Sistemi.xlsm` dosyasını açın.
   - Sol taraftaki parti ve ürün bilgilerini doldurun.
   - Basılacak toplam kart sayısını girin.
   - **KARTLARI GÖSTER** butonuna basarak kartları inceleyin veya **YAZDIR** / **PDF OLARAK KAYDET** butonlarını kullanın.
