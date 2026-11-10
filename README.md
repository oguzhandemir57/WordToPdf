# Word to PDF Dönüştürücü (.NET Windows Forms)

Bu uygulama, seçtiğiniz bir klasör içerisindeki tüm Word belgelerini (`.docx`, `.doc`, `.docm`) otomatik olarak bulur ve **Masaüstünüzde `PdfDosyaları`** adında bir klasör oluşturup tüm belgeleri aynı dosya isimleriyle PDF olarak kaydeder.

---

## 🚀 Özellikler

1. **Kolay Klasör Seçimi ve Sürükle-Bırak:**
   - "📁 Klasör Seç..." butonunu kullanabilir veya dönüştürmek istediğiniz klasörü doğrudan pencereye sürükleyip bırakabilirsiniz.
2. **Masaüstünde `PdfDosyaları` Klasörü:**
   - Hedef klasör (`Desktop\PdfDosyaları`) otomatik oluşturulur.
   - Dosya adları birebir aynı kalır (`Rapor.docx` -> `Rapor.pdf`).
3. **Akıllı Filtreleme:**
   - Word açıkken oluşan geçici kilit dosyalarını (`~$...`) otomatik olarak filtreler, hataları önler.
   - İsteğe bağlı olarak "Alt klasörleri de tara" seçeneği sunar.
4. **Çift Motor Desteği (Dual Engine):**
   - **Microsoft Word (Office COM):** Bilgisayarınızda Microsoft Office/Word kuruluysa tam uyumlu, orijinal sayfa düzeni, yazı tipi ve tabloları koruyarak dönüştürür.
   - **LibreOffice:** Alternatif olarak LibreOffice kuruluysa otomatik algılanıp kullanılır.
5. **Modern Arayüz ve Gelişmiş Takip:**
   - Dosya listesi, dosya boyutları, anlık durum ve hata bilgileri listelenir.
   - İlerleme çubuğu (Progress Bar) ve yüzdelik oran.
   - İstediğiniz zaman işlemi durdurabileceğiniz "⏹ İptal Et" butonu.
   - İşlem bittiğinde tek tıkla hedef klasörü açma seçeneği.

---

## 🛠️ Çalıştırma

### Yöntem 1: Doğrudan .exe ile Çalıştırma
`C:\Users\ouz57\.gemini\antigravity\scratch\WordToPdfApp\publish\WordToPdfApp.exe` dosyasını veya ana klasördeki `Baslat.bat` dosyasını çift tıklayarak çalıştırabilirsiniz.

### Yöntem 2: Terminal / Komut Satırından Çalıştırma
```powershell
cd C:\Users\ouz57\.gemini\antigravity\scratch\WordToPdfApp
dotnet run
```
