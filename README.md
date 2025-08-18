# Excel AI Assistant - Excel Yapay Zeka Yardımcısı

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Office Add-in](https://img.shields.io/badge/Office%20Add--in-Excel-blue.svg)](https://docs.microsoft.com/en-us/office/dev/add-ins/)

**İlk açık kaynak kodlu local yapay zeka destekli Excel yardımcı aracı**

The first open-source local AI-powered Excel assistant tool

## 🌟 Özellikler / Features

### 🤖 AI Entegrasyonu / AI Integration
- **Local AI Model Desteği**: LM Studio ile entegrasyon
- **Gerçek Zamanlı Yanıtlar**: ChatGPT benzeri streaming yanıtlar
- **Çoklu Model Seçimi**: Yüklü AI modelleri arasından seçim
- **Akıllı Veri Analizi**: Otomatik veri analizi ve Excel işlemleri

### 📊 Excel İşlemleri / Excel Operations
- **Otomatik Grafik Oluşturma**: Veriye göre akıllı grafik seçimi
- **Akıllı Formatlama**: Otomatik tablo formatlaması ve renklendirme
- **Veri Filtreleme**: Gelişmiş filtreleme ve sıralama
- **Hesaplama**: Otomatik toplam, ortalama ve istatistik hesaplamaları
- **Trend Analizi**: Veri trendlerini otomatik tespit etme
- **Anomali Tespiti**: Veri anormalliklerini bulma

### 💬 Sohbet Arayüzü / Chat Interface
- **Modern Dark Theme**: Cursor benzeri koyu tema
- **Sohbet Geçmişi**: Kalıcı sohbet kayıtları
- **Otomatik Kaydırma**: AI yazarken otomatik ekran kaydırma
- **Yanıt Durdurma**: AI yanıtını istediğiniz zaman durdurma
- **Ses ve Görsel Giriş**: Ses komutları ve resim yükleme desteği

### ⚙️ Ayarlar ve Özelleştirme / Settings & Customization
- **Tema Seçimi**: Koyu, açık ve mavi tema seçenekleri
- **Excel Ayarları**: Grafik türü, boyut, otomatik formatlama
- **AI Model Yönetimi**: Model bağlantı testi ve güncelleme
- **Kullanıcı Tercihleri**: Kişiselleştirilebilir ayarlar

## 🚀 Kurulum / Installation

### Gereksinimler / Requirements
- Microsoft Excel (Desktop veya Online)
- LM Studio (Local AI model server)
- Modern web tarayıcısı

### Adımlar / Steps

1. **LM Studio Kurulumu**
   ```bash
   # LM Studio'yu indirin ve kurun
   # https://lmstudio.ai/
   ```

2. **Proje Kurulumu**
   ```bash
   git clone https://github.com/[username]/excel-ai-assistant.git
   cd excel-ai-assistant
   npm install
   ```

3. **AI Model Yapılandırması**
   ```bash
   # LM Studio'da model yükleyin
   # API sunucusunu başlatın (port 1234)
   ```

4. **Excel Add-in Kurulumu**
   ```bash
   npm run start
   # Excel'de Developer > Add-ins > Upload My Add-in
   ```

## 🔧 Yapılandırma / Configuration

### LM Studio Bağlantısı
```javascript
// ai_client.js
const baseUrl = 'http://192.168.1.5:1234'; // Kendi IP adresinizi girin
```

### Excel Ayarları
```javascript
// Excel ayarları localStorage'da saklanır
{
  "chartType": "ColumnClustered",
  "autoFormatting": true,
  "zebraRows": true,
  "headerHighlighting": true
}
```

## 📖 Kullanım / Usage

### Temel Komutlar / Basic Commands
- **"Bu veriyi analiz et"** - Seçili veriyi akıllı analiz
- **"Grafik oluştur"** - Otomatik grafik oluşturma
- **"Toplam hesapla"** - Sütun toplamları
- **"Filtrele"** - Akıllı veri filtreleme
- **"Formatla"** - Otomatik tablo formatlaması

### Gelişmiş Özellikler / Advanced Features
- **Ses Komutları**: Mikrofon ile komut verme
- **Resim Analizi**: Resim yükleyerek AI analizi
- **Sohbet Geçmişi**: Önceki sohbetleri yeniden açma
- **Tema Özelleştirme**: Kişisel tema seçimi

## 🏗️ Mimari / Architecture

```
excel-ai-assistant/
├── app.js              # Ana uygulama mantığı
├── ai_client.js        # AI API istemcisi
├── index.html          # Kullanıcı arayüzü
├── styles.css          # Stil tanımları
├── manifest.xml        # Office Add-in manifest
├── package.json        # Proje bağımlılıkları
└── README.md           # Proje dokümantasyonu
```

## 🤝 Katkıda Bulunma / Contributing

Bu proje açık kaynak kodludur ve katkılarınızı bekliyoruz!

1. Fork yapın
2. Feature branch oluşturun (`git checkout -b feature/AmazingFeature`)
3. Değişikliklerinizi commit edin (`git commit -m 'Add some AmazingFeature'`)
4. Branch'inizi push edin (`git push origin feature/AmazingFeature`)
5. Pull Request oluşturun

## 📄 Lisans / License

Bu proje MIT lisansı altında lisanslanmıştır. Detaylar için [LICENSE](LICENSE) dosyasına bakın.

## 🙏 Teşekkürler / Acknowledgments

- Microsoft Office Add-ins ekibi
- LM Studio geliştiricileri
- Açık kaynak topluluğu
- Tüm katkıda bulunanlar

## 📞 İletişim / Contact

- **GitHub Issues**: [Proje Issues](https://github.com/[username]/excel-ai-assistant/issues)
- **Discussions**: [GitHub Discussions](https://github.com/[username]/excel-ai-assistant/discussions)

---

**⭐ Bu projeyi beğendiyseniz yıldız vermeyi unutmayın!**

**⭐ If you like this project, don't forget to give it a star!**
