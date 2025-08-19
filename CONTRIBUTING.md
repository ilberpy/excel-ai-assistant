# Katkıda Bulunma Rehberi / Contributing Guide

Bu projeye katkıda bulunmak istediğiniz için teşekkür ederiz! Bu rehber, projeye nasıl katkıda bulunacağınızı açıklar.

Thank you for wanting to contribute to this project! This guide explains how you can contribute.

## 🚀 Başlangıç / Getting Started

### Gereksinimler / Requirements
- Node.js 18 veya üzeri
- Git
- LM Studio (test için)
- Microsoft Excel

### Kurulum / Setup
```bash
# Repo'yu fork edin ve clone yapın
git clone https://github.com/ilberpy/excel-ai-assistant.git
cd excel-ai-assistant

# Bağımlılıkları yükleyin
npm install

# Geliştirme sunucusunu başlatın
npm run dev
```

## 🔧 Geliştirme / Development

### Kod Standartları / Code Standards
- **JavaScript**: ES6+ syntax kullanın
- **CSS**: BEM metodolojisi takip edin
- **HTML**: Semantic HTML kullanın
- **Türkçe**: Türkçe karakterler için UTF-8 encoding kullanın

### Dosya Yapısı / File Structure
```
excel-ai-assistant/
├── .github/           # GitHub Actions ve templates
├── assets/            # Resimler ve statik dosyalar
├── src/               # Kaynak kodlar (gelecekte)
├── app.js             # Ana uygulama mantığı
├── ai_client.js       # AI API istemcisi
├── index.html         # Ana HTML dosyası
├── styles.css         # CSS stilleri
├── manifest.xml       # Office Add-in manifest
└── package.json       # Proje konfigürasyonu
```

### Commit Mesajları / Commit Messages
```
feat: yeni özellik ekleme
fix: hata düzeltme
docs: dokümantasyon güncelleme
style: kod formatı düzenleme
refactor: kod yeniden düzenleme
test: test ekleme veya düzenleme
chore: bakım işlemleri
```

## 🐛 Hata Bildirimi / Bug Reporting

### Hata Raporu Şablonu / Bug Report Template
```markdown
## Hata Açıklaması / Bug Description
Kısa ve net bir açıklama yazın.

## Yeniden Üretme / Steps to Reproduce
1. Şu adımları takip edin...
2. Şu hatayı alırsınız...

## Beklenen Davranış / Expected Behavior
Ne olması gerekiyordu?

## Gerçek Davranış / Actual Behavior
Ne oldu?

## Ek Bilgiler / Additional Information
- Excel versiyonu:
- İşletim sistemi:
- Tarayıcı:
- Ekran görüntüleri:
```

## 💡 Özellik Önerisi / Feature Request

### Özellik Önerisi Şablonu / Feature Request Template
```markdown
## Özellik Açıklaması / Feature Description
Bu özellik ne yapacak?

## Problem / Problem
Hangi problemi çözecek?

## Çözüm / Solution
Nasıl çözülecek?

## Alternatifler / Alternatives
Başka hangi çözümler düşünüldü?

## Ek Bilgiler / Additional Information
Ekran görüntüleri, mockup'lar, vb.
```

## 🔄 Pull Request Süreci / Pull Request Process

### PR Oluşturma / Creating a PR
1. **Fork yapın** ve feature branch oluşturun
2. **Kodunuzu yazın** ve test edin
3. **Commit yapın** açıklayıcı mesajlarla
4. **Push yapın** branch'inizi
5. **Pull Request oluşturun** detaylı açıklamayla

### PR Şablonu / PR Template
```markdown
## Değişiklik Açıklaması / Change Description
Bu PR ne yapıyor?

## Test / Testing
Nasıl test edildi?

## Ekran Görüntüleri / Screenshots
Görsel değişiklikler varsa ekleyin.

## Checklist / Checklist
- [ ] Kod standartlarına uygun
- [ ] Testler geçiyor
- [ ] Dokümantasyon güncellendi
- [ ] Commit mesajları uygun
```

## 🧪 Test / Testing

### Test Çalıştırma / Running Tests
```bash
# Tüm testleri çalıştır
npm test

# Belirli test dosyasını çalıştır
npm test -- --grep "test name"

# Coverage raporu
npm run test:coverage
```

### Test Yazma / Writing Tests
```javascript
// Test örneği
describe('AI Client', () => {
  it('should connect to LM Studio', async () => {
    const result = await aiClient.testConnection();
    expect(result).toBe(true);
  });
});
```

## 📚 Dokümantasyon / Documentation

### Dokümantasyon Güncelleme / Updating Documentation
- README.md dosyasını güncelleyin
- Yeni özellikler için örnekler ekleyin
- API değişikliklerini belgelendirin
- Türkçe ve İngilizce açıklamalar ekleyin

### Çeviri / Translation
- Türkçe metinleri İngilizce'ye çevirin
- İngilizce metinleri Türkçe'ye çevirin
- Dil tutarlılığını koruyun

## 🎨 UI/UX Geliştirme / UI/UX Development

### Tasarım Prensipleri / Design Principles
- **Kullanıcı Dostu**: Basit ve sezgisel arayüz
- **Responsive**: Tüm ekran boyutlarında çalışma
- **Accessibility**: Erişilebilirlik standartlarına uygun
- **Dark Theme**: Modern koyu tema desteği

### CSS Kuralları / CSS Guidelines
```css
/* BEM metodolojisi */
.chat-message { }
.chat-message--user { }
.chat-message__content { }

/* CSS değişkenleri kullanın */
:root {
  --primary-color: #0078d4;
  --secondary-color: #106ebe;
}
```

## 🔒 Güvenlik / Security

### Güvenlik Açığı Bildirimi / Security Vulnerability Report
Güvenlik açığı bulduysanız:
1. **Özel olarak bildirin**: security@example.com
2. **Hemen yayınlamayın**
3. **Detaylı bilgi verin**
4. **Proof of concept ekleyin**

## 📞 İletişim / Communication

### Topluluk / Community
- **GitHub Issues**: Hata bildirimi ve özellik önerisi
- **GitHub Discussions**: Genel tartışmalar
- **Discord**: Canlı sohbet (gelecekte)

### Geliştirici Toplantıları / Developer Meetings
- **Haftalık sync**: Her Cuma 15:00 (TR)
- **Aylık review**: Her ayın ilk Pazartesi
- **Quarterly planning**: Her 3 ayda bir

## 🏆 Katkıda Bulunanlar / Contributors

Katkıda bulunanlar [CONTRIBUTORS.md](CONTRIBUTORS.md) dosyasında listelenir.

## 📄 Lisans / License

Bu proje MIT lisansı altında lisanslanmıştır. Katkıda bulunarak bu lisansı kabul etmiş olursunuz.

---

**Teşekkürler! / Thank you!**

Bu projeye katkıda bulunduğunuz için teşekkür ederiz. Birlikte daha iyi bir Excel AI Assistant yapabiliriz!
