#!/usr/bin/env node

const fs = require('fs');
const { execSync } = require('child_process');
const os = require('os');

// IP adresini otomatik olarak tespit et
function getLocalIP() {
    try {
        // Linux sistemlerde hostname -I komutu ile IP adresini al
        const ip = execSync('hostname -I', { encoding: 'utf8' }).trim().split(' ')[0];
        return ip;
    } catch (error) {
        // Fallback: OS network interface'lerini kontrol et
        const interfaces = os.networkInterfaces();
        for (const name of Object.keys(interfaces)) {
            for (const interface of interfaces[name]) {
                if (interface.family === 'IPv4' && !interface.internal) {
                    return interface.address;
                }
            }
        }
        return '127.0.0.1'; // Default fallback
    }
}

// Dosyadaki IP adresini güncelle
function updateIPInFile(filePath, oldIP, newIP) {
    try {
        let content = fs.readFileSync(filePath, 'utf8');
        let updated = false;
        
        // Farklı port kombinasyonlarını kontrol et
        const patterns = [
            new RegExp(`http://${oldIP}:\\d+`, 'g'),
            new RegExp(`http://\\d+\\.\\d+\\.\\d+\\.\\d+:\\d+`, 'g')
        ];
        
        for (const pattern of patterns) {
            if (pattern.test(content)) {
                content = content.replace(pattern, (match) => {
                    const port = match.split(':')[2];
                    return `http://${newIP}:${port}`;
                });
                updated = true;
            }
        }
        
        if (updated) {
            fs.writeFileSync(filePath, content, 'utf8');
            console.log(`✅ ${filePath} güncellendi`);
            return true;
        } else {
            console.log(`ℹ️  ${filePath} dosyasında IP adresi bulunamadı`);
            return false;
        }
    } catch (error) {
        console.error(`❌ ${filePath} güncellenirken hata:`, error.message);
        return false;
    }
}

// Ana fonksiyon
function main() {
    console.log('🌐 IP Adresi Otomatik Güncelleme Başlatılıyor...\n');
    
    const newIP = getLocalIP();
    console.log(`📍 Tespit edilen IP adresi: ${newIP}`);
    
    // Güncellenecek dosyalar
    const files = [
        'ai_client.js',
        'README.md'
    ];
    
    let updatedCount = 0;
    
    for (const file of files) {
        if (fs.existsSync(file)) {
            // Mevcut IP adresini bul (herhangi bir IP pattern'i)
            const content = fs.readFileSync(file, 'utf8');
            const ipMatch = content.match(/http:\/\/(\d+\.\d+\.\d+\.\d+):(\d+)/);
            
            if (ipMatch) {
                const oldIP = ipMatch[1];
                const port = ipMatch[2];
                console.log(`🔄 ${file} dosyasında ${oldIP}:${port} bulundu`);
                
                if (updateIPInFile(file, oldIP, newIP)) {
                    updatedCount++;
                }
            } else {
                console.log(`⚠️  ${file} dosyasında IP adresi bulunamadı`);
            }
        } else {
            console.log(`❌ ${file} dosyası bulunamadı`);
        }
    }
    
    console.log(`\n🎯 Güncelleme tamamlandı! ${updatedCount} dosya güncellendi.`);
    console.log(`📡 Yeni IP adresi: ${newIP}`);
    console.log(`🔗 Ollama bağlantı URL'i: http://${newIP}:11434`);
    console.log(`🌐 Web sunucu URL'i: http://${newIP}:3000`);
    
    // package.json'a yeni script ekle
    try {
        const packagePath = 'package.json';
        if (fs.existsSync(packagePath)) {
            const package = JSON.parse(fs.readFileSync(packagePath, 'utf8'));
            
            if (!package.scripts['update-ip']) {
                package.scripts['update-ip'] = 'node update_ip.js';
                fs.writeFileSync(packagePath, JSON.stringify(package, null, 2));
                console.log(`\n📦 package.json'a 'update-ip' script'i eklendi`);
                console.log(`💡 Gelecekte 'npm run update-ip' komutu ile IP adresini güncelleyebilirsiniz`);
            }
        }
    } catch (error) {
        console.log(`\n⚠️  package.json güncellenemedi: ${error.message}`);
    }
}

// Script çalıştırıldığında ana fonksiyonu çalıştır
if (require.main === module) {
    main();
}

module.exports = { getLocalIP, updateIPInFile };
