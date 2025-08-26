#!/bin/bash

echo "🚀 Excel AI Assistant Başlatılıyor..."
echo "======================================"

# IP adresini otomatik güncelle
echo "🌐 IP adresi güncelleniyor..."
node update_ip.js

echo ""
echo "📦 Bağımlılıklar kontrol ediliyor..."
npm install

echo ""
echo "🌐 Web sunucu başlatılıyor..."
echo "📍 URL: http://localhost:3000"
echo "🌐 Network URL: http://$(hostname -I | awk '{print $1}'):3000"
echo ""
echo "⏹️  Sunucuyu durdurmak için Ctrl+C tuşlayın"
echo ""

# Web sunucusunu başlat
npm start
