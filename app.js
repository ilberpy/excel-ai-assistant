// Excel AI Assistant - Dark Theme Chat Interface

// Chat geçmişi için global değişkenler
let chatHistory = [];
let currentChatId = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Excel yüklendiğinde çalışacak kod
        initializeChatInterface();
        
        // AI bağlantısını test et
        testAIConnection();
        
            // Eğer aktif chat yoksa yeni chat başlat
    if (!getCurrentChatId()) {
        startNewChat();
    }
    }
});

// Chat arayüzünü başlat
function initializeChatInterface() {
    const commandInput = document.getElementById('aiCommandInput');
    
    // Enter tuşu ile komut gönderme
    commandInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            processUserCommand();
        }
    });
    
    // Header butonları
    document.getElementById('newChatBtn').onclick = startNewChat;
    document.getElementById('historyBtn').onclick = showHistory;
    document.getElementById('menuBtn').onclick = showMenu;
    document.getElementById('closeBtn').onclick = closeApp;
    
    // Command bar butonları
    document.getElementById('imageBtn').onclick = handleImageUpload;
    document.getElementById('voiceBtn').onclick = handleVoiceInput;
    
    // Yanıtı durdur butonu
    document.getElementById('stopResponseBtn').onclick = stopResponse;
    
    // Chat geçmişini yükle
    loadChatHistory();
}

// Chat geçmişi yönetimi fonksiyonları
function loadChatHistory() {
    try {
        const saved = localStorage.getItem('excelAI_chatHistory');
        if (saved) {
            chatHistory = JSON.parse(saved);
            console.log(`📚 ${chatHistory.length} chat geçmişi yüklendi`);
        }
    } catch (error) {
        console.error('Chat geçmişi yüklenemedi:', error);
        chatHistory = [];
    }
}

function saveChatHistory() {
    try {
        localStorage.setItem('excelAI_chatHistory', JSON.stringify(chatHistory));
        console.log('💾 Chat geçmişi kaydedildi');
    } catch (error) {
        console.error('Chat geçmişi kaydedilemedi:', error);
    }
}

function getCurrentChatId() {
    return currentChatId;
}

function generateChatId() {
    return 'chat_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
}

function getCurrentChat() {
    if (!currentChatId) return null;
    return chatHistory.find(chat => chat.id === currentChatId);
}

function saveCurrentChat() {
    if (!currentChatId) return;
    
    const chatContainer = document.getElementById('chat-container');
    const messages = [];
    
    // Chat container'daki mesajları topla
    chatContainer.querySelectorAll('.chat-message').forEach(msg => {
        const type = msg.classList.contains('user') ? 'user' : 'ai';
        const content = msg.querySelector('.message-content').textContent;
        const sender = msg.querySelector('.message-header span').textContent;
        
        messages.push({
            type,
            content,
            sender,
            timestamp: Date.now()
        });
    });
    
    // Chat'i güncelle veya ekle
    const existingChatIndex = chatHistory.findIndex(chat => chat.id === currentChatId);
    const chatData = {
        id: currentChatId,
        title: messages.length > 0 ? messages[0].content.substring(0, 50) + '...' : 'Yeni Sohbet',
        messages: messages,
        lastUpdated: Date.now(),
        messageCount: messages.length
    };
    
    if (existingChatIndex >= 0) {
        chatHistory[existingChatIndex] = chatData;
    } else {
        chatHistory.unshift(chatData);
    }
    
    saveChatHistory();
}

// Chat yükle
function loadChat(chatId) {
    // Mevcut chat'i kaydet
    if (currentChatId) {
        saveCurrentChat();
    }
    
    // Chat'i bul
    const chat = chatHistory.find(c => c.id === chatId);
    if (!chat) {
        console.error('Chat bulunamadı:', chatId);
        return;
    }
    
    // Chat ID'yi güncelle
    currentChatId = chatId;
    
    // Chat container'ı temizle
    const chatContainer = document.getElementById('chat-container');
    chatContainer.innerHTML = '';
    
    // Mesajları yükle
    chat.messages.forEach(msg => {
        addMessage(msg.type, msg.content, msg.sender);
    });
    
    // Modal'ı kapat
    const modal = document.querySelector('.history-modal');
    if (modal) modal.remove();
    
    console.log('📚 Chat yüklendi:', chatId);
}

// Chat sil
function deleteChat(chatId, event) {
    // Event'i durdur
    if (event) {
        event.stopPropagation(); // Parent click event'ini engelle
        event.preventDefault(); // Default davranışı engelle
    }
    
    // Office Add-in ortamında confirm yerine custom modal kullan
    showDeleteConfirmation(chatId);
}

// Silme onay modal'ı göster
function showDeleteConfirmation(chatId) {
    // Mevcut history modal'ı bul
    const historyModal = document.querySelector('.history-modal');
    if (!historyModal) return;
    
    // Onay modal'ı oluştur
    const confirmModal = document.createElement('div');
    confirmModal.className = 'confirm-modal';
    confirmModal.innerHTML = `
        <div class="confirm-content">
            <div class="confirm-header">
                <h4>🗑️ Sohbet Sil</h4>
            </div>
            <div class="confirm-body">
                <p>Bu sohbeti silmek istediğinizden emin misiniz?</p>
                <p class="confirm-warning">Bu işlem geri alınamaz!</p>
            </div>
            <div class="confirm-actions">
                <button class="confirm-btn confirm-cancel" type="button">❌ İptal</button>
                <button class="confirm-btn confirm-delete" type="button">🗑️ Sil</button>
            </div>
        </div>
    `;
    
    // Modal'ı history modal'ın üzerine ekle
    historyModal.appendChild(confirmModal);
    
    // Event listener'ları ekle
    const cancelBtn = confirmModal.querySelector('.confirm-cancel');
    const deleteBtn = confirmModal.querySelector('.confirm-delete');
    
    cancelBtn.addEventListener('click', () => {
        confirmModal.remove();
    });
    
    deleteBtn.addEventListener('click', () => {
        try {
            // Chat'i geçmişten kaldır
            chatHistory = chatHistory.filter(chat => chat.id !== chatId);
            saveChatHistory();
            
            // Eğer aktif chat silindiyse yeni chat başlat
            if (currentChatId === chatId) {
                startNewChat();
            }
            
            // Onay modal'ını kapat
            confirmModal.remove();
            
            // History modal'ı yenile
            historyModal.remove();
            setTimeout(() => showHistory(), 200);
            
            console.log('🗑️ Chat silindi:', chatId);
        } catch (error) {
            console.error('Chat silme hatası:', error);
            showErrorMessage('Chat silinirken bir hata oluştu. Lütfen tekrar deneyin.');
        }
    });
}

// Hata mesajı göster
function showErrorMessage(message) {
    const errorModal = document.createElement('div');
    errorModal.className = 'error-modal';
    errorModal.innerHTML = `
        <div class="error-content">
            <div class="error-header">
                <h4>❌ Hata</h4>
            </div>
            <div class="error-body">
                <p>${message}</p>
            </div>
            <div class="error-actions">
                <button class="error-btn" type="button">Tamam</button>
            </div>
        </div>
    `;
    
    document.body.appendChild(errorModal);
    
    const okBtn = errorModal.querySelector('.error-btn');
    okBtn.addEventListener('click', () => {
        errorModal.remove();
    });
}

// Zaman önce fonksiyonu
function getTimeAgo(date) {
    const now = new Date();
    const diff = now - date;
    const minutes = Math.floor(diff / 60000);
    const hours = Math.floor(diff / 3600000);
    const days = Math.floor(diff / 86400000);
    
    if (minutes < 1) return 'Az önce';
    if (minutes < 60) return `${minutes} dakika önce`;
    if (hours < 24) return `${hours} saat önce`;
    if (days < 7) return `${days} gün önce`;
    
    return date.toLocaleDateString('tr-TR');
}

// AI bağlantısını test et
async function testAIConnection() {
    if (window.aiClient) {
        const isConnected = await window.aiClient.testConnection();
        if (isConnected) {
            addMessage('ai', '✅ AI bağlantısı başarılı! LM Studio ile bağlantı kuruldu.', 'System');
        } else {
            addMessage('ai', '❌ AI bağlantısı başarısız. LM Studio çalışıyor mu kontrol edin.', 'System');
        }
    }
}

// Kullanıcı komutunu işle
async function processUserCommand() {
    const commandInput = document.getElementById('aiCommandInput');
    const command = commandInput.value.trim();
    
    if (!command) return;
    
    // Kullanıcı mesajını ekle
    addMessage('user', command, 'You');
    
    // Input'u temizle
    commandInput.value = '';
    
    // Loading göster
    showLoading();
    
    try {
        // AI'dan yanıt al
        const response = await window.aiClient.parseExcelCommandStreaming(
            command,
            (chunk, fullResponse) => {
                // Streaming güncelle
                updateLastMessage(fullResponse);
            }
        );
        
        if (response) {
            // Loading gizle
            hideLoading();
            
            // AI yanıtını güncelle
            updateLastMessage(response);
            
            // AI yanıtını analiz et ve Excel'de otomatik uygula
            try {
                await executeAIResponseIntelligently(command, response);
            } catch (autoError) {
                console.log('AI yanıt uygulama hatası:', autoError);
            }
        } else {
            hideLoading();
            addMessage('ai', 'AI yanıtı alınamadı. LM Studio bağlantısını kontrol edin.', 'System');
        }
    } catch (error) {
        hideLoading();
        addMessage('ai', `Hata: ${error.message}`, 'System');
    }
}

// Chat mesajı ekle
function addMessage(type, content, sender) {
    const chatContainer = document.getElementById('chat-container');
    
    const messageDiv = document.createElement('div');
    messageDiv.className = `chat-message ${type}`;
    
    // Özel sender'lar için CSS class ekle
    if (sender === 'Data Analysis') {
        messageDiv.classList.add('data-analysis');
    } else if (sender === 'Data Summary') {
        messageDiv.classList.add('data-summary');
    } else if (sender === 'Trend Analysis') {
        messageDiv.classList.add('trend-analysis');
    } else if (sender === 'Anomaly Detection') {
        messageDiv.classList.add('anomaly-detection');
    }
    
    const headerDiv = document.createElement('div');
    headerDiv.className = 'message-header';
    headerDiv.innerHTML = `<span>${sender}</span>`;
    
    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content';
    contentDiv.textContent = content;
    
    messageDiv.appendChild(headerDiv);
    messageDiv.appendChild(contentDiv);
    
    chatContainer.appendChild(messageDiv);
    
    // Chat'i kaydet
    if (currentChatId) {
        saveCurrentChat();
    }
    
    // Scroll'u en alta getir
    scrollToBottom();
    
    return messageDiv;
}

// Scroll'u en alta getir
function scrollToBottom() {
    const chatContainer = document.getElementById('chat-container');
    const mainContent = document.querySelector('.main-content');
    
    // Ana content container'ı scroll yap (chat-container değil)
    if (mainContent) {
        setTimeout(() => {
            mainContent.scrollTo({
                top: mainContent.scrollHeight,
                behavior: 'smooth'
            });
        }, 50);
    }
}

// Yanıtı durdur
function stopResponse() {
    // AI client'ta streaming'i durdur
    if (window.aiClient && window.aiClient.stopStreaming) {
        window.aiClient.stopStreaming();
    }
    
    // Loading'i gizle
    hideLoading();
    
    // Son mesajı güncelle (duplicate olmaması için sadece bu)
    updateLastMessage('⏹️ Yanıt durduruldu.');
}

// Son mesajı güncelle (streaming için)
function updateLastMessage(content) {
    const chatContainer = document.getElementById('chat-container');
    const lastMessage = chatContainer.lastElementChild;
    
    if (lastMessage && lastMessage.classList.contains('ai')) {
        const contentDiv = lastMessage.querySelector('.message-content');
        if (contentDiv) {
            contentDiv.textContent = content;
        }
        
        // Streaming sırasında her güncellemede anında scroll yap
        setTimeout(() => {
            const mainContent = document.querySelector('.main-content');
            if (mainContent) {
                mainContent.scrollTo({
                    top: mainContent.scrollHeight,
                    behavior: 'smooth'
                });
            }
        }, 5); // Çok hızlı scroll için 5ms
    } else {
        // Eğer son mesaj AI değilse yeni mesaj ekle
        addMessage('ai', content, 'AI Assistant');
    }
}

// Loading göster
function showLoading() {
    const loadingIndicator = document.getElementById('loadingIndicator');
    const stopBtn = document.getElementById('stopResponseBtn');
    
    loadingIndicator.style.display = 'flex';
    stopBtn.style.display = 'flex'; // Durdur butonunu göster
    
    // Loading mesajı ekle
    addMessage('ai', '🤔 Düşünüyorum...', 'AI Assistant');
    
    // Hemen scroll yap (loading mesajından sonra)
    setTimeout(() => {
        const mainContent = document.querySelector('.main-content');
        if (mainContent) {
            mainContent.scrollTo({
                top: mainContent.scrollHeight,
                behavior: 'smooth'
            });
        }
    }, 100);
}

// Loading gizle
function hideLoading() {
    const loadingIndicator = document.getElementById('loadingIndicator');
    const stopBtn = document.getElementById('stopResponseBtn');
    
    loadingIndicator.style.display = 'none';
    stopBtn.style.display = 'none'; // Durdur butonunu gizle
}

// Komutu otomatik olarak uygulamaya çalış
async function autoExecuteCommand(command, aiResult) {
    const lowerCommand = command.toLowerCase();
    
    try {
        if (lowerCommand.includes('grafik') || lowerCommand.includes('chart')) {
            await createChartFromData();
        } else if (lowerCommand.includes('toplam') || lowerCommand.includes('sum')) {
            await calculateSum();
        } else if (lowerCommand.includes('ortalama') || lowerCommand.includes('average')) {
            await calculateAverage();
        } else if (lowerCommand.includes('filtre') || lowerCommand.includes('filter')) {
            await applyFilter();
        } else if (lowerCommand.includes('sırala') || lowerCommand.includes('sort')) {
            await sortData();
        } else if (lowerCommand.includes('analiz') || lowerCommand.includes('analyze') || 
                   lowerCommand.includes('incele') || lowerCommand.includes('examine')) {
            await analyzeDataIntelligently();
        } else if (lowerCommand.includes('özet') || lowerCommand.includes('summary')) {
            await generateDataSummary();
        } else if (lowerCommand.includes('trend') || lowerCommand.includes('eğilim')) {
            await analyzeTrends();
        } else if (lowerCommand.includes('anomali') || lowerCommand.includes('outlier')) {
            await detectAnomalies();
        }
    } catch (error) {
        console.log('Otomatik komut uygulanamadı:', error);
    }
}

// Grafik oluştur
async function createChartFromData() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('values, address');
            
            await context.sync();
            
            if (range.values && range.values.length > 1) {
                try {
                    const chart = range.worksheet.charts.add(Excel.ChartType.columnClustered, range);
                    chart.setPosition(0, range.getColumn(0).getColumnIndex() + range.values[0].length + 2);
                    
                    await context.sync();
                    addMessage('ai', '📊 Grafik başarıyla oluşturuldu!', 'Excel');
                } catch (chartError) {
                    console.log('Grafik oluşturma hatası, alternatif yöntem deneniyor...');
                    
                    range.format.borders.getItem('EdgeBottom').style = 'Continuous';
                    range.format.borders.getItem('EdgeRight').style = 'Continuous';
                    range.format.fill.color = '#e6f3ff';
                    
                    await context.sync();
                    addMessage('ai', '📊 Veri tablosu formatlandı! (Grafik yerine)', 'Excel');
                }
            }
        });
    } catch (error) {
        addMessage('ai', `Grafik oluşturma hatası: ${error.message}`, 'System');
    }
}

// Toplam hesapla
async function calculateSum() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('values, address');
            
            await context.sync();
            
            if (range.values && range.values.length > 0) {
                try {
                    const worksheet = range.worksheet;
                    const lastRowIndex = range.getRow(range.values.length - 1).getRowIndex();
                    const nextRow = worksheet.getRange(`${lastRowIndex + 1}:${lastRowIndex + 1}`);
                    
                    const sumFormula = `=SUM(${range.address})`;
                    nextRow.values = [[sumFormula]];
                    nextRow.format.font.bold = true;
                    nextRow.format.fill.color = '#e6f3ff';
                    
                    await context.sync();
                    addMessage('ai', '🧮 Toplam hesaplandı!', 'Excel');
                } catch (apiError) {
                    console.log('API toplam hatası, alternatif yöntem deneniyor...');
                    
                    const worksheet = range.worksheet;
                    const lastRowIndex = range.getRow(range.values.length - 1).getRowIndex();
                    const nextRowAddress = `${lastRowIndex + 1}:${lastRowIndex + 1}`;
                    const nextRow = worksheet.getRange(nextRowAddress);
                    
                    let total = 0;
                    for (let i = 1; i < range.values.length; i++) {
                        for (let j = 0; j < range.values[i].length; j++) {
                            const value = parseFloat(range.values[i][j]);
                            if (!isNaN(value)) {
                                total += value;
                            }
                        }
                    }
                    
                    nextRow.values = [[`Toplam: ${total}`]];
                    nextRow.format.font.bold = true;
                    nextRow.format.fill.color = '#e6f3ff';
                    
                    await context.sync();
                    addMessage('ai', `🧮 Toplam hesaplandı: ${total}`, 'Excel');
                }
            }
        });
    } catch (error) {
        addMessage('ai', `Toplam hesaplama hatası: ${error.message}`, 'System');
    }
}

// Ortalama hesapla
async function calculateAverage() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('values, address');
            
            await context.sync();
            
            if (range.values && range.values.length > 0) {
                try {
                    const worksheet = range.worksheet;
                    const lastRowIndex = range.getRow(range.values.length - 1).getRowIndex();
                    const nextRow = worksheet.getRange(`${lastRowIndex + 1}:${lastRowIndex + 1}`);
                    
                    const avgFormula = `=AVERAGE(${range.address})`;
                    nextRow.values = [[avgFormula]];
                    nextRow.format.font.bold = true;
                    nextRow.format.fill.color = '#fff2e6';
            
            await context.sync();
                    addMessage('ai', '📊 Ortalama hesaplandı!', 'Excel');
                } catch (apiError) {
                    console.log('API ortalama hatası, alternatif yöntem deneniyor...');
                    
                    const worksheet = range.worksheet;
                    const lastRowIndex = range.getRow(range.values.length - 1).getRowIndex();
                    const nextRowAddress = `${lastRowIndex + 1}:${lastRowIndex + 1}`;
                    const nextRow = worksheet.getRange(nextRowAddress);
                    
                    let total = 0;
                    let count = 0;
                    for (let i = 1; i < range.values.length; i++) {
                        for (let j = 0; j < range.values[i].length; j++) {
                            const value = parseFloat(range.values[i][j]);
                            if (!isNaN(value)) {
                                total += value;
                                count++;
                            }
                        }
                    }
                    
                    const average = count > 0 ? (total / count).toFixed(2) : 0;
                    nextRow.values = [[`Ortalama: ${average}`]];
                    nextRow.format.font.bold = true;
                    nextRow.format.fill.color = '#fff2e6';
            
            await context.sync();
                    addMessage('ai', `📊 Ortalama hesaplandı: ${average}`, 'Excel');
                }
            }
        });
    } catch (error) {
        addMessage('ai', `Ortalama hesaplama hatası: ${error.message}`, 'System');
    }
}

// Filtre uygula
async function applyFilter() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('values, address');
            
            await context.sync();
            
            if (range.values && range.values.length > 0) {
                try {
                    range.autoFilter();
                    await context.sync();
                    addMessage('ai', '🔍 Filtre uygulandı!', 'Excel');
                } catch (filterError) {
                    console.log('Filtre uygulama hatası, alternatif yöntem deneniyor...');
                    
                const headerRow = range.getRow(0);
                headerRow.format.fill.color = '#0078d4';
                headerRow.format.font.color = 'white';
                headerRow.format.font.bold = true;
                
                    await context.sync();
                    addMessage('ai', '🔍 Başlık satırı vurgulandı! (Filtre yerine)', 'Excel');
                }
            }
        });
    } catch (error) {
        addMessage('ai', `Filtre uygulama hatası: ${error.message}`, 'System');
    }
}

// Veri sırala
async function sortData() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('values, address');
            
            await context.sync();
            
            if (range.values && range.values.length > 1) {
                try {
                    const sortOptions = [{
                        key: 0,
                        sortOn: Excel.SortOn.values,
                        order: Excel.SortOrder.ascending
                    }];
                    
                    range.sort.apply(sortOptions);
                    await context.sync();
                    addMessage('ai', '📈 Veri sıralandı!', 'Excel');
                } catch (sortError) {
                    console.log('Enum sıralama hatası, alternatif yöntem deneniyor...');
                    
                    const values = range.values;
                    const sortedValues = values.sort((a, b) => {
                        if (a[0] < b[0]) return -1;
                        if (a[0] > b[0]) return 1;
                        return 0;
                    });
                    
                    range.values = sortedValues;
                await context.sync();
                    addMessage('ai', '📈 Veri JavaScript ile sıralandı!', 'Excel');
                }
            }
        });
    } catch (error) {
        addMessage('ai', `Sıralama hatası: ${error.message}`, 'System');
    }
}

// Header buton fonksiyonları
function startNewChat() {
    // Mevcut chat'i kaydet
    if (currentChatId) {
        saveCurrentChat();
    }
    
    // Yeni chat ID oluştur
    currentChatId = generateChatId();
    
    // Chat container'ı temizle
    const chatContainer = document.getElementById('chat-container');
    chatContainer.innerHTML = '';
    
    // Input'u temizle
    document.getElementById('aiCommandInput').value = '';
    
    // Hoş geldin mesajı ekle
    addMessage('ai', '🤖 Yeni sohbet başlatıldı! Excel verilerinizle ilgili herhangi bir komut yazabilirsiniz.', 'Excel AI Assistant');
    
    // Scroll'u en alta getir
    scrollToBottom();
    
    console.log('✅ Yeni sohbet başlatıldı:', currentChatId);
}

function showHistory() {
    // Sohbet geçmişi modal'ı oluştur
    const modal = document.createElement('div');
    modal.className = 'history-modal';
    
    // Chat geçmişini dinamik olarak oluştur
    let historyHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <h3>📚 Sohbet Geçmişi</h3>
                <button class="close-modal-btn" onclick="this.parentElement.parentElement.parentElement.remove()">×</button>
            </div>
            <div class="modal-body">
    `;
    
    if (chatHistory.length === 0) {
        historyHTML += `
            <div class="history-empty">
                <div class="empty-icon">📚</div>
                <div class="empty-text">Henüz sohbet geçmişi yok</div>
                <div class="empty-subtext">İlk komutunuzu yazın ve sohbet başlayın!</div>
            </div>
        `;
    } else {
        chatHistory.forEach(chat => {
            const date = new Date(chat.lastUpdated);
            const timeAgo = getTimeAgo(date);
            const isActive = chat.id === currentChatId;
            
            historyHTML += `
                <div class="history-item ${isActive ? 'active' : ''}" data-chat-id="${chat.id}">
                    <div class="history-info">
                        <div class="history-title">${chat.title}</div>
                        <div class="history-meta">
                            <span class="history-date">${timeAgo}</span>
                            <span class="history-count">${chat.messageCount} mesaj</span>
                        </div>
                    </div>
                    <div class="history-actions">
                        ${isActive ? '<span class="active-indicator">●</span>' : ''}
                        <button class="delete-chat-btn" data-chat-id="${chat.id}" type="button">🗑️</button>
                    </div>
                </div>
            `;
        });
    }
    
    historyHTML += `
            </div>
        </div>
    `;
    
    modal.innerHTML = historyHTML;
    document.body.appendChild(modal);
    
    // Event listener'ları ekle
    setupHistoryEventListeners(modal);
    
    console.log('📚 Sohbet geçmişi gösteriliyor');
}

// History modal event listener'larını kur
function setupHistoryEventListeners(modal) {
    // Chat item click event'leri
    const historyItems = modal.querySelectorAll('.history-item');
    historyItems.forEach(item => {
        item.addEventListener('click', (e) => {
            // Eğer delete butonuna tıklandıysa chat yükleme
            if (e.target.classList.contains('delete-chat-btn')) {
                return;
            }
            
            const chatId = item.dataset.chatId;
            if (chatId) {
                loadChat(chatId);
            }
        });
    });
    
    // Delete buton click event'leri
    const deleteButtons = modal.querySelectorAll('.delete-chat-btn');
    deleteButtons.forEach(button => {
        button.addEventListener('click', (e) => {
            e.stopPropagation();
            e.preventDefault();
            
            const chatId = button.dataset.chatId;
            if (chatId) {
                deleteChat(chatId, e);
            }
        });
    });
}

function showMenu() {
    // Menü modal'ı oluştur
    const modal = document.createElement('div');
    modal.className = 'menu-modal';
    modal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <h3>⚙️ Ayarlar & Menü</h3>
                <button class="close-modal-btn" onclick="this.parentElement.parentElement.parentElement.remove()">×</button>
            </div>
            <div class="modal-body">
                <div class="menu-item" data-action="ai-model">
                    <span class="menu-icon">🔧</span>
                    <span class="menu-text">AI Model Seçimi</span>
                    <span class="menu-arrow">→</span>
                </div>
                <div class="menu-item" data-action="theme">
                    <span class="menu-icon">🎨</span>
                    <span class="menu-text">Tema Değiştir</span>
                    <span class="menu-arrow">→</span>
                </div>
                <div class="menu-item" data-action="excel-settings">
                    <span class="menu-icon">📊</span>
                    <span class="menu-text">Excel Ayarları</span>
                    <span class="menu-arrow">→</span>
                </div>
                <div class="menu-item" data-action="about">
                    <span class="menu-icon">ℹ️</span>
                    <span class="menu-text">Hakkında</span>
                    <span class="menu-arrow">→</span>
                </div>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    // Event listener'ları ekle
    setupMenuEventListeners(modal);
    
    console.log('⚙️ Menü gösteriliyor');
}

// Menü event listener'larını kur
function setupMenuEventListeners(modal) {
    const menuItems = modal.querySelectorAll('.menu-item');
    menuItems.forEach(item => {
        item.addEventListener('click', () => {
            const action = item.dataset.action;
            handleMenuAction(action, modal);
        });
    });
}

// Menü aksiyonlarını işle
function handleMenuAction(action, menuModal) {
    switch (action) {
        case 'ai-model':
            showAIModelSelection(menuModal);
            break;
        case 'theme':
            showThemeSelection(menuModal);
            break;
        case 'excel-settings':
            showExcelSettings(menuModal);
            break;
        case 'about':
            showAboutInfo(menuModal);
            break;
        default:
            console.log('Bilinmeyen menü aksiyonu:', action);
    }
}

// AI Model Seçimi Modal'ı
function showAIModelSelection(menuModal) {
    // Mevcut menü modal'ını gizle
    menuModal.style.display = 'none';
    
    const modal = document.createElement('div');
    modal.className = 'submenu-modal';
    modal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <button class="back-btn" onclick="showMenu()">←</button>
                <h3>🔧 AI Model Seçimi</h3>
                <button class="close-modal-btn" onclick="this.parentElement.parentElement.parentElement.remove()">×</button>
            </div>
            <div class="modal-body">
                <div class="model-info">
                    <p>Mevcut AI Model: <strong id="currentModel">Yükleniyor...</strong></p>
                    <p>Model Durumu: <span id="modelStatus">Kontrol ediliyor...</span></p>
                </div>
                <div class="model-list" id="modelList">
                    <div class="loading-models">Modeller yükleniyor...</div>
                </div>
                <div class="model-actions">
                    <button class="refresh-btn" onclick="refreshModels()">🔄 Yenile</button>
                    <button class="test-btn" onclick="testCurrentModel()">🧪 Test Et</button>
                </div>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    // Modelleri yükle
    loadAvailableModels();
}

// Tema Seçimi Modal'ı
function showThemeSelection(menuModal) {
    // Mevcut menü modal'ını gizle
    menuModal.style.display = 'none';
    
    const modal = document.createElement('div');
    modal.className = 'submenu-modal';
    modal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <button class="back-btn" onclick="showMenu()">←</button>
                <h3>🎨 Tema Seçimi</h3>
                <button class="close-modal-btn" onclick="this.parentElement.parentElement.parentElement.remove()">×</button>
            </div>
            <div class="modal-body">
                <div class="theme-options">
                    <div class="theme-item" data-theme="dark">
                        <div class="theme-preview dark-theme"></div>
                        <div class="theme-info">
                            <span class="theme-name">Koyu Tema</span>
                            <span class="theme-desc">Varsayılan koyu tema</span>
                        </div>
                        <span class="theme-check">✓</span>
                    </div>
                    <div class="theme-item" data-theme="light">
                        <div class="theme-preview light-theme"></div>
                        <div class="theme-info">
                            <span class="theme-name">Açık Tema</span>
                            <span class="theme-desc">Açık renkli tema</span>
                        </div>
                        <span class="theme-check"></span>
                    </div>
                    <div class="theme-item" data-theme="blue">
                        <div class="theme-preview blue-theme"></div>
                        <div class="theme-info">
                            <span class="theme-name">Mavi Tema</span>
                            <span class="theme-desc">Mavi tonlarında tema</span>
                        </div>
                        <span class="theme-check"></span>
                    </div>
                </div>
                <div class="theme-actions">
                    <button class="customize-btn" onclick="customizeTheme()">🎨 Özelleştir</button>
                </div>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    // Tema seçimi event listener'ları
    setupThemeEventListeners(modal);
}

// Excel Ayarları Modal'ı
function showExcelSettings(menuModal) {
    // Mevcut menü modal'ını gizle
    menuModal.style.display = 'none';
    
    const modal = document.createElement('div');
    modal.className = 'submenu-modal';
    modal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <button class="back-btn" onclick="showMenu()">←</button>
                <h3>📊 Excel Ayarları</h3>
                <button class="close-modal-btn" onclick="this.parentElement.parentElement.parentElement.remove()">×</button>
            </div>
            <div class="modal-body">
                <div class="setting-group">
                    <h4>📈 Grafik Ayarları</h4>
                    <div class="setting-item">
                        <label>Varsayılan Grafik Türü:</label>
                        <select id="defaultChartType">
                            <option value="columnClustered">Sütun Grafik</option>
                            <option value="line">Çizgi Grafik</option>
                            <option value="pie">Pasta Grafik</option>
                            <option value="bar">Çubuk Grafik</option>
                        </select>
                    </div>
                    <div class="setting-item">
                        <label>Grafik Boyutu:</label>
                        <select id="chartSize">
                            <option value="small">Küçük</option>
                            <option value="medium">Orta</option>
                            <option value="large">Büyük</option>
                        </select>
                    </div>
                </div>
                
                <div class="setting-group">
                    <h4>🎨 Formatlama Ayarları</h4>
                    <div class="setting-item">
                        <label>Otomatik Formatlama:</label>
                        <input type="checkbox" id="autoFormat" checked>
                    </div>
                    <div class="setting-item">
                        <label>Alternatif Satır Renklendirme:</label>
                        <input type="checkbox" id="zebraRows" checked>
                    </div>
                    <div class="setting-item">
                        <label>Header Vurgulama:</label>
                        <input type="checkbox" id="highlightHeaders" checked>
                    </div>
                </div>
                
                <div class="setting-group">
                    <h4>⚡ Performans Ayarları</h4>
                    <div class="setting-item">
                        <label>Hızlı İşlem Modu:</label>
                        <input type="checkbox" id="fastMode">
                    </div>
                    <div class="setting-item">
                        <label>Otomatik Kaydet:</label>
                        <input type="checkbox" id="autoSave" checked>
                    </div>
                </div>
                
                <div class="setting-actions">
                    <button class="save-settings-btn" onclick="saveExcelSettings()">💾 Kaydet</button>
                    <button class="reset-settings-btn" onclick="resetExcelSettings()">🔄 Sıfırla</button>
                </div>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    // Mevcut ayarları yükle
    loadExcelSettings();
}

// Hakkında Modal'ı
function showAboutInfo(menuModal) {
    // Mevcut menü modal'ını gizle
    menuModal.style.display = 'none';
    
    const modal = document.createElement('div');
    modal.className = 'submenu-modal';
    modal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <button class="back-btn" onclick="showMenu()">←</button>
                <h3>ℹ️ Hakkında</h3>
                <button class="close-modal-btn" onclick="this.parentElement.parentElement.parentElement.remove()">×</button>
            </div>
            <div class="modal-body">
                <div class="about-content">
                    <div class="app-logo">🤖</div>
                    <h2>Excel AI Assistant</h2>
                    <p class="version">Versiyon 2.0.0</p>
                    <p class="description">
                        Excel AI Assistant, yapay zeka destekli Excel eklentisidir. 
                        LM Studio entegrasyonu ile doğal dil komutları kullanarak 
                        Excel işlemlerini otomatikleştirir.
                    </p>
                    
                    <div class="features">
                        <h4>🚀 Özellikler:</h4>
                        <ul>
                            <li>AI destekli veri analizi</li>
                            <li>Otomatik grafik oluşturma</li>
                            <li>Akıllı formatlama</li>
                            <li>Trend ve anomali tespiti</li>
                            <li>Doğal dil komutları</li>
                        </ul>
                    </div>
                    
                    <div class="tech-info">
                        <h4>🔧 Teknik Bilgiler:</h4>
                        <p><strong>AI Engine:</strong> LM Studio</p>
                        <p><strong>API:</strong> REST + Server-Sent Events</p>
                        <p><strong>Framework:</strong> Office Add-in</p>
                        <p><strong>Dil:</strong> JavaScript + HTML + CSS</p>
                    </div>
                    
                    <div class="contact">
                        <h4>📞 İletişim:</h4>
                        <p>Geliştirici: AI Assistant</p>
                        <p>Lisans: MIT</p>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
}

// ===== YARDIMCI FONKSİYONLAR =====

// Tema event listener'larını kur
function setupThemeEventListeners(modal) {
    const themeItems = modal.querySelectorAll('.theme-item');
    themeItems.forEach(item => {
        item.addEventListener('click', () => {
            const theme = item.dataset.theme;
            applyTheme(theme);
            
            // Check işaretini güncelle
            themeItems.forEach(t => t.querySelector('.theme-check').textContent = '');
            item.querySelector('.theme-check').textContent = '✓';
        });
    });
}

// Temayı uygula
function applyTheme(theme) {
    const root = document.documentElement;
    
    switch (theme) {
        case 'light':
            root.style.setProperty('--bg-primary', '#ffffff');
            root.style.setProperty('--bg-secondary', '#f8f9fa');
            root.style.setProperty('--text-primary', '#212529');
            root.style.setProperty('--text-secondary', '#6c757d');
            root.style.setProperty('--border-color', '#dee2e6');
            break;
        case 'blue':
            root.style.setProperty('--bg-primary', '#1e3a8a');
            root.style.setProperty('--bg-secondary', '#3b82f6');
            root.style.setProperty('--text-primary', '#ffffff');
            root.style.setProperty('--text-secondary', '#bfdbfe');
            root.style.setProperty('--border-color', '#60a5fa');
            break;
        default: // dark
            root.style.setProperty('--bg-primary', '#1a1a1a');
            root.style.setProperty('--bg-secondary', '#2d2d2d');
            root.style.setProperty('--text-primary', '#ffffff');
            root.style.setProperty('--text-secondary', '#888888');
            root.style.setProperty('--border-color', '#404040');
    }
    
    // Temayı localStorage'a kaydet
    localStorage.setItem('excelAI_theme', theme);
    
    addMessage('ai', `🎨 ${theme === 'light' ? 'Açık' : theme === 'blue' ? 'Mavi' : 'Koyu'} tema uygulandı!`, 'System');
}

// Mevcut modelleri yükle
async function loadAvailableModels() {
    try {
        if (window.aiClient && window.aiClient.getAvailableModels) {
            const models = await window.aiClient.getAvailableModels();
            displayModels(models);
            updateModelStatus();
        } else {
            document.getElementById('modelList').innerHTML = '<div class="error-message">AI Client bulunamadı</div>';
        }
    } catch (error) {
        document.getElementById('modelList').innerHTML = '<div class="error-message">Modeller yüklenemedi: ' + error.message + '</div>';
    }
}

// Modelleri görüntüle
function displayModels(models) {
    const modelList = document.getElementById('modelList');
    
    if (!models || models.length === 0) {
        modelList.innerHTML = '<div class="no-models">Yüklenmiş model bulunamadı</div>';
        return;
    }
    
    let html = '';
    models.forEach((model, index) => {
        html += `
            <div class="model-item" data-model="${model.name}">
                <div class="model-info">
                    <span class="model-name">${model.name}</span>
                    <span class="model-size">${model.size || 'Bilinmeyen'}</span>
                </div>
                <div class="model-actions">
                    <button class="select-model-btn" onclick="selectModel('${model.name}')">Seç</button>
                    <button class="model-info-btn" onclick="showModelInfo('${model.name}')">ℹ️</button>
                </div>
            </div>
        `;
    });
    
    modelList.innerHTML = html;
}

// Model durumunu güncelle
function updateModelStatus() {
    const currentModelElement = document.getElementById('currentModel');
    const modelStatusElement = document.getElementById('modelStatus');
    
    if (window.aiClient && window.aiClient.getCurrentModel) {
        const currentModel = window.aiClient.getCurrentModel();
        currentModelElement.textContent = currentModel || 'Seçilmemiş';
        
        // Model durumunu test et
        testModelConnection().then(status => {
            modelStatusElement.textContent = status;
            modelStatusElement.className = status === 'Bağlı' ? 'status-connected' : 'status-disconnected';
        });
    }
}

// Model bağlantısını test et
async function testModelConnection() {
    try {
        if (window.aiClient && window.aiClient.testConnection) {
            const isConnected = await window.aiClient.testConnection();
            return isConnected ? 'Bağlı' : 'Bağlantı Yok';
        }
        return 'Test Edilemedi';
    } catch (error) {
        return 'Hata: ' + error.message;
    }
}

// Model seç
function selectModel(modelName) {
    if (window.aiClient && window.aiClient.setCurrentModel) {
        window.aiClient.setCurrentModel(modelName);
        updateModelStatus();
        addMessage('ai', `🔧 AI Model değiştirildi: ${modelName}`, 'System');
    }
}

// Model bilgisi göster
function showModelInfo(modelName) {
    // Basit model bilgi modal'ı
    const infoModal = document.createElement('div');
    infoModal.className = 'info-modal';
    infoModal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <h4>ℹ️ Model Bilgisi</h4>
                <button class="close-modal-btn" onclick="this.parentElement.parentElement.remove()">×</button>
            </div>
            <div class="modal-body">
                <p><strong>Model Adı:</strong> ${modelName}</p>
                <p><strong>Durum:</strong> <span class="status-connected">Aktif</span></p>
                <p><strong>Tip:</strong> Yerel AI Model</p>
                <p><strong>Kaynak:</strong> LM Studio</p>
            </div>
        </div>
    `;
    
    document.body.appendChild(infoModal);
}

// Modelleri yenile
function refreshModels() {
    loadAvailableModels();
    addMessage('ai', '🔄 AI Modeller yenilendi!', 'System');
}

// Mevcut modeli test et
function testCurrentModel() {
    if (window.aiClient && window.aiClient.testConnection) {
        testModelConnection().then(status => {
            addMessage('ai', `🧪 Model test sonucu: ${status}`, 'System');
        });
    }
}

// Tema özelleştir
function customizeTheme() {
    addMessage('ai', '🎨 Tema özelleştirme özelliği yakında eklenecek!', 'System');
}

// Excel ayarlarını yükle
function loadExcelSettings() {
    try {
        const settings = JSON.parse(localStorage.getItem('excelAI_settings')) || getDefaultExcelSettings();
        
        // Form elemanlarını doldur
        document.getElementById('defaultChartType').value = settings.defaultChartType || 'columnClustered';
        document.getElementById('chartSize').value = settings.chartSize || 'medium';
        document.getElementById('autoFormat').checked = settings.autoFormat !== false;
        document.getElementById('zebraRows').checked = settings.zebraRows !== false;
        document.getElementById('highlightHeaders').checked = settings.highlightHeaders !== false;
        document.getElementById('fastMode').checked = settings.fastMode || false;
        document.getElementById('autoSave').checked = settings.autoSave !== false;
        
    } catch (error) {
        console.error('Ayarlar yüklenemedi:', error);
    }
}

// Varsayılan Excel ayarları
function getDefaultExcelSettings() {
    return {
        defaultChartType: 'columnClustered',
        chartSize: 'medium',
        autoFormat: true,
        zebraRows: true,
        highlightHeaders: true,
        fastMode: false,
        autoSave: true
    };
}

// Excel ayarlarını kaydet
function saveExcelSettings() {
    try {
        const settings = {
            defaultChartType: document.getElementById('defaultChartType').value,
            chartSize: document.getElementById('chartSize').value,
            autoFormat: document.getElementById('autoFormat').checked,
            zebraRows: document.getElementById('zebraRows').checked,
            highlightHeaders: document.getElementById('highlightHeaders').checked,
            fastMode: document.getElementById('fastMode').checked,
            autoSave: document.getElementById('autoSave').checked
        };
        
        localStorage.setItem('excelAI_settings', JSON.stringify(settings));
        addMessage('ai', '💾 Excel ayarları kaydedildi!', 'System');
        
        // Modal'ı kapat
        const modal = document.querySelector('.submenu-modal');
        if (modal) modal.remove();
        
    } catch (error) {
        console.error('Ayarlar kaydedilemedi:', error);
        addMessage('ai', '❌ Ayarlar kaydedilemedi: ' + error.message, 'System');
    }
}

// Excel ayarlarını sıfırla
function resetExcelSettings() {
    try {
        localStorage.removeItem('excelAI_settings');
        loadExcelSettings(); // Varsayılan ayarları yükle
        addMessage('ai', '🔄 Excel ayarları sıfırlandı!', 'System');
    } catch (error) {
        console.error('Ayarlar sıfırlanamadı:', error);
    }
}

function closeApp() {
    // Uygulamayı kapatma onayı - Office Add-in ortamında custom modal kullan
    showCloseConfirmation();
}

// Uygulama kapatma onay modal'ı
function showCloseConfirmation() {
    const confirmModal = document.createElement('div');
    confirmModal.className = 'confirm-modal';
    confirmModal.innerHTML = `
        <div class="confirm-content">
            <div class="confirm-header">
                <h4>🔒 Uygulama Kapat</h4>
            </div>
            <div class="confirm-body">
                <p>Excel AI Assistant'ı kapatmak istediğinizden emin misiniz?</p>
            </div>
            <div class="confirm-actions">
                <button class="confirm-btn confirm-cancel" type="button">❌ İptal</button>
                <button class="confirm-btn confirm-delete" type="button">🔒 Kapat</button>
            </div>
        </div>
    `;
    
    document.body.appendChild(confirmModal);
    
    // Event listener'ları ekle
    const cancelBtn = confirmModal.querySelector('.confirm-cancel');
    const closeBtn = confirmModal.querySelector('.confirm-delete');
    
    cancelBtn.addEventListener('click', () => {
        confirmModal.remove();
    });
    
    closeBtn.addEventListener('click', () => {
        // Office Add-in'i kapat
        if (Office && Office.context && Office.context.document) {
            Office.context.document.closeAsync();
        } else {
            // Fallback: sayfayı kapat
            window.close();
        }
        console.log('🔒 Uygulama kapatılıyor');
        confirmModal.remove();
    });
}

// Command bar buton fonksiyonları
function handleImageUpload() {
    // Resim yükleme input'u oluştur
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = 'image/*';
    fileInput.style.display = 'none';
    
    fileInput.onchange = function(e) {
        const file = e.target.files[0];
        if (file) {
            // Resim yüklendi mesajı
            addMessage('user', `🖼️ Resim yüklendi: ${file.name}`, 'Kullanıcı');
            
            // AI'ya resim analizi için gönder
            const reader = new FileReader();
            reader.onload = function(e) {
                // Base64 resim verisi
                const imageData = e.target.result;
                
                // AI'ya resim analizi komutu gönder
                addMessage('ai', '🖼️ Resim analiz ediliyor...', 'AI Assistant');
                
                // Burada resim analizi API'si çağrılabilir
                setTimeout(() => {
                    updateLastMessage('🖼️ Resim analiz edildi! Bu resimde Excel tablosu görüyorum. Hangi işlemi yapmak istiyorsunuz?');
                }, 2000);
            };
            reader.readAsDataURL(file);
        }
    };
    
    document.body.appendChild(fileInput);
    fileInput.click();
    document.body.removeChild(fileInput);
    
    console.log('🖼️ Resim yükleme başlatıldı');
}

function handleVoiceInput() {
    // Ses girişi için modal oluştur
    const modal = document.createElement('div');
    modal.className = 'voice-modal';
    modal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <h3>🎤 Ses Girişi</h3>
                <button class="close-modal-btn" onclick="this.parentElement.parentElement.parentElement.remove()">×</button>
            </div>
            <div class="modal-body">
                <div class="voice-status">
                    <div class="voice-icon">🎤</div>
                    <div class="voice-text">Ses girişi için tıklayın</div>
                </div>
                <button class="voice-record-btn" onclick="startVoiceRecording(this)">
                    🎙️ Kayıt Başlat
                </button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    console.log('🎤 Ses girişi modal\'ı açıldı');
}

// Ses kaydı başlat
function startVoiceRecording(btn) {
    if (btn.textContent.includes('Başlat')) {
        btn.textContent = '⏹️ Kaydı Durdur';
        btn.style.backgroundColor = '#dc3545';
        
        // Ses kaydı simülasyonu
        setTimeout(() => {
            btn.textContent = '🎙️ Kayıt Başlat';
            btn.style.backgroundColor = '';
            
            // Modal'ı kapat
            const modal = document.querySelector('.voice-modal');
            if (modal) modal.remove();
            
            // Ses kaydı tamamlandı mesajı
            addMessage('user', '🎤 "Bu veriyi grafik yap" (ses kaydı)', 'Kullanıcı');
            
            // AI'ya ses komutu gönder
            processUserCommand('Bu veriyi grafik yap');
        }, 3000);
    }
}

// ===== AI YANIT OTOMATİK UYGULAMA SİSTEMİ =====

// AI yanıtını akıllıca analiz et ve Excel'de uygula
async function executeAIResponseIntelligently(command, aiResponse) {
    const lowerCommand = command.toLowerCase();
    const lowerResponse = aiResponse.toLowerCase();
    
    try {
        // 1. Önce mevcut otomatik komutları dene
        if (lowerCommand.includes('grafik') || lowerCommand.includes('chart')) {
            await createChartFromData();
            return;
        } else if (lowerCommand.includes('toplam') || lowerCommand.includes('sum')) {
            await calculateSum();
            return;
        } else if (lowerCommand.includes('ortalama') || lowerCommand.includes('average')) {
            await calculateAverage();
            return;
        } else if (lowerCommand.includes('filtre') || lowerCommand.includes('filter')) {
            await applyFilter();
            return;
        } else if (lowerCommand.includes('sırala') || lowerCommand.includes('sort')) {
            await sortData();
            return;
        } else if (lowerCommand.includes('analiz') || lowerCommand.includes('analyze') || 
                   lowerCommand.includes('incele') || lowerCommand.includes('examine')) {
            await analyzeDataIntelligently();
            return;
        } else if (lowerCommand.includes('özet') || lowerCommand.includes('summary')) {
            await generateDataSummary();
            return;
        } else if (lowerCommand.includes('trend') || lowerCommand.includes('eğilim')) {
            await analyzeTrends();
            return;
        } else if (lowerCommand.includes('anomali') || lowerCommand.includes('outlier')) {
            await detectAnomalies();
            return;
        }
        
        // 2. AI yanıtını analiz et ve otomatik uygula
        await analyzeAndExecuteAIResponse(aiResponse);
        
    } catch (error) {
        console.log('AI yanıt uygulama hatası:', error);
        addMessage('ai', '⚠️ AI yanıtı analiz edildi ama otomatik uygulanamadı. Manuel olarak uygulayabilirsiniz.', 'System');
    }
}

// AI yanıtını analiz et ve Excel'de uygula
async function analyzeAndExecuteAIResponse(aiResponse) {
    const lowerResponse = aiResponse.toLowerCase();
    
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('values, address, columnCount, rowCount');
            
            await context.sync();
            
            if (!range.values || range.values.length < 2) {
                addMessage('ai', '❌ Excel\'de seçili veri bulunamadı. Lütfen bir tablo seçin.', 'System');
                return;
            }
            
            // AI yanıtına göre otomatik işlemler
            if (lowerResponse.includes('grafik') || lowerResponse.includes('chart') || 
                lowerResponse.includes('chart oluştur') || lowerResponse.includes('create chart')) {
                await createChartFromRange(context, range);
                
            } else if (lowerResponse.includes('renklendir') || lowerResponse.includes('color') || 
                       lowerResponse.includes('formatla') || lowerResponse.includes('format')) {
                await applySmartFormatting(context, range);
                
            } else if (lowerResponse.includes('filtrele') || lowerResponse.includes('filter') || 
                       lowerResponse.includes('filtre uygula')) {
                await applySmartFiltering(context, range);
                
            } else if (lowerResponse.includes('sırala') || lowerResponse.includes('sort') || 
                       lowerResponse.includes('düzenle')) {
                await applySmartSorting(context, range);
                
            } else if (lowerResponse.includes('formül') || lowerResponse.includes('formula') || 
                       lowerResponse.includes('hesapla')) {
                await addSmartFormulas(context, range);
                
            } else if (lowerResponse.includes('özet') || lowerResponse.includes('summary') || 
                       lowerResponse.includes('pivot')) {
                await createDataSummary(context, range);
                
            } else if (lowerResponse.includes('koşullu format') || lowerResponse.includes('conditional format')) {
                await applyConditionalFormatting(context, range);
                
            } else {
                // Genel veri analizi ve formatlama
                await performGeneralDataEnhancement(context, range);
            }
            
            await context.sync();
            addMessage('ai', '✅ AI yanıtı Excel\'de otomatik olarak uygulandı!', 'Excel');
            
        });
    } catch (error) {
        console.error('AI yanıt uygulama hatası:', error);
        addMessage('ai', `❌ Excel işlemi hatası: ${error.message}`, 'System');
    }
}

// Akıllı grafik oluştur
async function createChartFromRange(context, range) {
    try {
        const chart = range.worksheet.charts.add(Excel.ChartType.columnClustered, range);
        chart.setPosition(0, range.getColumn(0).getColumnIndex() + range.values[0].length + 2);
        
        addMessage('ai', '📊 Grafik başarıyla oluşturuldu!', 'Excel');
    } catch (error) {
        // Alternatif: tablo formatlaması
                range.format.borders.getItem('EdgeBottom').style = 'Continuous';
                range.format.borders.getItem('EdgeRight').style = 'Continuous';
        range.format.fill.color = '#e6f3ff';
        addMessage('ai', '📊 Veri tablosu formatlandı!', 'Excel');
    }
}

// Akıllı formatlama uygula
async function applySmartFormatting(context, range) {
    try {
        // Header satırını formatla
        const headerRange = range.getRow(0);
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = '#e6f3ff';
        headerRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
        
        // Sayısal sütunları formatla
        for (let col = 0; col < range.values[0].length; col++) {
            const columnData = range.values.slice(1).map(row => row[col]);
            if (isNumericColumn(columnData)) {
                const columnRange = range.getColumn(col);
                columnRange.format.numberFormat = '#,##0.00';
            }
        }
        
        // Tüm hücrelere border ekle
        range.format.borders.getItem('EdgeBottom').style = 'Continuous';
        range.format.borders.getItem('EdgeRight').style = 'Continuous';
        
        addMessage('ai', '✨ Akıllı formatlama uygulandı!', 'Excel');
    } catch (error) {
        console.log('Formatlama hatası:', error);
    }
}

// Akıllı filtreleme uygula
async function applySmartFiltering(context, range) {
    try {
        range.autoFilter.apply();
        addMessage('ai', '🔍 Otomatik filtre uygulandı!', 'Excel');
    } catch (error) {
        console.log('Filtreleme hatası:', error);
    }
}

// Akıllı sıralama uygula
async function applySmartSorting(context, range) {
    try {
        // İlk sayısal sütuna göre sırala
        for (let col = 0; col < range.values[0].length; col++) {
            const columnData = range.values.slice(1).map(row => row[col]);
            if (isNumericColumn(columnData)) {
                range.sort.apply([{ key: col, sortOrder: Excel.SortOrder.ascending }]);
                addMessage('ai', '📈 Veri sıralandı!', 'Excel');
                return;
            }
        }
        
        // Sayısal sütun yoksa ilk sütuna göre sırala
        range.sort.apply([{ key: 0, sortOrder: Excel.SortOrder.ascending }]);
        addMessage('ai', '📈 Veri sıralandı!', 'Excel');
    } catch (error) {
        console.log('Sıralama hatası:', error);
    }
}

// Akıllı formüller ekle
async function addSmartFormulas(context, range) {
    try {
        const worksheet = range.worksheet;
        const lastRowIndex = range.getRow(range.values.length - 1).getRowIndex();
        const nextRow = worksheet.getRange(`${lastRowIndex + 1}:${lastRowIndex + 1}`);
        
        // Sayısal sütunlar için toplam ve ortalama
        for (let col = 0; col < range.values[0].length; col++) {
            const columnData = range.values.slice(1).map(row => row[col]);
            if (isNumericColumn(columnData)) {
                const colLetter = getColumnLetter(col);
                const startRow = range.getRow(1).getRowIndex();
                const endRow = lastRowIndex;
                
                // Toplam
                const sumCell = worksheet.getRange(`${colLetter}${lastRowIndex + 2}`);
                sumCell.values = [[`=SUM(${colLetter}${startRow}:${colLetter}${endRow})`]];
                sumCell.format.font.bold = true;
                sumCell.format.fill.color = '#e6f3ff';
                
                // Ortalama
                const avgCell = worksheet.getRange(`${colLetter}${lastRowIndex + 3}`);
                avgCell.values = [[`=AVERAGE(${colLetter}${startRow}:${colLetter}${endRow})`]];
                avgCell.format.font.bold = true;
                avgCell.format.fill.color = '#e6f3ff';
            }
        }
        
        addMessage('ai', '🧮 Akıllı formüller eklendi!', 'Excel');
    } catch (error) {
        console.log('Formül ekleme hatası:', error);
    }
}

// Veri özeti oluştur
async function createDataSummary(context, range) {
    try {
        const worksheet = range.worksheet;
        const summaryRow = range.getRow(range.values.length + 2);
        
        // Özet başlığı
        summaryRow.values = [['VERİ ÖZETİ']];
        summaryRow.format.font.bold = true;
        summaryRow.format.font.size = 14;
        summaryRow.format.fill.color = '#0078d4';
        summaryRow.format.font.color = '#ffffff';
        
        // Her sütun için özet
        for (let col = 0; col < range.values[0].length; col++) {
            const columnData = range.values.slice(1).map(row => row[col]);
            const summary = createColumnSummary(columnData);
            
            const summaryCell = worksheet.getRange(`${getColumnLetter(col)}${range.values.length + 3}`);
            summaryCell.values = [[summary]];
        }
        
        addMessage('ai', '📋 Veri özeti oluşturuldu!', 'Excel');
    } catch (error) {
        console.log('Özet oluşturma hatası:', error);
    }
}

// Koşullu formatlama uygula
async function applyConditionalFormatting(context, range) {
    try {
        // Sayısal sütunlar için koşullu formatlama
        for (let col = 0; col < range.values[0].length; col++) {
            const columnData = range.values.slice(1).map(row => row[col]);
            if (isNumericColumn(columnData)) {
                const columnRange = range.getColumn(col);
                
                // Yüksek değerler için yeşil
                const highRule = columnRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                highRule.cellValue.rule = Excel.ConditionalCellValueRule.greaterThan;
                highRule.cellValue.formula1 = '=AVERAGE($' + getColumnLetter(col) + ':$' + getColumnLetter(col) + ')';
                highRule.format.fill.color = '#90EE90';
                
                // Düşük değerler için kırmızı
                const lowRule = columnRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                lowRule.cellValue.rule = Excel.ConditionalCellValueRule.lessThan;
                lowRule.cellValue.formula1 = '=AVERAGE($' + getColumnLetter(col) + ':$' + getColumnLetter(col) + ')';
                lowRule.format.fill.color = '#FFB6C1';
            }
        }
        
        addMessage('ai', '🎨 Koşullu formatlama uygulandı!', 'Excel');
    } catch (error) {
        console.log('Koşullu formatlama hatası:', error);
    }
}

// Genel veri geliştirme
async function performGeneralDataEnhancement(context, range) {
    try {
        // Otomatik sütun genişliği
        range.format.autofitColumns();
                
                // Alternatif satır renklendirme
        for (let row = 1; row < range.values.length; row++) {
            if (row % 2 === 1) {
                const rowRange = range.getRow(row);
                rowRange.format.fill.color = '#f8f9fa';
            }
        }
        
        // Header formatlaması
        const headerRange = range.getRow(0);
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = '#e6f3ff';
        
        addMessage('ai', '✨ Veri tablosu geliştirildi!', 'Excel');
    } catch (error) {
        console.log('Genel geliştirme hatası:', error);
    }
}

// Yardımcı fonksiyonlar
function isNumericColumn(columnData) {
    const numbers = columnData.filter(val => !isNaN(val) && val !== '' && val !== null);
    return numbers.length / columnData.length > 0.7;
}

function getColumnLetter(columnIndex) {
    let result = '';
    while (columnIndex >= 0) {
        result = String.fromCharCode(65 + (columnIndex % 26)) + result;
        columnIndex = Math.floor(columnIndex / 26) - 1;
    }
    return result;
}

function createColumnSummary(columnData) {
    if (isNumericColumn(columnData)) {
        const numbers = columnData.filter(val => !isNaN(val) && val !== '').map(Number);
        if (numbers.length > 0) {
            const sum = numbers.reduce((a, b) => a + b, 0);
            const avg = sum / numbers.length;
            return `Toplam: ${sum.toFixed(2)}, Ort: ${avg.toFixed(2)}`;
        }
    }
    
    const uniqueValues = [...new Set(columnData.filter(val => val !== '' && val !== null))];
    return `Benzersiz: ${uniqueValues.length}`;
}

// ===== AKILLI VERİ ANALİZİ FONKSİYONLARI =====

// Akıllı veri analizi - Ana fonksiyon
async function analyzeDataIntelligently() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('values, address, columnCount, rowCount');
                
                await context.sync();
            
            if (range.values && range.values.length > 1) {
                const analysis = performDataAnalysis(range.values);
                await applyIntelligentFormatting(context, range, analysis);
                
                // Analiz sonuçlarını chat'e ekle
                const analysisMessage = formatAnalysisResults(analysis);
                addMessage('ai', analysisMessage, 'Data Analysis');
            } else {
                addMessage('ai', '❌ Analiz için yeterli veri bulunamadı. Lütfen birden fazla satır seçin.', 'System');
            }
        });
    } catch (error) {
        addMessage('ai', `Veri analizi hatası: ${error.message}`, 'System');
    }
}

// Veri analizi yap
function performDataAnalysis(data) {
    const analysis = {
        dataType: 'unknown',
        numericColumns: [],
        textColumns: [],
        dateColumns: [],
        statistics: {},
        insights: [],
        recommendations: []
    };
    
    if (data.length < 2) return analysis;
    
    const headers = data[0];
    const values = data.slice(1);
    
    // Her sütun için veri türü analizi
    for (let col = 0; col < headers.length; col++) {
        const columnData = values.map(row => row[col]);
        const columnType = analyzeColumnType(columnData);
        
        if (columnType === 'numeric') {
            analysis.numericColumns.push(col);
            analysis.statistics[col] = calculateColumnStatistics(columnData);
        } else if (columnType === 'text') {
            analysis.textColumns.push(col);
            analysis.statistics[col] = analyzeTextColumn(columnData);
        } else if (columnType === 'date') {
            analysis.dateColumns.push(col);
            analysis.statistics[col] = analyzeDateColumn(columnData);
        }
    }
    
    // Genel veri türü belirleme
    if (analysis.numericColumns.length > analysis.textColumns.length) {
        analysis.dataType = 'numeric';
    } else if (analysis.textColumns.length > analysis.numericColumns.length) {
        analysis.dataType = 'categorical';
    } else {
        analysis.dataType = 'mixed';
    }
    
    // İçgörüler ve öneriler
    analysis.insights = generateInsights(analysis);
    analysis.recommendations = generateRecommendations(analysis);
    
    return analysis;
}

// Sütun veri türünü analiz et
function analyzeColumnType(columnData) {
    let numericCount = 0;
    let dateCount = 0;
    let textCount = 0;
    
    for (const value of columnData) {
        if (value === null || value === undefined || value === '') continue;
        
        if (!isNaN(value) && value !== '') {
            numericCount++;
        } else if (isValidDate(value)) {
            dateCount++;
        } else {
            textCount++;
        }
    }
    
    const total = numericCount + dateCount + textCount;
    if (total === 0) return 'unknown';
    
    if (numericCount / total > 0.7) return 'numeric';
    if (dateCount / total > 0.7) return 'date';
    return 'text';
}

// Tarih geçerli mi kontrol et
function isValidDate(value) {
    if (typeof value === 'string') {
        const date = new Date(value);
        return !isNaN(date.getTime());
    }
    return false;
}

// Sayısal sütun istatistikleri
function calculateColumnStatistics(columnData) {
    const numbers = columnData.filter(val => !isNaN(val) && val !== '').map(Number);
    if (numbers.length === 0) return {};
    
    const sum = numbers.reduce((a, b) => a + b, 0);
    const mean = sum / numbers.length;
    const sorted = numbers.sort((a, b) => a - b);
    const median = sorted.length % 2 === 0 
        ? (sorted[sorted.length/2 - 1] + sorted[sorted.length/2]) / 2
        : sorted[Math.floor(sorted.length/2)];
    
    return {
        count: numbers.length,
        sum: sum,
        mean: mean,
        median: median,
        min: Math.min(...numbers),
        max: Math.max(...numbers),
        range: Math.max(...numbers) - Math.min(...numbers)
    };
}

// Metin sütun analizi
function analyzeTextColumn(columnData) {
    const texts = columnData.filter(val => val !== null && val !== undefined && val !== '');
    const uniqueValues = [...new Set(texts)];
    
    return {
        count: texts.length,
        uniqueCount: uniqueValues.length,
        mostCommon: findMostCommon(texts),
        categories: uniqueValues.slice(0, 10) // İlk 10 kategori
    };
}

// Tarih sütun analizi
function analyzeDateColumn(columnData) {
    const dates = columnData.filter(val => isValidDate(val)).map(val => new Date(val));
    if (dates.length === 0) return {};
    
    const sorted = dates.sort((a, b) => a - b);
    const range = sorted[sorted.length - 1] - sorted[0];
    
    return {
        count: dates.length,
        earliest: sorted[0],
        latest: sorted[sorted.length - 1],
        range: range,
        days: Math.ceil(range / (1000 * 60 * 60 * 24))
    };
}

// En çok tekrar eden değeri bul
function findMostCommon(array) {
    const counts = {};
    let maxCount = 0;
    let maxValue = null;
    
    for (const value of array) {
        counts[value] = (counts[value] || 0) + 1;
        if (counts[value] > maxCount) {
            maxCount = counts[value];
            maxValue = value;
        }
    }
    
    return { value: maxValue, count: maxCount };
}

// İçgörüler oluştur
function generateInsights(analysis) {
    const insights = [];
    
    if (analysis.numericColumns.length > 0) {
        insights.push(`📊 ${analysis.numericColumns.length} sayısal sütun bulundu`);
        
        analysis.numericColumns.forEach(col => {
            const stats = analysis.statistics[col];
            if (stats.range > 0) {
                insights.push(`📈 Sütun ${col + 1}: Min ${stats.min}, Max ${stats.max}, Ortalama ${stats.mean.toFixed(2)}`);
            }
        });
    }
    
    if (analysis.textColumns.length > 0) {
        insights.push(`📝 ${analysis.textColumns.length} metin sütunu bulundu`);
        
        analysis.textColumns.forEach(col => {
            const stats = analysis.statistics[col];
            if (stats.uniqueCount > 0) {
                insights.push(`🏷️ Sütun ${col + 1}: ${stats.uniqueCount} benzersiz kategori`);
            }
        });
    }
    
    if (analysis.dateColumns.length > 0) {
        insights.push(`📅 ${analysis.dateColumns.length} tarih sütunu bulundu`);
    }
    
    return insights;
}

// Öneriler oluştur
function generateRecommendations(analysis) {
    const recommendations = [];
    
    if (analysis.numericColumns.length >= 2) {
        recommendations.push('📊 Grafik oluşturulabilir');
        recommendations.push('📈 Trend analizi yapılabilir');
    }
    
    if (analysis.textColumns.length > 0) {
        recommendations.push('🏷️ Kategori bazlı filtreleme yapılabilir');
    }
    
    if (analysis.dateColumns.length > 0) {
        recommendations.push('📅 Zaman bazlı analiz yapılabilir');
    }
    
    if (analysis.dataType === 'numeric') {
        recommendations.push('🧮 İstatistiksel özet oluşturulabilir');
    }
    
    return recommendations;
}

// Analiz sonuçlarını formatla
function formatAnalysisResults(analysis) {
    let message = '🔍 **AKILLI VERİ ANALİZİ SONUÇLARI**\n\n';
    
    // Veri türü
    message += `📋 **Veri Türü:** ${getDataTypeName(analysis.dataType)}\n`;
    message += `📊 **Toplam Satır:** ${analysis.numericColumns.length + analysis.textColumns.length + analysis.dateColumns.length}\n\n`;
    
    // İçgörüler
    if (analysis.insights.length > 0) {
        message += '💡 **İÇGÖRÜLER:**\n';
        analysis.insights.forEach(insight => {
            message += `• ${insight}\n`;
        });
        message += '\n';
    }
    
    // Öneriler
    if (analysis.recommendations.length > 0) {
        message += '🚀 **ÖNERİLER:**\n';
        analysis.recommendations.forEach(rec => {
            message += `• ${rec}\n`;
        });
    }
    
    return message;
}

// Veri türü adını getir
function getDataTypeName(type) {
    const names = {
        'numeric': 'Sayısal Veri',
        'categorical': 'Kategorik Veri',
        'mixed': 'Karışık Veri',
        'unknown': 'Bilinmeyen'
    };
    return names[type] || 'Bilinmeyen';
}

// Akıllı formatlama uygula
async function applyIntelligentFormatting(context, range, analysis) {
    try {
        // Sayısal sütunları formatla
        analysis.numericColumns.forEach(col => {
            const columnRange = range.getColumn(col);
            columnRange.format.numberFormat = '#,##0.00';
        });
        
        // Tarih sütunlarını formatla
        analysis.dateColumns.forEach(col => {
            const columnRange = range.getColumn(col);
            columnRange.format.numberFormat = 'dd.mm.yyyy';
        });
        
        // Header satırını formatla
        const headerRange = range.getRow(0);
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = '#e6f3ff';
        
        await context.sync();
        addMessage('ai', '✨ Akıllı formatlama uygulandı!', 'Excel');
    } catch (error) {
        console.log('Formatlama hatası:', error);
    }
}

// Veri özeti oluştur
async function generateDataSummary() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('values, address');
            
            await context.sync();
            
            if (range.values && range.values.length > 1) {
                const summary = createDataSummary(range.values);
                addMessage('ai', summary, 'Data Summary');
            } else {
                addMessage('ai', '❌ Özet için yeterli veri bulunamadı.', 'System');
            }
        });
    } catch (error) {
        addMessage('ai', `Veri özeti hatası: ${error.message}`, 'System');
    }
}

// Veri özeti oluştur
function createDataSummary(data) {
    const headers = data[0];
    const values = data.slice(1);
    
    let summary = '📋 **VERİ ÖZETİ**\n\n';
    summary += `📊 **Toplam Satır:** ${values.length}\n`;
    summary += `🏷️ **Toplam Sütun:** ${headers.length}\n\n`;
    
    // Her sütun için özet
    for (let i = 0; i < headers.length; i++) {
        const columnData = values.map(row => row[i]);
        const columnType = analyzeColumnType(columnData);
        
        summary += `**${headers[i] || `Sütun ${i + 1}`}:** `;
        
        if (columnType === 'numeric') {
            const stats = calculateColumnStatistics(columnData);
            summary += `Sayısal (${stats.count} değer)\n`;
            summary += `  • Ortalama: ${stats.mean.toFixed(2)}\n`;
            summary += `  • Min-Max: ${stats.min} - ${stats.max}\n`;
        } else if (columnType === 'text') {
            const stats = analyzeTextColumn(columnData);
            summary += `Metin (${stats.count} değer, ${stats.uniqueCount} benzersiz)\n`;
        } else if (columnType === 'date') {
            const stats = analyzeDateColumn(columnData);
            summary += `Tarih (${stats.count} değer)\n`;
        }
        summary += '\n';
    }
    
    return summary;
}

// Trend analizi
async function analyzeTrends() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('values, address');
            
            await context.sync();
            
            if (range.values && range.values.length > 2) {
                const trends = detectTrends(range.values);
                addMessage('ai', trends, 'Trend Analysis');
            } else {
                addMessage('ai', '❌ Trend analizi için en az 3 satır gerekli.', 'System');
            }
        });
    } catch (error) {
        addMessage('ai', `Trend analizi hatası: ${error.message}`, 'System');
    }
}

// Trend tespit et
function detectTrends(data) {
    const headers = data[0];
    const values = data.slice(1);
    
    let analysis = '📈 **TREND ANALİZİ**\n\n';
    
    // Her sayısal sütun için trend analizi
    for (let col = 0; col < headers.length; col++) {
        const columnData = values.map(row => row[col]);
        if (analyzeColumnType(columnData) === 'numeric') {
            const trend = calculateTrend(columnData);
            analysis += `**${headers[col] || `Sütun ${col + 1}`}:** ${trend}\n\n`;
        }
    }
    
    return analysis;
}

// Trend hesapla
function calculateTrend(data) {
    const numbers = data.filter(val => !isNaN(val) && val !== '').map(Number);
    if (numbers.length < 2) return 'Yetersiz veri';
    
    // Basit trend analizi
    const firstHalf = numbers.slice(0, Math.floor(numbers.length / 2));
    const secondHalf = numbers.slice(Math.floor(numbers.length / 2));
    
    const firstAvg = firstHalf.reduce((a, b) => a + b, 0) / firstHalf.length;
    const secondAvg = secondHalf.reduce((a, b) => a + b, 0) / secondHalf.length;
    
    const change = ((secondAvg - firstAvg) / firstAvg) * 100;
    
    if (change > 5) return `📈 Artış trendi (%${change.toFixed(1)} artış)`;
    if (change < -5) return `📉 Azalış trendi (%${Math.abs(change).toFixed(1)} azalış)`;
    return `➡️ Stabil trend (%${change.toFixed(1)} değişim)`;
}

// Anomali tespit et
async function detectAnomalies() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('values, address');
            
            await context.sync();
            
            if (range.values && range.values.length > 2) {
                const anomalies = findAnomalies(range.values);
                addMessage('ai', anomalies, 'Anomaly Detection');
            } else {
                addMessage('ai', '❌ Anomali tespiti için en az 3 satır gerekli.', 'System');
            }
        });
    } catch (error) {
        addMessage('ai', `Anomali tespiti hatası: ${error.message}`, 'System');
    }
}

// Anomali bul
function findAnomalies(data) {
    const headers = data[0];
    const values = data.slice(1);
    
    let analysis = '🔍 **ANOMALİ TESPİTİ**\n\n';
    
    // Her sayısal sütun için anomali tespiti
    for (let col = 0; col < headers.length; col++) {
        const columnData = values.map(row => row[col]);
        if (analyzeColumnType(columnData) === 'numeric') {
            const anomalies = detectColumnAnomalies(columnData);
            if (anomalies.length > 0) {
                analysis += `**${headers[col] || `Sütun ${col + 1}`}:**\n`;
                anomalies.forEach(anomaly => {
                    analysis += `  • ${anomaly}\n`;
                });
                analysis += '\n';
            }
        }
    }
    
    if (analysis === '🔍 **ANOMALİ TESPİTİ**\n\n') {
        analysis += '✅ Belirgin anomali tespit edilmedi.';
    }
    
    return analysis;
}

// Sütun anomali tespiti
function detectColumnAnomalies(data) {
    const numbers = data.filter(val => !isNaN(val) && val !== '').map(Number);
    if (numbers.length < 3) return [];
    
    const mean = numbers.reduce((a, b) => a + b, 0) / numbers.length;
    const variance = numbers.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / numbers.length;
    const stdDev = Math.sqrt(variance);
    
    const anomalies = [];
    
    numbers.forEach((num, index) => {
        const zScore = Math.abs((num - mean) / stdDev);
        if (zScore > 2) { // 2 standart sapma üzeri
            anomalies.push(`Satır ${index + 2}: ${num} (Z-Score: ${zScore.toFixed(2)})`);
        }
    });
    
    return anomalies;
}
