// LM Studio AI Client
class LMStudioClient {
    constructor() {
        this.baseUrl = 'http://192.168.1.5:1234';
        this.apiEndpoint = `${this.baseUrl}/v1/chat/completions`;
        this.currentModel = 'openai/gpt-oss-20b';
        this.availableModels = [];
        this.maxTokens = 2000;
        this.temperature = 0.7;
        this.timeout = 120000; // 2 dakika
        
        // Streaming kontrolü için flag'ler
        this.isStreaming = false;
        this.currentReader = null;
        
        // Modelleri yükle
        this.loadAvailableModels();
    }

    // API bağlantısını test et
    async testConnection() {
        try {
            const response = await fetch(`${this.baseUrl}/v1/models`, {
                method: 'GET',
                timeout: this.timeout
            });
            return response.ok;
        } catch (error) {
            console.error('LM Studio bağlantı hatası:', error);
            return false;
        }
    }

    // Mevcut modelleri yükler
    async loadAvailableModels() {
        try {
            const response = await fetch(`${this.baseUrl}/v1/models`, {
                method: 'GET',
                timeout: this.timeout
            });
            
            if (response.ok) {
                const modelsData = await response.json();
                this.availableModels = modelsData.data.map(model => model.id);
                
                if (this.availableModels.length > 0 && !this.availableModels.includes(this.currentModel)) {
                    this.currentModel = this.availableModels[0];
                }
                
                console.log(`✅ ${this.availableModels.length} model yüklendi`);
                return this.availableModels;
            } else {
                console.log(`⚠️ Modeller yüklenemedi: ${response.status}`);
                return [];
            }
        } catch (error) {
            console.error('⚠️ Model yükleme hatası:', error);
            return [];
        }
    }

    // Mevcut modelleri döndürür
    getAvailableModels() {
        return this.availableModels;
    }

    // Aktif modeli değiştirir
    setCurrentModel(modelId) {
        if (this.availableModels.includes(modelId)) {
            this.currentModel = modelId;
            console.log(`✅ Model değiştirildi: ${modelId}`);
            return true;
        } else {
            console.log(`❌ Model bulunamadı: ${modelId}`);
            return false;
        }
    }

    // Mevcut modeli döndürür
    getCurrentModel() {
        return this.currentModel;
    }
    
    // Streaming'i durdur
    stopStreaming() {
        if (this.isStreaming && this.currentReader) {
            this.isStreaming = false;
            this.currentReader.cancel();
            console.log('Streaming durduruldu');
        }
    }

    // AI'dan yanıt al (streaming olmadan)
    async generateResponse(prompt, systemMessage = "") {
        try {
            const messages = [];
            if (systemMessage) {
                messages.push({ role: "system", content: systemMessage });
            }
            messages.push({ role: "user", content: prompt });

            const payload = {
                model: this.currentModel,
                messages: messages,
                max_tokens: this.maxTokens,
                temperature: this.temperature,
                stream: false
            };

            const response = await fetch(this.apiEndpoint, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(payload)
            });

            if (response.ok) {
                const result = await response.json();
                return result.choices[0].message.content;
            } else {
                throw new Error(`API Hatası: ${response.status}`);
            }
        } catch (error) {
            console.error('AI yanıt hatası:', error);
            return null;
        }
    }

    // AI'dan streaming yanıt al (ChatGPT gibi)
    async generateStreamingResponse(prompt, systemMessage = "", onChunk) {
        try {
            const messages = [];
            if (systemMessage) {
                messages.push({ role: "system", content: systemMessage });
            }
            messages.push({ role: "user", content: prompt });

            const payload = {
                model: this.currentModel,
                messages: messages,
                max_tokens: this.maxTokens,
                temperature: this.temperature,
                stream: true
            };

            const response = await fetch(this.apiEndpoint, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(payload)
            });

            if (response.ok) {
                const reader = response.body.getReader();
                const decoder = new TextDecoder();
                let fullResponse = '';
                
                // Streaming'i durdurma kontrolü için flag
                this.isStreaming = true;
                this.currentReader = reader;

                while (true) {
                    // Streaming durduruldu mu kontrol et
                    if (!this.isStreaming) {
                        await reader.cancel();
                        break;
                    }
                    
                    const { done, value } = await reader.read();
                    if (done) break;

                    const chunk = decoder.decode(value);
                    const lines = chunk.split('\n');

                    for (const line of lines) {
                        if (line.startsWith('data: ')) {
                            const data = line.slice(6);
                            if (data === '[DONE]') {
                                this.isStreaming = false;
                                this.currentReader = null;
                                return fullResponse;
                            }

                            try {
                                const parsed = JSON.parse(data);
                                if (parsed.choices && parsed.choices[0] && parsed.choices[0].delta && parsed.choices[0].delta.content) {
                                    const content = parsed.choices[0].delta.content;
                                    fullResponse += content;
                                    
                                    // Her chunk'ı callback ile gönder
                                    if (onChunk) {
                                        onChunk(content, fullResponse);
                                    }
                                }
                            } catch (e) {
                                // JSON parse hatası, devam et
                            }
                        }
                    }
                }

                this.isStreaming = false;
                this.currentReader = null;
                return fullResponse;
            } else {
                throw new Error(`API Hatası: ${response.status}`);
            }
        } catch (error) {
            this.isStreaming = false;
            this.currentReader = null;
            console.error('AI streaming yanıt hatası:', error);
            return null;
        }
    }

    // Excel veri analizi için özel prompt
    async analyzeExcelData(dataDescription, question) {
        const systemMsg = "Sen bir Excel uzmanısın. Verilen veri seti üzerinde işlem yapmak için gerekli Excel komutlarını ve formüllerini önerirsin.";
        const prompt = `Veri: ${dataDescription}\n\nSoru: ${question}\n\nBu veri seti için hangi Excel işlemlerini önerirsin?`;
        
        return await this.generateResponse(prompt, systemMsg);
    }

    // Excel veri analizi için streaming prompt
    async analyzeExcelDataStreaming(dataDescription, question, onChunk) {
        const systemMsg = "Sen bir Excel uzmanısın. Verilen veri seti üzerinde işlem yapmak için gerekli Excel komutlarını ve formüllerini önerirsin.";
        const prompt = `Veri: ${dataDescription}\n\nSoru: ${question}\n\nBu veri seti için hangi Excel işlemlerini önerirsin?`;
        
        return await this.generateStreamingResponse(prompt, systemMsg, onChunk);
    }

    // Excel formül önerileri
    async suggestFormulas(dataDescription, goal) {
        const systemMsg = "Sen bir Excel formül uzmanısın. Verilen hedef için uygun Excel formüllerini ve nasıl kullanılacağını açıklarsın.";
        const prompt = `Veri: ${dataDescription}\nHedef: ${goal}\n\nHangi Excel formüllerini önerirsin ve nasıl kullanılır?`;
        
        return await this.generateResponse(prompt, systemMsg);
    }

    // Grafik önerileri
    async suggestCharts(dataDescription) {
        const systemMsg = "Sen bir veri görselleştirme uzmanısın. Veri seti için en uygun grafik türlerini önerirsin.";
        const prompt = `Veri: ${dataDescription}\n\nBu veri seti için hangi grafik türlerini önerirsin? Neden uygun olduğunu açıkla.`;
        
        return await this.generateResponse(prompt, systemMsg);
    }

    // Excel işlem komutlarını çözümle
    async parseExcelCommand(command) {
        const systemMsg = "Sen bir Excel komut çözümleyicisisin. Verilen doğal dil komutunu Excel işlemlerine çevirirsin.";
        const prompt = `Komut: "${command}"\n\nBu komutu Excel'de nasıl gerçekleştirirsin? Adım adım açıkla.`;
        
        return await this.generateResponse(prompt, systemMsg);
    }

    // Excel işlem komutlarını streaming olarak çözümle
    async parseExcelCommandStreaming(command, onChunk) {
        const systemMsg = "Sen bir Excel komut çözümleyicisisin. Verilen doğal dil komutunu Excel işlemlerine çevirirsin.";
        const prompt = `Komut: "${command}"\n\nBu komutu Excel'de nasıl gerçekleştirirsin? Adım adım açıkla.`;
        
        return await this.generateStreamingResponse(prompt, systemMsg, onChunk);
    }
}

// Global AI client instance
window.aiClient = new LMStudioClient();
